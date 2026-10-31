#!/usr/bin/env python3
"""
Pipeline de ventas: ingesta -> limpieza y transformación -> reporte ejecutivo.

Recibe un export de ventas con los errores típicos de un ERP (fechas en varios
formatos, importes con símbolos, nombres mal escritos, duplicados) y genera:
  - un reporte Excel para gerencia (KPIs, gráficos, rankings y auditoría)
  - un CSV limpio para otros procesos

Criterio central: ninguna fila desaparece en silencio. Lo que se descarta queda
en la hoja "Excluidos" con su motivo y su texto original, y al final se
reconcilian las cifras (leídos = válidos + excluidos).

Uso:
    python generar_datos_sucios.py
    python pipeline.py
    python pipeline.py --input otro.csv --max-exclusion 0.25 --abrir -v
"""
from __future__ import annotations

import argparse
import contextlib
import csv
import difflib
import io
import logging
import os
import re
import subprocess
import sys
import time
import unicodedata
from collections.abc import Iterable
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import xlsxwriter
from xlsxwriter.exceptions import XlsxWriterException

# ---------------------------------------------------------------------------
# Rutas y reglas de negocio. Se ajustan aquí, sin tocar la lógica de las fases.
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
ENTRADA_POR_DEFECTO = BASE_DIR / "data" / "raw" / "ventas_sucias.csv"
SALIDA_POR_DEFECTO = BASE_DIR / "output"
LOG_POR_DEFECTO = BASE_DIR / "logs" / "pipeline.log"

CATALOGO = {  # producto -> categoría
    "Laptop Pro 14": "Cómputo",
    "Docking USB-C": "Cómputo",
    "Monitor 27 4K": "Pantallas",
    "Teclado Mecánico": "Periféricos",
    "Mouse Inalámbrico": "Periféricos",
    "Webcam HD 1080p": "Periféricos",
    "Auriculares ANC": "Periféricos",
    "Silla Ergonómica": "Mobiliario",
}
REGIONES = ["Norte", "Sur", "Centro", "Oriente"]
ESTADOS = ["Pagado", "Pendiente", "Anulado"]
ALIAS_ESTADO = {"pend": "Pendiente", "pdte": "Pendiente", "cancelado": "Anulado", "cancelada": "Anulado"}

COLUMNAS = ["id_pedido", "fecha", "cliente", "region", "producto",
            "cantidad", "precio_unitario", "descuento", "vendedor", "estado"]
ALIAS_COLUMNAS = {  # nombres alternativos frecuentes en exports de ERP
    "id": "id_pedido", "pedido": "id_pedido", "nro_pedido": "id_pedido",
    "cant": "cantidad", "qty": "cantidad",
    "precio": "precio_unitario", "precio_unit": "precio_unitario",
    "dcto": "descuento", "desc": "descuento",
    "sucursal": "region", "zona": "region",
    "status": "estado",
}

FORMATOS_FECHA = ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y")
FECHA_MINIMA = pd.Timestamp("2020-01-01")
MAX_CANTIDAD = 1_000
MAX_PRECIO = 50_000.0
MAX_DESCUENTO = 0.50
UMBRAL_SIMILITUD = 0.75         # 0 a 1: cuánto debe parecerse un texto para corregirlo
FACTOR_ATIPICO = 3.0            # atípica: supera Q3 + 3 x rango intercuartílico del producto...
MULTIPLO_MEDIANA = 5            # ...y además es al menos 5 veces la mediana de ese producto
DUPLICADOS_CONSERVAR = "first"  # "first" o "last": qué aparición de un ID se queda
MAX_EXCLUSION_POR_DEFECTO = 0.30
SIN_ASIGNAR = "Sin asignar"
MOTIVO_ESTRUCTURA = "Estructura inválida (nº de columnas)"

log = logging.getLogger("pipeline")


# ---------------------------------------------------------------------------
# Errores y estructuras de datos
# ---------------------------------------------------------------------------
class PipelineError(Exception):
    """Error esperado: se informa en una línea y el proceso termina con código 1."""


class IngestaError(PipelineError):
    """El archivo no existe, está vacío o no tiene el esquema esperado."""


class TransformacionError(PipelineError):
    """Una verificación interna de la limpieza no cuadra."""


class CalidadError(PipelineError):
    """Los datos están demasiado dañados para producir un reporte confiable."""


class ExportacionError(PipelineError):
    """No se pudo escribir el reporte o el CSV."""


@dataclass
class Ingesta:
    datos: pd.DataFrame           # columnas canónicas, valores sin tocar, índice = nº de línea
    originales: pd.Series         # texto de cada registro, para la hoja de auditoría
    codificacion: str
    separador: str
    lineas_vacias: int
    malformadas: dict[int, str]   # nº de línea -> contenido
    encabezados_corregidos: int


@dataclass
class Metricas:
    archivo: str
    codificacion: str
    separador: str
    leidos: int = 0
    validos: int = 0
    lineas_vacias: int = 0
    exclusiones: dict[str, int] = field(default_factory=dict)
    correcciones: dict[str, int] = field(default_factory=dict)

    @property
    def excluidos(self) -> int:
        return sum(self.exclusiones.values())

    @property
    def tasa_exclusion(self) -> float:
        return self.excluidos / self.leidos if self.leidos else 0.0


@dataclass
class ResultadoPipeline:
    ruta_reporte: Path
    ruta_csv: Path
    metricas: Metricas
    duracion_s: float


# ---------------------------------------------------------------------------
# Utilidades de texto
# ---------------------------------------------------------------------------
def _clave(texto: object) -> str:
    """Forma comparable de un texto: minúsculas, sin acentos, sin puntos y con espacios simples."""
    sin_acentos = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", " ", sin_acentos.lower().replace(".", "")).strip()


# ---------------------------------------------------------------------------
# Fase 1: ingesta
# ---------------------------------------------------------------------------
CODIFICACIONES = {"utf-8": "utf-8-sig", "cp1252": "cp1252"}  # nombre mostrado -> códec ("-sig" ignora el BOM)
SEPARADORES = (";", ",", "\t", "|")


def _leer_texto(ruta: Path) -> tuple[str, str]:
    """Decodifica el archivo probando codificaciones comunes; latin-1 acepta cualquier byte."""
    datos = ruta.read_bytes()
    for nombre, codec in CODIFICACIONES.items():
        try:
            return datos.decode(codec), nombre
        except UnicodeDecodeError:
            log.debug("%s no se pudo leer como %s, pruebo otra codificación", ruta.name, nombre)
    return datos.decode("latin-1"), "latin-1"


def _detectar_separador(texto: str) -> str:
    """Elige el separador que más veces aparece en el encabezado."""
    encabezado = next((linea for linea in texto.splitlines() if linea.strip()), "")
    separador = max(SEPARADORES, key=encabezado.count)
    if encabezado.count(separador) == 0:
        raise IngestaError("No se detectó el separador de columnas en la primera línea del archivo")
    return separador


def _leer_registros(texto: str, separador: str) -> tuple[list[str] | None, dict[int, list[str]], dict[int, str], int]:
    """Recorre el CSV registro a registro y anota el número de línea de cada uno."""
    lector = csv.reader(io.StringIO(texto), delimiter=separador)
    encabezado: list[str] | None = None
    registros: dict[int, list[str]] = {}
    malformadas: dict[int, str] = {}
    vacias = 0
    for campos in lector:
        if not any(c.strip() for c in campos):
            vacias += 1
        elif encabezado is None:
            encabezado = campos
        elif len(campos) == len(encabezado):
            registros[lector.line_num] = campos
        else:
            malformadas[lector.line_num] = separador.join(campos)
    return encabezado, registros, malformadas, vacias


def _normalizar_encabezado(encabezado: list[str]) -> tuple[list[str], int]:
    """'Precio Unit.' -> 'precio_unitario'. Valida que estén todas las columnas obligatorias."""
    claves = [_clave(c).replace(" ", "_") for c in encabezado]
    nombres = [ALIAS_COLUMNAS.get(k, k) for k in claves]
    repetidos = sorted({n for n in nombres if n and nombres.count(n) > 1})
    if repetidos:
        raise IngestaError(f"Encabezados repetidos: {', '.join(repetidos)}")
    faltantes = [c for c in COLUMNAS if c not in nombres]
    if faltantes:
        raise IngestaError(f"Faltan columnas obligatorias: {', '.join(faltantes)}")
    corregidos = sum(1 for original, nuevo in zip(encabezado, nombres) if original != nuevo)
    return nombres, corregidos


def ingestar(ruta: Path) -> Ingesta:
    """Fase 1. Lee el archivo tal como viene: todo texto, sin interpretar valores."""
    if not ruta.is_file():
        raise IngestaError(f"No existe el archivo de entrada: {ruta}\n"
                           "Para crear datos de prueba ejecuta: python generar_datos_sucios.py")
    if ruta.stat().st_size == 0:
        raise IngestaError(f"El archivo está vacío: {ruta}")

    try:
        texto, codificacion = _leer_texto(ruta)
        separador = _detectar_separador(texto)
        encabezado, registros, malformadas, vacias = _leer_registros(texto, separador)
    except (OSError, csv.Error) as exc:
        raise IngestaError(f"No se pudo leer {ruta.name}: {exc}") from exc

    if encabezado is None or not registros:
        raise IngestaError("El archivo no contiene filas de datos después del encabezado")

    nombres, corregidos = _normalizar_encabezado(encabezado)
    lineas = pd.Index(list(registros), name="linea")
    crudo = pd.DataFrame(list(registros.values()), columns=nombres, index=lineas)
    originales = pd.Series([separador.join(c) for c in registros.values()], index=lineas)
    return Ingesta(crudo[COLUMNAS], originales, codificacion, separador, vacias, malformadas, corregidos)


# ---------------------------------------------------------------------------
# Fase 2: limpieza y transformación
# ---------------------------------------------------------------------------
_RE_ID = re.compile(r"^PED[\s_-]*(\d{1,7})$", re.IGNORECASE)


def normalizar_id(valor: str) -> str | None:
    """'ped-12', ' PED00012 ' y 'PED_12' pasan a 'PED-00012'. None si no parece un ID."""
    coincidencia = _RE_ID.match(valor.strip())
    return f"PED-{int(coincidencia.group(1)):05d}" if coincidencia else None


def parsear_fechas(serie: pd.Series) -> pd.Series:
    """Prueba los formatos conocidos uno por uno; lo que ninguno reconoce queda como NaT."""
    resultado = pd.to_datetime(serie, format=FORMATOS_FECHA[0], errors="coerce")
    for formato in FORMATOS_FECHA[1:]:
        resultado = resultado.fillna(pd.to_datetime(serie, format=formato, errors="coerce"))
    return resultado


def parsear_monto(valor: str) -> float:
    """Convierte '$1,250.50', '1.250,50', 'USD 980' o '78,5' a float. NaN si no se puede leer."""
    texto = re.sub(r"[^\d.,-]", "", valor)
    if not re.search(r"\d", texto):
        return np.nan
    ultimo_punto, ultima_coma = texto.rfind("."), texto.rfind(",")
    if ultimo_punto >= 0 and ultima_coma >= 0:
        # Con los dos separadores presentes, el último es el decimal
        decimal = "." if ultimo_punto > ultima_coma else ","
        miles = "," if decimal == "." else "."
        texto = texto.replace(miles, "").replace(decimal, ".")
    elif ultimo_punto >= 0 or ultima_coma >= 0:
        separador = "." if ultimo_punto >= 0 else ","
        grupos = texto.split(separador)
        # '1.250' o '1,250,000' son miles; '78,5' o '24.90' son decimales
        es_miles = len(grupos) > 2 or len(grupos[-1]) == 3
        texto = texto.replace(separador, "" if es_miles else ".")
    try:
        return float(texto)
    except ValueError:
        return np.nan


def parsear_descuento(valor: str) -> float:
    """'10%', '10', '0.1' y '0,10' dan 0.10. Vacío significa sin descuento; ilegible da NaN."""
    if not valor:
        return 0.0
    try:
        numero = float(re.sub(r"[^\d.,]", "", valor).replace(",", "."))
    except ValueError:
        return np.nan
    return numero / 100 if "%" in valor or numero >= 1 else numero


def estandarizar_categoria(serie: pd.Series, validos: Iterable[str], alias: dict[str, str] | None = None) -> pd.Series:
    """Asigna cada texto al valor válido más parecido. Sin coincidencia confiable queda en None."""
    por_clave = {_clave(v): v for v in validos}
    equivalencias = {_clave(k): v for k, v in (alias or {}).items()}
    cache: dict[str, str | None] = {}

    def resolver(texto: str) -> str | None:
        clave = _clave(texto)
        if clave not in cache:
            if clave in por_clave:
                cache[clave] = por_clave[clave]
            elif clave in equivalencias:
                cache[clave] = equivalencias[clave]
            else:
                cercanos = difflib.get_close_matches(clave, list(por_clave), n=1, cutoff=UMBRAL_SIMILITUD)
                cache[clave] = por_clave[cercanos[0]] if cercanos else None
        return cache[clave]

    return serie.map(resolver)


def unificar_nombres(serie: pd.Series) -> pd.Series:
    """Une variantes del mismo nombre (mayúsculas, acentos, puntuación) en su grafía más frecuente."""
    claves = serie.map(_clave)

    def mejor_grafia(grafias: pd.Series) -> str:
        frecuencia = grafias.value_counts()
        # Desempate: se prefiere la grafía con mayúsculas y minúsculas mezcladas
        return max(frecuencia.index, key=lambda g: (frecuencia[g], not (g.isupper() or g.islower())))

    return claves.map(serie.groupby(claves).agg(mejor_grafia))


def limpiar_espacios(df: pd.DataFrame) -> pd.DataFrame:
    """Quita espacios sobrantes (al inicio, al final y repetidos) en todas las celdas."""
    return df.apply(lambda col: col.str.replace(r"\s+", " ", regex=True).str.strip())


def tipar_y_estandarizar(texto: pd.DataFrame) -> pd.DataFrame:
    """Lleva cada columna a su tipo final y unifica categorías. No elimina filas."""
    def con_default(col: str) -> pd.Series:
        return texto[col].where(texto[col] != "", SIN_ASIGNAR)

    return pd.DataFrame({
        "id_pedido": texto["id_pedido"].map(normalizar_id),
        "fecha": parsear_fechas(texto["fecha"]),
        "cliente": unificar_nombres(con_default("cliente")),
        "region": estandarizar_categoria(texto["region"], REGIONES).fillna(SIN_ASIGNAR),
        "vendedor": unificar_nombres(con_default("vendedor")),
        "producto": estandarizar_categoria(texto["producto"], CATALOGO),
        "cantidad": texto["cantidad"].map(parsear_monto),
        "precio_unitario": texto["precio_unitario"].map(parsear_monto),
        "descuento": texto["descuento"].map(parsear_descuento),
        "estado": estandarizar_categoria(texto["estado"], ESTADOS, ALIAS_ESTADO),
    })


def validar(t: pd.DataFrame, hoy: pd.Timestamp) -> pd.Series:
    """Devuelve el motivo de exclusión de cada fila ('' si pasa todas las reglas)."""
    reglas = pd.DataFrame({
        "ID inválido": t["id_pedido"].isna(),
        "Fecha inválida": t["fecha"].isna(),
        "Fecha fuera de rango": (t["fecha"] < FECHA_MINIMA) | (t["fecha"] > hoy),
        "Cantidad inválida": ~(t["cantidad"].between(1, MAX_CANTIDAD) & (t["cantidad"] % 1 == 0)),
        "Precio inválido": ~(t["precio_unitario"] > 0) | (t["precio_unitario"] > MAX_PRECIO),
        "Descuento inválido": ~t["descuento"].between(0, MAX_DESCUENTO),
        "Producto no reconocido": t["producto"].isna(),
        "Estado inválido": t["estado"].isna(),
    })
    # Una fila puede romper varias reglas: se listan todas, en orden de prioridad
    return reglas.apply(lambda fila: "; ".join(fila.index[fila]), axis=1)


def marcar_duplicados(t: pd.DataFrame, motivo: pd.Series) -> pd.Series:
    """Entre las filas válidas, conserva una sola aparición de cada ID."""
    validas = t[motivo == ""]
    repetidas = validas.index[validas.duplicated(subset="id_pedido", keep=DUPLICADOS_CONSERVAR)]
    return motivo.mask(motivo.index.isin(repetidas), "ID duplicado")


def marcar_anulados(t: pd.DataFrame, motivo: pd.Series) -> pd.Series:
    """Los pedidos anulados son datos correctos, pero no deben sumar al ingreso."""
    return motivo.mask((motivo == "") & (t["estado"] == "Anulado"), "Pedido anulado (regla de negocio)")


def apartar_atipicos(t: pd.DataFrame, motivo: pd.Series) -> pd.Series:
    """Aparta para revisión las cantidades desproporcionadas para su producto (suelen ser ceros de más).

    Una cantidad es atípica si supera Q3 + k x IQR dentro de su producto y además es varias veces
    la mediana. La segunda condición evita falsos positivos cuando casi todos los pedidos son iguales.
    """
    candidatas = t[motivo == ""]
    if candidatas.empty:
        return motivo
    q = candidatas.groupby("producto")["cantidad"].quantile([0.25, 0.5, 0.75]).unstack()
    limite = np.maximum(q[0.75] + FACTOR_ATIPICO * (q[0.75] - q[0.25]), MULTIPLO_MEDIANA * q[0.5])
    atipicas = candidatas.index[candidatas["cantidad"] > candidatas["producto"].map(limite)]
    return motivo.mask(motivo.index.isin(atipicas), "Cantidad atípica (revisar en origen)")


def enriquecer(validos: pd.DataFrame) -> pd.DataFrame:
    """Agrega categoría, importes calculados y mes."""
    df = validos.copy()
    df["cantidad"] = df["cantidad"].astype(int)
    df["categoria"] = df["producto"].map(CATALOGO)
    df["total_bruto"] = (df["cantidad"] * df["precio_unitario"]).round(2)
    df["descuento_monto"] = (df["total_bruto"] * df["descuento"]).round(2)
    df["ingreso_neto"] = (df["total_bruto"] - df["descuento_monto"]).round(2)
    df["mes"] = df["fecha"].dt.strftime("%Y-%m")
    orden = ["id_pedido", "fecha", "mes", "cliente", "region", "vendedor", "producto", "categoria",
             "cantidad", "precio_unitario", "descuento", "total_bruto", "descuento_monto",
             "ingreso_neto", "estado"]
    limpio = df[orden].sort_values(["fecha", "id_pedido"]).rename_axis("linea_origen").reset_index()
    return limpio[[*orden, "linea_origen"]]


def construir_excluidos(ingesta: Ingesta, motivo: pd.Series) -> pd.DataFrame:
    """Tabla de auditoría: línea, motivo y texto original de cada registro descartado."""
    excluidas = motivo != ""
    partes = [pd.DataFrame({"motivo": motivo[excluidas], "registro_original": ingesta.originales[excluidas]})]
    if ingesta.malformadas:
        partes.append(pd.DataFrame(
            {"motivo": MOTIVO_ESTRUCTURA, "registro_original": list(ingesta.malformadas.values())},
            index=pd.Index(list(ingesta.malformadas), name="linea")))
    con_datos = [p for p in partes if not p.empty]
    if not con_datos:
        return pd.DataFrame(columns=["linea", "motivo", "registro_original"])
    return pd.concat(con_datos).sort_index().rename_axis("linea").reset_index()


def contar_exclusiones(motivo: pd.Series, malformadas: int) -> dict[str, int]:
    """Cuenta por motivo principal (el primero de la lista) para que la suma cuadre."""
    principales = motivo[motivo != ""].str.split("; ").str[0]
    conteo = principales.value_counts().to_dict()
    if malformadas:
        conteo[MOTIVO_ESTRUCTURA] = malformadas
    return dict(sorted(conteo.items(), key=lambda par: -par[1]))


def contar_correcciones(texto: pd.DataFrame, tipado: pd.DataFrame, validas: pd.Series,
                        encabezados: int) -> dict[str, int]:
    """Cuántos registros válidos tuvieron algún valor corregido, por tipo de corrección."""
    t, f = texto[validas], tipado[validas]

    def cambiados(col: str) -> int:
        return int(((t[col] != "") & (t[col] != f[col])).sum())

    return {
        "Encabezados normalizados": encabezados,
        "IDs de pedido normalizados": cambiados("id_pedido"),
        "Fechas en formato no ISO convertidas": int((~t["fecha"].str.fullmatch(r"\d{4}-\d{2}-\d{2}")).sum()),
        "Importes con símbolo o separador atípico": int((~t["precio_unitario"].str.fullmatch(r"\d+(\.\d+)?")).sum()),
        "Nombres de cliente unificados": cambiados("cliente"),
        "Nombres de vendedor unificados": cambiados("vendedor"),
        "Regiones corregidas": cambiados("region"),
        "Productos corregidos": cambiados("producto"),
        "Estados estandarizados": cambiados("estado"),
        "Vacíos imputados como 'Sin asignar'": int((t[["cliente", "region", "vendedor"]] == "").sum().sum()),
    }


def transformar(ingesta: Ingesta, archivo: str,
                hoy: pd.Timestamp | None = None) -> tuple[pd.DataFrame, pd.DataFrame, Metricas]:
    """Fase 2. Devuelve (registros limpios, registros excluidos con motivo, métricas)."""
    if hoy is None:
        hoy = pd.Timestamp.today().normalize()

    texto = limpiar_espacios(ingesta.datos)
    tipado = tipar_y_estandarizar(texto)
    motivo = validar(tipado, hoy)
    motivo = marcar_duplicados(tipado, motivo)
    motivo = marcar_anulados(tipado, motivo)
    motivo = apartar_atipicos(tipado, motivo)
    validas = motivo == ""
    if not validas.any():
        raise CalidadError("Ninguna fila superó las validaciones. Revisa que el archivo sea el correcto.")

    limpios = enriquecer(tipado[validas])
    metricas = Metricas(
        archivo=archivo, codificacion=ingesta.codificacion, separador=ingesta.separador,
        leidos=len(texto) + len(ingesta.malformadas), validos=len(limpios),
        lineas_vacias=ingesta.lineas_vacias,
        exclusiones=contar_exclusiones(motivo, len(ingesta.malformadas)),
        correcciones=contar_correcciones(texto, tipado, validas, ingesta.encabezados_corregidos),
    )
    # Reconciliación: cada registro leído termina en un único destino
    if metricas.validos + metricas.excluidos != metricas.leidos:
        raise TransformacionError(
            f"Las cifras no cuadran: {metricas.leidos} leídos, {metricas.validos} válidos "
            f"y {metricas.excluidos} excluidos")
    return limpios, construir_excluidos(ingesta, motivo), metricas


# ---------------------------------------------------------------------------
# Fase 3: exportación
# ---------------------------------------------------------------------------
FUENTE = "Arial"
AZUL, AZUL_CLARO, GRIS = "#1F3864", "#DCE6F1", "#595959"
FORMATOS_NUMERICOS = {"usd": "$#,##0.00", "usd0": "$#,##0", "int": "#,##0", "pct": "0.0%", "fecha": "dd/mm/yyyy"}
FILAS_DE_GRAFICO = 15  # alto del gráfico (300 px) medido en filas de 20 px

ETIQUETAS_DATOS = {
    "id_pedido": "ID pedido", "fecha": "Fecha", "mes": "Mes", "cliente": "Cliente", "region": "Región",
    "vendedor": "Vendedor", "producto": "Producto", "categoria": "Categoría", "cantidad": "Cantidad",
    "precio_unitario": "Precio unitario", "descuento": "Descuento", "total_bruto": "Total bruto",
    "descuento_monto": "Monto descuento", "ingreso_neto": "Ingreso neto", "estado": "Estado",
    "linea_origen": "Línea origen",
}
FORMATO_POR_COLUMNA = {  # etiqueta de columna -> formato de número
    "Fecha": "fecha", "Cantidad": "int", "Precio unitario": "usd", "Descuento": "pct",
    "Total bruto": "usd", "Monto descuento": "usd", "Ingreso neto": "usd", "Pedidos": "int",
    "Unidades": "int", "Ticket promedio": "usd", "% del total": "pct", "Registros": "int",
    "% de lo leído": "pct", "Línea": "int", "Línea origen": "int",
}


def _py(valor):
    """Convierte tipos de numpy y pandas a tipos nativos que xlsxwriter sabe escribir."""
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
        return None
    if isinstance(valor, pd.Timestamp):
        return valor.to_pydatetime()
    if isinstance(valor, np.generic):
        return valor.item()
    return valor


def _crear_formatos(libro: xlsxwriter.Workbook) -> dict:
    def fmt(**props):
        return libro.add_format({"font_name": FUENTE, "font_size": 10, "valign": "vcenter", **props})

    formatos = {
        "texto": fmt(),
        "texto_der": fmt(align="right"),
        "titulo": fmt(font_size=18, bold=True, font_color=AZUL),
        "subtitulo": fmt(font_color=GRIS),
        "seccion": fmt(font_size=12, bold=True, font_color=AZUL, bottom=2, bottom_color=AZUL),
        "encabezado": fmt(bold=True, font_color="#FFFFFF", bg_color=AZUL, align="center",
                          text_wrap=True, border=1, border_color="#FFFFFF"),
        "kpi_etiqueta": fmt(font_size=9, font_color=GRIS, bg_color=AZUL_CLARO, align="center",
                            top=1, left=1, right=1, border_color="#BFBFBF"),
    }
    for nombre, numero in FORMATOS_NUMERICOS.items():
        formatos[nombre] = fmt(num_format=numero)
        formatos[f"kpi_{nombre}"] = fmt(font_size=16, bold=True, font_color=AZUL, bg_color=AZUL_CLARO,
                                        align="center", num_format=numero, bottom=1, left=1, right=1,
                                        border_color="#BFBFBF")
    return formatos


def _escribir_tabla(hoja, fila: int, col: int, df: pd.DataFrame, f: dict) -> int:
    """Escribe un DataFrame con encabezado desde (fila, col). Devuelve la última fila usada."""
    for j, nombre in enumerate(df.columns):
        hoja.write(fila, col + j, nombre, f["encabezado"])
    for i, registro in enumerate(df.itertuples(index=False, name=None), start=1):
        for j, valor in enumerate(registro):
            formato = f[FORMATO_POR_COLUMNA.get(df.columns[j], "texto")]
            hoja.write(fila + i, col + j, _py(valor), formato)
    return fila + len(df)


def _ajustar_anchos(hoja, df: pd.DataFrame, tope: int = 45) -> None:
    for j, nombre in enumerate(df.columns):
        mayor = int(df[nombre].astype(str).str.len().max()) if len(df) else 0
        hoja.set_column(j, j, min(tope, max(len(nombre), mayor) + 3))


def calcular_kpis(df: pd.DataFrame) -> dict[str, float]:
    ingreso = float(df["ingreso_neto"].sum())
    bruto = float(df["total_bruto"].sum())
    por_cobrar = float(df.loc[df["estado"] == "Pendiente", "ingreso_neto"].sum())
    return {
        "ingreso": ingreso,
        "pedidos": len(df),
        "ticket": ingreso / len(df),
        "unidades": int(df["cantidad"].sum()),
        "descuento": float(df["descuento_monto"].sum()) / bruto if bruto else 0.0,
        "por_cobrar": por_cobrar / ingreso if ingreso else 0.0,
    }


def agrupar(df: pd.DataFrame, por: str, etiqueta: str, por_ingreso: bool = True) -> pd.DataFrame:
    """Ingreso, pedidos, unidades, ticket promedio y participación según la dimensión indicada."""
    t = df.groupby(por, as_index=False).agg(
        ingreso=("ingreso_neto", "sum"), pedidos=("id_pedido", "count"), unidades=("cantidad", "sum"))
    t["ticket"] = t["ingreso"] / t["pedidos"]
    t["participacion"] = t["ingreso"] / t["ingreso"].sum()
    t = t.sort_values("ingreso", ascending=False) if por_ingreso else t.sort_values(por)
    return t.rename(columns={por: etiqueta, "ingreso": "Ingreso neto", "pedidos": "Pedidos",
                             "unidades": "Unidades", "ticket": "Ticket promedio",
                             "participacion": "% del total"}).reset_index(drop=True)


def _bloque_con_grafico(libro, hoja, f, fila: int, titulo: str, tabla: pd.DataFrame, tipo: str) -> int:
    """Sección del resumen: tabla a la izquierda y gráfico a la derecha. Devuelve la última fila usada."""
    hoja.merge_range(fila, 1, fila, 12, titulo, f["seccion"])
    ultima = _escribir_tabla(hoja, fila + 1, 1, tabla, f)

    serie = {"name": "Ingreso neto",
             "categories": ["Resumen", fila + 2, 1, ultima, 1],
             "values": ["Resumen", fila + 2, 2, ultima, 2]}
    if tipo == "line":
        serie.update({"line": {"color": AZUL, "width": 2.5},
                      "marker": {"type": "circle", "size": 6, "fill": {"color": AZUL}, "border": {"color": AZUL}}})
    else:
        serie.update({"fill": {"color": AZUL}, "gap": 80,
                      "data_labels": {"value": True, "num_format": "$#,##0", "font": {"name": FUENTE, "size": 9}}})
    grafico = libro.add_chart({"type": tipo})
    grafico.add_series(serie)
    grafico.set_title({"none": True})
    grafico.set_legend({"none": True})
    grafico.set_x_axis({"num_font": {"name": FUENTE, "size": 9}})
    grafico.set_y_axis({"num_format": "$#,##0", "num_font": {"name": FUENTE, "size": 9},
                        "major_gridlines": {"visible": True, "line": {"color": "#E0E0E0"}}})
    grafico.set_size({"width": 618, "height": 300})
    hoja.insert_chart(fila + 1, 7, grafico)
    return max(ultima, fila + FILAS_DE_GRAFICO)


def _hoja_resumen(libro, hoja, f: dict, limpios: pd.DataFrame, m: Metricas) -> None:
    kpis = calcular_kpis(limpios)
    desde, hasta = limpios["fecha"].min(), limpios["fecha"].max()
    hoja.hide_gridlines(2)
    hoja.set_column(0, 0, 2)
    hoja.set_column(1, 12, 14)
    hoja.set_row(0, 30)
    hoja.write(0, 1, "Reporte ejecutivo de ventas", f["titulo"])
    hoja.write(1, 1, f"Período {desde:%d/%m/%Y} al {hasta:%d/%m/%Y}   |   Fuente: {m.archivo}   |   "
                     f"Generado {datetime.now():%d/%m/%Y %H:%M}", f["subtitulo"])
    hoja.write(2, 1, f"Datos: {m.leidos:,} registros leídos, {m.validos:,} válidos "
                     f"({m.validos / m.leidos:.1%}) y {m.excluidos:,} excluidos con su motivo (hoja Excluidos)",
               f["subtitulo"])

    tarjetas = [("Ingreso neto", kpis["ingreso"], "kpi_usd0"), ("Pedidos", kpis["pedidos"], "kpi_int"),
                ("Ticket promedio", kpis["ticket"], "kpi_usd0"), ("Unidades vendidas", kpis["unidades"], "kpi_int"),
                ("Descuento promedio", kpis["descuento"], "kpi_pct"), ("Ingreso por cobrar", kpis["por_cobrar"], "kpi_pct")]
    for i, (etiqueta, valor, formato) in enumerate(tarjetas):
        col = 1 + 2 * i
        hoja.merge_range(4, col, 4, col + 1, etiqueta, f["kpi_etiqueta"])
        hoja.merge_range(5, col, 5, col + 1, valor, f[formato])
    hoja.set_row(5, 30)

    columnas = ["Ingreso neto", "Pedidos", "Ticket promedio", "% del total"]
    meses = agrupar(limpios, "mes", "Mes", por_ingreso=False)
    fila = _bloque_con_grafico(libro, hoja, f, 7, "Evolución mensual del ingreso neto",
                               meses[["Mes", *columnas]], "line") + 2
    regiones = agrupar(limpios, "region", "Región")
    _bloque_con_grafico(libro, hoja, f, fila, "Ingreso neto por región",
                        regiones[["Región", *columnas]], "column")
    categorias = agrupar(limpios, "categoria", "Categoría")
    inicio = fila + len(regiones) + 3
    hoja.merge_range(inicio, 1, inicio, 5, "Ingreso neto por categoría", f["seccion"])
    _escribir_tabla(hoja, inicio + 1, 1, categorias[["Categoría", *columnas]], f)
    hoja.set_landscape()
    hoja.fit_to_pages(1, 0)


def _hoja_rankings(hoja, f: dict, limpios: pd.DataFrame) -> None:
    hoja.hide_gridlines(2)
    hoja.set_column(0, 0, 2)
    hoja.set_column(1, 1, 30)
    hoja.set_column(2, 6, 15)
    bloques = [("Productos por ingreso neto", "producto", "Producto", None),
               ("Top 10 clientes", "cliente", "Cliente", 10),
               ("Vendedores", "vendedor", "Vendedor", None)]
    fila = 0
    for titulo, columna, etiqueta, limite in bloques:
        tabla = agrupar(limpios, columna, etiqueta)
        if limite:
            tabla = tabla.head(limite)
        hoja.merge_range(fila, 1, fila, 6, titulo, f["seccion"])
        ultima = _escribir_tabla(hoja, fila + 1, 1, tabla, f)
        hoja.conditional_format(fila + 2, 2, ultima, 2,
                                {"type": "data_bar", "bar_color": "#8EA9DB", "bar_solid": True})
        fila = ultima + 3
    hoja.set_landscape()
    hoja.fit_to_pages(1, 0)


def _hoja_datos(hoja, f: dict, limpios: pd.DataFrame) -> None:
    export = limpios.rename(columns=ETIQUETAS_DATOS)
    filas = [[_py(v) for v in fila] for fila in export.itertuples(index=False, name=None)]
    columnas = [{"header": c, "format": f[FORMATO_POR_COLUMNA.get(c, "texto")],
                 "header_format": f["encabezado"]} for c in export.columns]
    hoja.add_table(0, 0, len(export), len(export.columns) - 1,
                   {"data": filas, "columns": columnas, "style": "Table Style Medium 2"})
    hoja.freeze_panes(1, 0)
    _ajustar_anchos(hoja, export)


def _hoja_calidad(hoja, f: dict, m: Metricas) -> None:
    hoja.hide_gridlines(2)
    hoja.set_column(0, 0, 2)
    hoja.set_column(1, 1, 58)
    hoja.set_column(2, 5, 18)
    hoja.set_row(0, 30)
    hoja.write(0, 1, "Calidad de datos y auditoría del procesamiento", f["titulo"])

    hoja.merge_range(2, 1, 2, 2, "Resumen del procesamiento", f["seccion"])
    pares = [("Archivo de origen", m.archivo, "texto_der"),
             ("Codificación detectada", m.codificacion, "texto_der"),
             ("Separador detectado", repr(m.separador), "texto_der"),
             ("Registros leídos", m.leidos, "int"),
             ("Registros válidos", m.validos, "int"),
             ("Registros excluidos", m.excluidos, "int"),
             ("Tasa de aprovechamiento", m.validos / m.leidos, "pct"),
             ("Líneas vacías ignoradas", m.lineas_vacias, "int")]
    for i, (etiqueta, valor, formato) in enumerate(pares, start=3):
        hoja.write(i, 1, etiqueta, f["texto"])
        hoja.write(i, 2, valor, f[formato])

    fila = 3 + len(pares) + 1
    hoja.merge_range(fila, 1, fila, 3, "Exclusiones por motivo principal", f["seccion"])
    exclusiones = pd.DataFrame({"Motivo": list(m.exclusiones), "Registros": list(m.exclusiones.values())})
    exclusiones["% de lo leído"] = exclusiones["Registros"] / m.leidos
    fila = _escribir_tabla(hoja, fila + 1, 1, exclusiones, f) + 2

    hoja.merge_range(fila, 1, fila, 2, "Correcciones aplicadas a los registros válidos", f["seccion"])
    correcciones = pd.DataFrame({"Corrección": list(m.correcciones), "Registros": list(m.correcciones.values())})
    fila = _escribir_tabla(hoja, fila + 1, 1, correcciones, f) + 2

    hoja.merge_range(fila, 1, fila, 5, "Reglas aplicadas", f["seccion"])
    orden_duplicados = "primera" if DUPLICADOS_CONSERVAR == "first" else "última"
    reglas = [("Cantidad", f"entera, de 1 a {MAX_CANTIDAD:,}"),
              ("Precio unitario", f"mayor que 0 y hasta ${MAX_PRECIO:,.0f}"),
              ("Descuento", f"de 0% a {MAX_DESCUENTO:.0%}; vacío se toma como 0%"),
              ("Fecha", f"desde {FECHA_MINIMA:%d/%m/%Y} hasta el día de la ejecución"),
              ("Textos con error", f"se corrigen si se parecen al menos {UMBRAL_SIMILITUD:.0%} a un valor válido"),
              ("ID duplicado", f"se conserva la {orden_duplicados} aparición"),
              ("Pedido anulado", "se excluye del ingreso y queda registrado en la hoja Excluidos"),
              ("Cantidad atípica", f"supera Q3 + {FACTOR_ATIPICO:g} x IQR del producto y es {MULTIPLO_MEDIANA}x su mediana o más; se aparta para revisión")]
    for i, (regla, detalle) in enumerate(reglas, start=fila + 1):
        hoja.write(i, 1, regla, f["texto"])
        hoja.merge_range(i, 2, i, 5, detalle, f["texto"])
    hoja.set_landscape()
    hoja.fit_to_pages(1, 0)


def _hoja_excluidos(hoja, f: dict, excluidos: pd.DataFrame) -> None:
    hoja.set_column(0, 0, 10)
    hoja.set_column(1, 1, 52)
    hoja.set_column(2, 2, 110)
    if excluidos.empty:
        hoja.write(0, 0, "No hubo registros excluidos en esta ejecución.", f["texto"])
        return
    export = excluidos.rename(columns={"linea": "Línea", "motivo": "Motivo", "registro_original": "Registro original"})
    filas = [[_py(v) for v in fila] for fila in export.itertuples(index=False, name=None)]
    columnas = [{"header": c, "format": f[FORMATO_POR_COLUMNA.get(c, "texto")],
                 "header_format": f["encabezado"]} for c in export.columns]
    hoja.add_table(0, 0, len(export), len(export.columns) - 1,
                   {"data": filas, "columns": columnas, "style": "Table Style Medium 2"})
    hoja.freeze_panes(1, 0)


def _escribir_excel(ruta: Path, limpios: pd.DataFrame, excluidos: pd.DataFrame, m: Metricas) -> None:
    with xlsxwriter.Workbook(str(ruta), {"default_date_format": "dd/mm/yyyy"}) as libro:
        libro.set_properties({"title": "Reporte ejecutivo de ventas", "comments": "Generado por pipeline.py"})
        f = _crear_formatos(libro)
        resumen = libro.add_worksheet("Resumen")
        rankings = libro.add_worksheet("Rankings")
        datos = libro.add_worksheet("Datos limpios")
        calidad = libro.add_worksheet("Calidad de datos")
        auditoria = libro.add_worksheet("Excluidos")
        _hoja_resumen(libro, resumen, f, limpios, m)
        _hoja_rankings(rankings, f, limpios)
        _hoja_datos(datos, f, limpios)
        _hoja_calidad(calidad, f, m)
        _hoja_excluidos(auditoria, f, excluidos)


def exportar(limpios: pd.DataFrame, excluidos: pd.DataFrame, metricas: Metricas,
             salida: Path) -> tuple[Path, Path]:
    """Fase 3. Escribe el reporte Excel y el CSV limpio. Con marca de tiempo para no pisar corridas previas."""
    sello = datetime.now().strftime("%Y%m%d_%H%M%S")
    ruta_xlsx = salida / f"reporte_ejecutivo_{sello}.xlsx"
    ruta_csv = salida / f"ventas_limpias_{sello}.csv"
    try:
        salida.mkdir(parents=True, exist_ok=True)
        limpios.to_csv(ruta_csv, index=False, encoding="utf-8")
        _escribir_excel(ruta_xlsx, limpios, excluidos, metricas)
    except (OSError, XlsxWriterException) as exc:
        for parcial in (ruta_xlsx, ruta_csv):  # no dejar archivos a medias
            with contextlib.suppress(OSError):
                parcial.unlink(missing_ok=True)
        pista = " Si el reporte está abierto en Excel, ciérralo y vuelve a ejecutar." if "ermission" in str(exc) else ""
        raise ExportacionError(f"No se pudo escribir en {salida}: {exc}.{pista}") from exc
    return ruta_xlsx, ruta_csv


def abrir_archivo(ruta: Path) -> None:
    """Abre el archivo con la aplicación predeterminada del sistema (útil para la demo)."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(ruta)  # type: ignore[attr-defined]
        else:
            subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", str(ruta)], check=False)
    except OSError as exc:
        log.warning("No se pudo abrir %s automáticamente: %s", ruta.name, exc)


# ---------------------------------------------------------------------------
# Orquestación y línea de comandos
# ---------------------------------------------------------------------------
def _dato(etiqueta: str, valor: object, nivel: int = 1) -> None:
    """Línea de log con la etiqueta a la izquierda y el valor alineado."""
    izquierda = f"{'  ' * nivel}{etiqueta} ".ljust(50, ".")
    log.info("%s %s", izquierda, valor)


def _resumir(conteos: dict[str, int], maximo: int, nivel: int) -> None:
    """Muestra los mayores conteos y agrupa el resto en una línea. Con -v se ve todo."""
    ordenados = sorted(((k, v) for k, v in conteos.items() if v), key=lambda par: -par[1])
    visibles = ordenados if log.isEnabledFor(logging.DEBUG) else ordenados[:maximo]
    for nombre, cantidad in visibles:
        _dato(nombre, f"{cantidad:,}", nivel)
    resto = ordenados[len(visibles):]
    if resto:
        _dato(f"otros ({len(resto)} tipos)", f"{sum(c for _, c in resto):,}", nivel)


def _mostrar(ruta: Path) -> str:
    try:
        return str(ruta.relative_to(Path.cwd()))
    except ValueError:
        return str(ruta)


def ejecutar_pipeline(entrada: Path, salida: Path, max_exclusion: float) -> ResultadoPipeline:
    inicio = time.perf_counter()
    log.info("Pipeline de ventas | entrada: %s", entrada.name)

    log.info("[1/3] INGESTA")
    ingesta = ingestar(entrada)
    _dato("codificación / separador", f"{ingesta.codificacion} / {ingesta.separador!r}")
    _dato("registros leídos", f"{len(ingesta.datos) + len(ingesta.malformadas):,}")
    _dato("líneas vacías ignoradas", ingesta.lineas_vacias)
    _dato("líneas con estructura inválida", len(ingesta.malformadas))

    log.info("[2/3] TRANSFORMACIÓN")
    limpios, excluidos, metricas = transformar(ingesta, entrada.name)
    _resumir(metricas.correcciones, maximo=5, nivel=1)
    _dato("registros válidos", f"{metricas.validos:,} ({metricas.validos / metricas.leidos:.1%})")
    _dato("registros excluidos", f"{metricas.excluidos:,} ({metricas.tasa_exclusion:.1%})")
    _resumir(metricas.exclusiones, maximo=6, nivel=2)
    if metricas.tasa_exclusion > max_exclusion:
        raise CalidadError(
            f"Se excluyó el {metricas.tasa_exclusion:.1%} de los registros, por encima del máximo "
            f"permitido ({max_exclusion:.0%}). No se generó reporte: revisa el archivo de origen "
            "o sube el umbral con --max-exclusion.")

    log.info("[3/3] EXPORTACIÓN")
    reporte, ruta_csv = exportar(limpios, excluidos, metricas, salida)
    _dato("reporte Excel", _mostrar(reporte))
    _dato("CSV limpio", _mostrar(ruta_csv))

    duracion = time.perf_counter() - inicio
    log.info("LISTO en %.1f s | %s leídos, %s válidos (%.1f%%) | ingreso neto $%s",
             duracion, f"{metricas.leidos:,}", f"{metricas.validos:,}",
             100 * metricas.validos / metricas.leidos, f"{limpios['ingreso_neto'].sum():,.2f}")
    return ResultadoPipeline(reporte, ruta_csv, metricas, duracion)


def configurar_logging(verbose: bool, ruta_log: Path = LOG_POR_DEFECTO) -> None:
    """Consola con hora corta y archivo de log con fecha completa."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    consola = logging.StreamHandler(sys.stdout)
    consola.setFormatter(logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%H:%M:%S"))
    manejadores: list[logging.Handler] = [consola]
    try:
        ruta_log.parent.mkdir(parents=True, exist_ok=True)
        archivo = logging.FileHandler(ruta_log, encoding="utf-8")
        archivo.setFormatter(logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s"))
        manejadores.append(archivo)
    except OSError as exc:
        print(f"Aviso: no se pudo abrir el archivo de log ({exc}). Se usa solo la consola.", file=sys.stderr)
    logging.basicConfig(level=logging.DEBUG if verbose else logging.INFO, handlers=manejadores, force=True)


def leer_argumentos(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Pipeline de ventas: ingesta, limpieza y reporte ejecutivo.")
    p.add_argument("--input", "-i", type=Path, default=ENTRADA_POR_DEFECTO,
                   help="CSV de entrada (por defecto: data/raw/ventas_sucias.csv)")
    p.add_argument("--output-dir", "-o", type=Path, default=SALIDA_POR_DEFECTO,
                   help="carpeta de salida (por defecto: output/)")
    p.add_argument("--max-exclusion", type=float, default=MAX_EXCLUSION_POR_DEFECTO,
                   help="fracción máxima de registros excluidos antes de abortar, de 0 a 1 (por defecto: 0.30)")
    p.add_argument("--abrir", action="store_true", help="abre el reporte al terminar")
    p.add_argument("--verbose", "-v", action="store_true", help="muestra mensajes de depuración")
    args = p.parse_args(argv)
    if not 0 <= args.max_exclusion <= 1:
        p.error("--max-exclusion debe estar entre 0 y 1")
    return args


def main(argv: list[str] | None = None) -> int:
    args = leer_argumentos(argv)
    configurar_logging(args.verbose)
    try:
        resultado = ejecutar_pipeline(args.input, args.output_dir, args.max_exclusion)
    except PipelineError as exc:
        log.error("%s", exc)
        return 1
    except KeyboardInterrupt:
        log.warning("Interrumpido por el usuario")
        return 130
    except Exception:  # error no previsto: se registra el traceback completo
        log.exception("Error inesperado")
        return 2
    if args.abrir:
        abrir_archivo(resultado.ruta_reporte)
    return 0


if __name__ == "__main__":
    sys.exit(main())
