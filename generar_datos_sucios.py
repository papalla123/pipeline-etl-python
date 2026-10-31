#!/usr/bin/env python3
"""
Genera ventas_sucias.csv: un export de ventas con los errores típicos de un ERP.

Los datos son sintéticos y reproducibles (misma semilla, mismo archivo). Las
primeras filas son casos fijos que muestran cada tipo de error; el resto se
genera al azar con una proporción realista de suciedad.

Uso:
    python generar_datos_sucios.py
    python generar_datos_sucios.py --filas 2000 --semilla 7 --encoding latin-1
"""
from __future__ import annotations

import argparse
import csv
import io
import random
import unicodedata
from datetime import date, timedelta
from pathlib import Path

DESTINO = Path(__file__).resolve().parent / "data" / "raw" / "ventas_sucias.csv"
ENCABEZADO = ["ID Pedido", "Fecha ", "Cliente", "Región", "Producto", "Cant.",
              "Precio Unit.", "Descuento", "Vendedor", "Estado"]

# producto -> (precio mínimo, precio máximo, cantidad máxima por pedido)
PRODUCTOS = {
    "Laptop Pro 14": (950, 1350, 8),
    "Docking USB-C": (90, 160, 20),
    "Monitor 27 4K": (280, 420, 15),
    "Teclado Mecánico": (60, 110, 30),
    "Mouse Inalámbrico": (20, 45, 40),
    "Webcam HD 1080p": (35, 70, 25),
    "Auriculares ANC": (70, 180, 20),
    "Silla Ergonómica": (150, 320, 10),
}
# cliente -> (región, vendedor asignado)
CLIENTES = {
    "Nova Distribuciones": ("Norte", "Lucía Paredes"),
    "Grupo Aurora": ("Norte", "Lucía Paredes"),
    "Logística Delta": ("Sur", "Marcos Quiroz"),
    "Cobalto Industrial": ("Centro", "Valeria Soto"),
    "Alfa Retail": ("Oriente", "Diego Rivas"),
    "Tecnoandina": ("Sur", "Marcos Quiroz"),
    "Corporación Zenit": ("Centro", "Andrés Molina"),
    "Mercantil Horizonte": ("Oriente", "Camila Herrera"),
    "Soluciones Orbital": ("Sur", "Diego Rivas"),
    "Importadora Brisa": ("Norte", "Camila Herrera"),
    "Textiles Pampa": ("Centro", "Valeria Soto"),
    "Comercial Andes Sur": ("Sur", "Marcos Quiroz"),
    "Inversiones Titán": ("Centro", "Andrés Molina"),
    "Servicios Meridiano": ("Oriente", "Diego Rivas"),
    "Constructora Roble": ("Norte", "Lucía Paredes"),
    "Agroexport Valle": ("Oriente", "Camila Herrera"),
}

# Casos fijos: cada fila ilustra un problema distinto (ver comentario a la derecha).
CASOS_FIJOS = [
    ["PED-00001", "2026-01-05", "Nova Distribuciones", "Norte", "Laptop Pro 14", "3", "1250.50", "10%", "Lucía Paredes", "Pagado"],
    [" ped-00002 ", "07/01/2026", "NOVA DISTRIBUCIONES ", "norte", "laptop pro14", "2", "1,180.00", "0", "lucía paredes", "pagado"],  # ID, fecha, texto y producto sucios
    ["PED00003", "08-01-26", "Grupo Aurora", "Nrte", "Monitor 27 4k", "5", "USD 340", "5 %", "Marcos Quiroz", "Pend."],  # región con error, importe con USD
    ["PED-00004", "09.01.2026", "Logística  Delta", "Sur", "Teclado Mecanico", "12", "78,5", "0.05", "", "Pagado"],  # coma decimal, vendedor vacío
    ["PED-00005", "31/02/2026", "Cobalto Industrial", "Centro", "Mouse Inalámbrico", "10", "$ 24.90", "", "Valeria Soto", "Pagado"],  # fecha imposible
    ["PED-00006", "12/01/2026", "Alfa Retail", "Oriente", "Silla Ergonómica", "-2", "210", "0", "Diego Rivas", "Pagado"],  # cantidad negativa
    ["PED-00002", "07/01/2026", "Nova Distribuciones", "Norte", "Laptop Pro 14", "2", "1180.00", "0", "Lucía Paredes", "Pagado"],  # duplicado de la fila 2
    ["PED-00008", "15/01/2026", "Tecnoandina", "Sur", "Laptop Pro 14", "1", "1.250,50", "15%", "Camila Herrera", "Anulado"],  # formato europeo, pedido anulado
    ["PED-00009", "15/03/2062", "Corporación Zenit", "Centro", "Docking USB-C", "4", "125", "0", "Andrés Molina", "Pagado"],  # año mal digitado
    ["PED-00010", "20/01/2026", "Corporación Zenit", "Centro", "Webcam HD 1080p", "99999", "45", "0", "Andrés Molina", "Pagado"],  # cantidad absurda
    ["PED-00011", "21/01/2026", "Mercantil Horizonte", "Marte", "Auriculares ANC", "6", "120", "10%", "Camila Herrera", "Pagado"],  # región inexistente
    ["PED-00012", "22/01/2026", "Soluciones Orbital", "Sur", "Cable HDMI 2m", "1", "99", "0", "Diego Rivas", "Pagado"],  # producto fuera del catálogo
]
LINEA_ROTA = "PED-00350;12/05/2026;Soluciones; Orbital;Sur;Silla Ergonómica;2;210;0;Diego Rivas;Pagado"  # ';' sin comillas
LINEA_TRUNCADA = "PED-00351;13/05/2026;Alfa Retail;Oriente;Docking USB-C"
PIE_DE_REPORTE = "TOTAL;;;;;;;;;"


def sin_acentos(texto: str) -> str:
    return unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii")


# ---------------------------------------------------------------------------
# Cada función recibe el valor limpio y lo devuelve como lo exportaría un ERP
# descuidado: a veces correcto, a veces con ruido.
# ---------------------------------------------------------------------------
def sucio_id(rng: random.Random, n: int) -> str:
    base = f"PED-{n:05d}"
    return rng.choices([base, base.lower(), f" {base} ", base.replace("-", ""), ""],
                       weights=[90, 3, 3, 3, 1])[0]


def sucio_fecha(rng: random.Random, d: date) -> str:
    r = rng.random()
    if r < 0.02:
        return rng.choice(["31/02/2026", "2026-13-05", "ayer", "", "00/00/0000"])
    if r < 0.03:
        return d.replace(year=2062).strftime("%d/%m/%Y")
    formatos = [d.isoformat(), d.strftime("%d/%m/%Y"), d.strftime("%d-%m-%y"),
                d.strftime("%d.%m.%Y"), d.strftime("%d-%m-%Y")]
    return rng.choices(formatos, weights=[30, 40, 10, 10, 10])[0]


def sucio_nombre(rng: random.Random, nombre: str, p_vacio: float) -> str:
    if rng.random() < p_vacio:
        return ""
    variantes = [nombre, nombre.upper(), nombre.lower(), f"{nombre} ", f" {nombre}",
                 nombre.replace(" ", "  "), f"{nombre}."]
    return rng.choices(variantes, weights=[76, 6, 4, 5, 3, 3, 3])[0]


def sucio_region(rng: random.Random, region: str) -> str:
    r = rng.random()
    if r < 0.03:
        return ""
    if r < 0.05:
        return rng.choice(["N/A", "???", "Marte"])
    errores = {"Norte": ["Nrte", "N0rte"], "Sur": ["Sr", "Sur."],
               "Centro": ["Centrp", "Cntro"], "Oriente": ["Orient", "Oriemte"]}
    variantes = [region, region.lower(), region.upper(), f"{region} "] + errores[region]
    return rng.choices(variantes, weights=[70, 6, 5, 5, 7, 7])[0]


def sucio_producto(rng: random.Random, nombre: str) -> str:
    if rng.random() < 0.01:
        return rng.choice(["Cable HDMI 2m", "Producto X", "Servicio instalación"])
    i = rng.randrange(1, len(nombre))
    con_error = nombre[:i] + nombre[i + 1:]  # falta una letra
    variantes = [nombre, nombre.lower(), nombre.upper(), sin_acentos(nombre),
                 nombre.replace(" ", "  "), con_error]
    return rng.choices(variantes, weights=[72, 6, 4, 8, 3, 7])[0]


def sucio_cantidad(rng: random.Random, c: int) -> str:
    r = rng.random()
    if r < 0.010:
        return rng.choice(["-3", "-1", "0"])
    if r < 0.015:
        return ""
    if r < 0.020:
        return "tres"
    if r < 0.024:
        return str(c * 100)  # ceros de más: pasa la validación, queda como atípico
    if r < 0.026:
        return "99999"
    return rng.choices([str(c), f"{c} unid.", f"{c:02d}", f"{c}.0"], weights=[85, 6, 5, 4])[0]


def sucio_precio(rng: random.Random, p: float) -> str:
    r = rng.random()
    if r < 0.010:
        return ""
    if r < 0.015:
        return rng.choice(["N/A", "gratis", f"-{p:.2f}"])
    miles_us = f"{p:,.2f}"
    miles_eu = miles_us.translate(str.maketrans(",.", ".,"))
    variantes = [f"{p:.2f}", miles_us, miles_eu, f"${miles_us}", f"USD {p:.0f}",
                 f"{p:.2f}".replace(".", ","), f"US$ {miles_us}"]
    return rng.choices(variantes, weights=[40, 15, 12, 10, 8, 10, 5])[0]


def sucio_descuento(rng: random.Random, d: float) -> str:
    if rng.random() < 0.005:
        return rng.choice(["250%", "abc"])
    pct = round(d * 100)
    variantes = [f"{pct}%", f"{pct} %", f"{d:g}", f"{d:.2f}".replace(".", ","),
                 f"{pct}", "" if pct == 0 else f"{pct}%"]
    return rng.choices(variantes, weights=[30, 10, 15, 10, 20, 15])[0]


def sucio_estado(rng: random.Random, estado: str) -> str:
    if rng.random() < 0.003:
        return "???"
    variantes = {
        "Pagado": ["Pagado", "pagado", "PAGADO ", "Pagada"],
        "Pendiente": ["Pendiente", "PENDIENTE", "pend.", "Pdte"],
        "Anulado": ["Anulado", "ANULADO", "anulada", "Cancelado"],
    }[estado]
    return rng.choices(variantes, weights=[70, 12, 8, 10])[0]


# ---------------------------------------------------------------------------
# Construcción del archivo
# ---------------------------------------------------------------------------
def fechas_con_tendencia(rng: random.Random, n: int, desde: date, hasta: date) -> list[date]:
    """Más ventas hacia el final del período y menos en fin de semana."""
    dias = [desde + timedelta(days=i) for i in range((hasta - desde).days + 1)]
    pesos = [(1 + 1.2 * i / len(dias)) * (0.35 if d.weekday() >= 5 else 1) for i, d in enumerate(dias)]
    return sorted(rng.choices(dias, weights=pesos, k=n))


def fila_sucia(rng: random.Random, n: int, d: date) -> list[str]:
    cliente, (region, vendedor) = rng.choice(list(CLIENTES.items()))
    producto, (pmin, pmax, cmax) = rng.choice(list(PRODUCTOS.items()))
    cantidad = rng.randint(1, cmax)
    precio = round(rng.uniform(pmin, pmax), 2)
    descuento = rng.choice([0, 0, 0, 0.05, 0.05, 0.10, 0.15])
    estado = rng.choices(["Pagado", "Pendiente", "Anulado"], weights=[68, 24, 8])[0]
    return [
        sucio_id(rng, n), sucio_fecha(rng, d), sucio_nombre(rng, cliente, 0.01),
        sucio_region(rng, region), sucio_producto(rng, producto), sucio_cantidad(rng, cantidad),
        sucio_precio(rng, precio), sucio_descuento(rng, descuento),
        sucio_nombre(rng, vendedor, 0.03), sucio_estado(rng, estado),
    ]


def construir_registros(rng: random.Random, filas: int) -> list[list[str] | str]:
    """Devuelve filas (listas de campos) y líneas crudas (str) que rompen el formato."""
    fechas = fechas_con_tendencia(rng, filas, date(2026, 1, 1), date(2026, 9, 30))
    aleatorias = [fila_sucia(rng, n, d) for n, d in enumerate(fechas, start=len(CASOS_FIJOS) + 1)]

    # Duplicados: copias exactas que aparecen más abajo en el archivo
    for _ in range(int(len(aleatorias) * 0.05)):
        origen = rng.randrange(len(aleatorias))
        aleatorias.insert(rng.randrange(origen + 1, len(aleatorias) + 1), list(aleatorias[origen]))

    registros: list[list[str] | str] = [*CASOS_FIJOS, *aleatorias]
    # Líneas en blanco y líneas con estructura rota, en posiciones al azar
    for linea in ["", "", "", LINEA_ROTA, LINEA_TRUNCADA]:
        registros.insert(rng.randrange(len(CASOS_FIJOS) + 1, len(registros)), linea)
    registros.append(PIE_DE_REPORTE)
    return registros


def escribir_csv(destino: Path, registros: list[list[str] | str], encoding: str) -> None:
    buffer = io.StringIO()
    escritor = csv.writer(buffer, delimiter=";", lineterminator="\n")
    escritor.writerow(ENCABEZADO)
    for registro in registros:
        if isinstance(registro, str):
            buffer.write(registro + "\n")
        else:
            escritor.writerow(registro)
    destino.parent.mkdir(parents=True, exist_ok=True)
    destino.write_text(buffer.getvalue(), encoding=encoding)


def main() -> None:
    p = argparse.ArgumentParser(description="Genera un CSV de ventas con datos sucios para probar el pipeline.")
    p.add_argument("--filas", type=int, default=600, help="filas aleatorias además de los casos fijos (por defecto: 600)")
    p.add_argument("--semilla", type=int, default=42, help="semilla del generador (por defecto: 42)")
    p.add_argument("--salida", type=Path, default=DESTINO, help="ruta del CSV a crear")
    p.add_argument("--encoding", default="utf-8", help="codificación del archivo (por defecto: utf-8)")
    args = p.parse_args()

    registros = construir_registros(random.Random(args.semilla), args.filas)
    escribir_csv(args.salida, registros, args.encoding)
    print(f"Archivo generado: {args.salida}  ({len(registros)} líneas de datos, semilla {args.semilla})")


if __name__ == "__main__":
    main()
