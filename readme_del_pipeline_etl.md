# ⚙️ Pipeline Automatizado de Extracción, Procesamiento y Reporte de Datos (ETL B2B)

![Python](https://img.shields.io/badge/Python-3.9%2B-blue.svg)
![Pandas](https://img.shields.io/badge/Pandas-Data%20Analysis-150458.svg)
![Architecture](https://img.shields.io/badge/Architecture-ETL%20Modular-green.svg)
![License](https://img.shields.io/badge/License-MIT-lightgrey.svg)

## 📌 Descripción del Proyecto

Este proyecto consiste en un **Pipeline de Ingeniería de Datos (ETL) automatizado, modular y robusto** desarrollado en Python. Ha sido diseñado para resolver un problema crítico recurrente en entornos corporativos B2B: la ingesta e integración de información de ventas o transacciones desordenadas, incompletas y corruptas procedentes de sistemas legados o exports de ERPs.

El sistema ejecuta de forma autónoma la ingesta, limpieza profunda, validación de reglas de negocio, manejo de excepciones y la generación automatizada de reportes ejecutivos consolidados en formato **Excel (`.xlsx`)** y **CSV**, acompañados de un sistema integral de **logs de auditoría**.

---

## 🎬 Demostración en Vivo

> **[🎥 Haz clic aquí para ver la demostración en video del Pipeline ETL (35s)](https://youtu.be/Ad95OagzKYw)**

---

## 🏗️ Arquitectura del Sistema

El pipeline implementa una arquitectura modular desacoplada basada en 3 fases secuenciales:

```text
┌────────────────────────┐
│     FASE 1: INGESTA    │ ──► Lectura e inspección de CSV/Excel crudos (data/raw/)
└───────────┬────────────┘
            │
┌───────────▼────────────┐
│ FASE 2: TRANSFORMACIÓN │ ──► Limpieza, normalización, imputación y reglas de negocio
└───────────┬────────────┘
            │
┌───────────▼────────────┐
│  FASE 3: EXPORTACIÓN   │ ──► Generación de Excel Ejecutivo (output/) y logs de auditoría (logs/)
└────────────────────────┘
```

1. **Ingesta (`data/raw/`):** Captura segura del dataset crudo sin alterar los archivos fuente.
2. **Transformación (Pandas/NumPy):**
   * Normalización de cadenas de texto (mayúsculas, minúsculas, eliminación de espacios en blanco).
   * Conversión vectorizada de tipos de datos (`datetime`, `float`).
   * Eliminación de duplicados operativos y tratamiento de valores nulos (`NaN`).
   * Aplicación de filtros según reglas de negocio (ej. descarte de pedidos anulados o montos no numéricos).
3. **Exportación (`output/` & `logs/`):**
   * Generación de reportes ejecutivos consolidados en Excel con formato profesional mediante `XlsxWriter`.
   * Trazabilidad completa mediante archivos de registro (`.log`) para auditoría técnica.

---

## 🛠️ Tecnologías Utilizadas

* **Lenguaje:** Python 3.9+
* **Procesamiento de Datos:** `pandas`, `numpy`
* **Generación de Reportes Excel:** `xlsxwriter`, `openpyxl`
* **Manejo de Entorno y Rutas:** `pathlib`, `os`, `sys`
* **Trazabilidad:** `logging`

---

## 📁 Estructura del Repositorio

```text
pipeline_proyecto/
│
├── data/
│   └── raw/                  # Datos de entrada crudos (ventas_sucias.csv)
├── logs/                     # Archivos de auditoría técnica (.log)
├── output/                   # Reportes finales generados (.xlsx y .csv)
│
├── generar_datos_sucios.py   # Script generador de datasets sintéticos de prueba
├── pipeline.py               # Motor principal del Pipeline ETL
├── requirements.txt          # Dependencias y librerías necesarias
└── README.md                 # Documentación del proyecto
```

---

## 🚀 Instalación y Uso Local

### Prerrequisitos
Asegúrate de tener instalado Python 3.9 o superior en tu sistema.

### 1. Clonar el repositorio y navegar a la carpeta
```bash
git clone https://github.com/papalla123/pipeline-etl-python.git
cd pipeline-etl-python
```

### 2. Instalar dependencias
Se recomienda utilizar la versión directa del gestor de paquetes de Python:
```bash
python -m pip install -r requirements.txt
```

### 3. Generar dataset de prueba (Datos Sucios)
Ejecuta el script para simular un export crudo de ERP con errores intencionados:
```bash
python generar_datos_sucios.py
```
*Esto generará el archivo `ventas_sucias.csv` dentro de la carpeta `data/raw/`.*

### 4. Ejecutar el Pipeline ETL
Inicia el proceso completo de limpieza y exportación:
```bash
python pipeline.py
```

---

## 📊 Resultados Esperados

Al finalizar la ejecución (`python pipeline.py`), el sistema generará automáticamente:

* **En `output/`:**
  * `reporte_ejecutivo_YYYYMMDD_HHMMSS.xlsx`: Dashboard analítico consolidado por cliente, producto y región.
  * `ventas_limpias_YYYYMMDD_HHMMSS.csv`: Dataset estandarizado y listo para ser consumido por herramientas de BI (Power BI / Tableau).
* **En `logs/`:**
  * Registro de auditoría detallado con el recuento de registros procesados, porcentaje de datos válidos vs. excluidos y tiempo total de ejecución.

---

## 🔄 Adaptación para Nuevos Clientes / Datos Reales

Para procesar datos de un cliente real:
1. Reemplaza el archivo situado en `data/raw/` por el CSV entregado por el cliente, asegurándote de nombrarlo `ventas_sucias.csv`.
2. Ejecuta `python pipeline.py`.
3. El sistema procesará el nuevo conjunto de datos y depositará los resultados limpios en la carpeta `output/`.

---

## 📜 Licencia

Este proyecto se encuentra bajo la Licencia [MIT](LICENSE). Libre para uso, modificación y distribución.