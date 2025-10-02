# Automatización de Generación de Reportes en Excel

Este proyecto contiene un conjunto de scripts diseñados para automatizar la creación de reportes financieros y de planificación. El sistema lee datos de múltiples hojas y archivos de Excel, realiza cálculos, mapea datos a una plantilla prediseñada y genera reportes individuales por cada "Cost Center".

## ✨ Características Principales

* **Procesamiento por Lotes:** Genera automáticamente un reporte de Excel para cada `Cost Center` encontrado en el archivo de origen.
* **Múltiples Fuentes de Datos:** Capaz de leer y combinar información de un archivo de Excel estático y de múltiples archivos dinámicos (uno por cada `Cost Center`).
* **Mapeo y Cálculos:**
  * Mapea datos de hojas de origen a tablas específicas en una plantilla de Excel.
  * Realiza cálculos complejos (`SUMAR.SI`, copias de valores) y los inserta en celdas específicas.
* **Copia de Hojas con Formato:** Replica hojas completas de un archivo de origen a los reportes de salida, preservando el formato visual (colores, bordes, celdas combinadas) y la estructura de tablas.
* **Interactividad en Excel:** Añade menús desplegables y fórmulas condicionales a los reportes generados para permitir una interacción controlada por el usuario final.
* **Automatización Completa:** Utiliza un script de PowerShell como lanzador para crear un entorno virtual, instalar dependencias y ejecutar el proceso de Python con un solo clic.
* **Registro y Monitoreo:** Incluye una barra de progreso visual durante la ejecución y mide el tiempo total del proceso.

## 📁 Estructura del Proyecto

Para que el script funcione correctamente, la estructura de carpetas debe ser la siguiente:

```
Mi_Proyecto/
│
├── 📁 Datos de Entrada/
│   ├── ScenarioPlanningDB.xlsx       # Archivo principal con datos de BWP, Envelope, etc.
│   ├── AdaptiveBLZ.xlsx              # Archivo dinámico para el Cost Center 'BLZ'
│   ├── AdaptiveARG.xlsx              # Archivo dinámico para el Cost Center 'ARG'
│   └── .gitkeep
│
├── 📁 Plantilla/
│   ├── Plantilla OP2627ScenarioPlanning.xlsx # Plantilla base para los reportes
│   └── .gitkeep
│
├── 📁 Reportes Generados/
│   └── .gitkeep                      # Aquí se guardarán los reportes finales
│
├── 🚀 EjecutarProceso.ps1             # Script de PowerShell para iniciar todo el proceso
├── 🐍 GenerarReportes.py                # Script principal de Python con toda la lógica
├── 📋 requirements.txt               # Lista de librerías de Python necesarias
├── 📄 config.json                     # Archivo de configuración para todos los parámetros
└── 📄 README.md                       # Este archivo
```

## ⚙️ Configuración

Toda la configuración del script se gestiona desde el archivo `config.json`. Esto permite modificar parámetros sin tocar el código Python. Las secciones principales son:

* **`archivos`**: Rutas de los archivos de entrada, plantilla y carpeta de salida.
* **`parametros_globales`**: Define la hoja maestra y las columnas clave para identificar los `Cost Centers`.
* **`plantilla_salida`**: Configura detalles del archivo de salida, como las celdas de inicio y los nombres de las tablas.
* **`interactividad`**: Parámetros para la creación de menús desplegables y fórmulas en Excel.
* **`mapeo_principal`**: Define cómo se llenan las tablas principales en la plantilla a partir de los datos de origen.
* **`lista_calculos`**: Una lista detallada de todas las operaciones (`SUMAR.SI`, `COPIA`) que el script debe realizar, especificando hojas, celdas y columnas.

## 🚀 Requisitos y Ejecución

### Requisitos

* **Python 3.7+** instalado y añadido al PATH de Windows.
* **Windows Terminal** (recomendado para una correcta visualización de la salida del script).

### Ejecución

1. **Clonar el Repositorio:** Descarga o clona este repositorio en tu máquina.
2. **Poblar las Carpetas:** Coloca los archivos de Excel requeridos en las carpetas `Datos de Entrada` y `Plantilla`.
3. **Ejecutar el Lanzador:** Haz clic derecho en el archivo `ejecutar_proceso.ps1` y selecciona **"Ejecutar con PowerShell"**. También puedes usar el acceso directo o el archivo `.bat` si lo has creado.

El script de PowerShell se encargará automáticamente de:

* Verificar que los archivos necesarios existan.
* Crear un entorno virtual de Python en la carpeta `entorno_virtual/`.
* Instalar las librerías listadas en `requirements.txt`.
* Ejecutar el script principal de Python (`separar_excel.py`) para generar los reportes.

Los reportes generados aparecerán en la carpeta `Reportes Generados`.

---

Autor: Dennis Fraile
