# --- COMPATIBILIDAD DE CARACTERES ---
$OutputEncoding = [System.Text.Encoding]::UTF8
# ---------------------------------------------------

# =================================================================================
#  SCRIPT DE POWERSHELL PARA AUTOMATIZAR LA EJECUCIÓN DE UN SCRIPT DE PYTHON
#  Fecha: 23 de septiembre de 2025
# =================================================================================

# --- Configuración ---
$pythonScriptName = "GenerarReportes.py"
$venvName = "entorno_virtual" # Nombre de la carpeta para el entorno virtual
$requirementsFile = "requirements.txt"

# --- Inicio del Script ---
Clear-Host
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "  Automatizador de Reportes en Excel          " -ForegroundColor Cyan
Write-Host "=============================================="

# --- # Paso 1: VALIDACIÓN ---
Write-Host "[1/7] Validando archivos de entrada..." -ForegroundColor Yellow
if (-not (Test-Path ".\$inputFolder")) {
    Write-Host "❌ ERROR: No se encontró la carpeta de entrada '$inputFolder'." -ForegroundColor Red
    Read-Host "Presiona Enter para salir."
    exit
}
if (-not (Test-Path ".\$mainInputFile")) {
    Write-Host "❌ ERROR: El archivo principal '$mainInputFile' no se encontró. El script no puede continuar." -ForegroundColor Red
    Read-Host "Presiona Enter para salir."
    exit
}
Write-Host "✅ Archivos principales encontrados." -ForegroundColor Green
# --- FIN DE LA NUEVA SECCIÓN ---

# Paso 2: Verificar que Python esté instalado y en el PATH
Write-Host "[2/7] Verificando la instalación de Python..." -ForegroundColor Yellow
$pythonExe = Get-Command python -ErrorAction SilentlyContinue
if (-not $pythonExe) {
    Write-Host "❌ ERROR: Python no se encontró. Por favor, instálalo y asegúrate de que esté en el PATH del sistema." -ForegroundColor Red
    Read-Host "Presiona Enter para salir."
    exit
}
Write-Host "✅ Python encontrado." -ForegroundColor Green

# Paso 3: Verificar que el script de Python exista
Write-Host "[3/7] Verificando que '$pythonScriptName' exista..." -ForegroundColor Yellow
if (-not (Test-Path ".\$pythonScriptName")) {
    Write-Host "❌ ERROR: El script '$pythonScriptName' no se encuentra en esta carpeta." -ForegroundColor Red
    Read-Host "Presiona Enter para salir."
    exit
}
Write-Host "✅ Script de Python encontrado." -ForegroundColor Green

# Paso 4: Verificar que el archivo de requerimientos exista
Write-Host "[4/7] Verificando que '$requirementsFile' exista..." -ForegroundColor Yellow
if (-not (Test-Path ".\$requirementsFile")) {
    Write-Host "❌ ERROR: El archivo '$requirementsFile' no se encuentra en esta carpeta." -ForegroundColor Red
    Read-Host "Presiona Enter para salir."
    exit
}
Write-Host "✅ Archivo de requerimientos encontrado." -ForegroundColor Green

# Paso 5: Crear el entorno virtual si no existe
Write-Host "[5/7] Configurando el entorno virtual ('$venvName')..." -ForegroundColor Yellow
if (-not (Test-Path ".\$venvName")) {
    Write-Host "    Creando entorno virtual (esto puede tardar un momento)..."
    python -m venv $venvName
    Write-Host "✅ Entorno virtual creado." -ForegroundColor Green
} else {
    Write-Host "✅ El entorno virtual ya existe." -ForegroundColor Green
}

# Paso 6: Instalar dependencias desde requirements.txt
Write-Host "[6/7] Instalando/verificando librerías desde '$requirementsFile'..." -ForegroundColor Yellow
$pipPath = ".\$venvName\Scripts\pip.exe"
# pip instalará solo lo que falte. Si ya está todo instalado, no hará nada.
& $pipPath install -r $requirementsFile --quiet --no-warn-script-location

if ($?) { # Verifica si el último comando fue exitoso
   Write-Host "✅ Librerías sincronizadas correctamente." -ForegroundColor Green
} else {
   Write-Host "❌ ERROR: Hubo un problema al instalar las librerías." -ForegroundColor Red
   Read-Host "Presiona Enter para salir."
   exit
}

# --- FORZAR A PYTHON A USAR UTF-8 ---
$env:PYTHONUTF8 = "1"
# ---------------------------------------------------

#Guardamos la hora de inicio
$inicio = Get-Date

# Paso 7: Ejecutar el script de Python
Write-Host "[7/7] Ejecutando el script de Python para generar los reportes..." -ForegroundColor Yellow
# Usamos el intérprete de Python del entorno virtual
$pythonInterpreter = ".\$venvName\Scripts\python.exe"
# Ejecutamos el script de Python directamente (con -u)
# Esto permitirá que la salida se muestre en tiempo real.
& $pythonInterpreter -u $pythonScriptName  
# Guardamos la hora de finalización
$fin = Get-Date
# Calculamos la diferencia
$duracion = $fin - $inicio
# Formateamos y mostramos el tiempo total
$tiempoFormateado = "{0:hh}:{0:mm}:{0:ss}" -f $duracion
Write-Host "Tiempo total de ejecución: $tiempoFormateado" -ForegroundColor Green

Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "  PROCESO FINALIZADO. Revisa la carpeta 'reportes_generados'." -ForegroundColor Green
Write-Host "=============================================="
Read-Host "Presiona Enter para cerrar esta ventana."