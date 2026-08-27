@echo off
setlocal

set "DIR_PROYECTO=%~dp0"
set "DIR_PLANTILLAS=%DIR_PROYECTO%plantillas"
set "DIR_DIST=%DIR_PROYECTO%dist\main"
set "PYTHON=%HOMEDRIVE%%HOMEPATH%\AppData\Local\Programs\Python\Python313\python.exe"

echo ============================================
echo  Instalando dependencias...
echo ============================================
"%PYTHON%" -m pip install -r "%DIR_PROYECTO%requirements.txt"

echo.
echo ============================================
echo  Compilando ejecutable...
echo ============================================
"%PYTHON%" -m PyInstaller --noconfirm --noconsole --onedir ^
    --collect-all customtkinter ^
    --collect-all docxtpl ^
    --collect-all pdfplumber ^
    --collect-all openpyxl ^
    --distpath "%DIR_PROYECTO%dist" ^
    "%DIR_PROYECTO%main.py"

echo.
echo ============================================
echo  Copiando carpeta plantillas al ejecutable...
echo ============================================
if exist "%DIR_PLANTILLAS%" (
    if not exist "%DIR_DIST%\plantillas" mkdir "%DIR_DIST%\plantillas"
    xcopy /E /Y /I "%DIR_PLANTILLAS%\*" "%DIR_DIST%\plantillas\"
    echo Carpeta "plantillas" copiada correctamente.
) else (
    echo AVISO: No se encontro la carpeta "plantillas". Creandola...
    mkdir "%DIR_DIST%\plantillas"
)

echo.
echo ============================================
echo  Creando carpeta Informes_Generados...
echo ============================================
if not exist "%DIR_DIST%\Informes_Generados" mkdir "%DIR_DIST%\Informes_Generados"

echo.
echo ============================================
echo  Desbloqueando ejecutable (evita alerta de Windows)...
echo ============================================
powershell -Command "Unblock-File -Path '%DIR_DIST%\main.exe'" 2>nul
echo Ejecutable desbloqueado.

echo.
echo ============================================
echo  Limpiando archivos temporales...
echo ============================================
if exist "%DIR_PROYECTO%build" rmdir /s /q "%DIR_PROYECTO%build"
if exist "%DIR_PROYECTO%__pycache__" rmdir /s /q "%DIR_PROYECTO%__pycache__"
echo Limpieza completada.

echo.
echo ============================================
echo  Proceso finalizado.
echo  Tu aplicacion esta lista en:
echo  %DIR_DIST%\main.exe
echo ============================================
pause
