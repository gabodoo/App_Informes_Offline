@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ====================================================================
echo  Sincronizando cambios con GitHub: App_Informes_Offline
echo ====================================================================
echo.
echo [1/5] Estado de archivos en el repositorio:
git status -s
echo.
echo [2/5] Agregando archivos de codigo, configuracion y plantillas...
git add main.py requirements.txt gui.py procesador.py procesador_actualizacion.py generador_word.py generador_word_actualizacion.py generador_excel.py generador_excel_actualizacion.py generador_oficio.py main.spec generar_ejecutable.bat subir_a_github.bat README.md .gitignore plantillas
echo.
echo [3/5] Creando commit con cambios pendientes...
git commit -m "feat: unificacion de modulos, actualizacion de dependencias, spec relativo y plantillas"
echo.
echo [4/5] Sincronizando con cambios remotos de GitHub (pull --rebase)...
git pull --rebase origin main
echo.
echo [5/5] Subiendo cambios a GitHub...
git push origin main
if %ERRORLEVEL% neq 0 (
    echo.
    echo [Aviso] Intentando subir a rama master...
    git push origin master
)
echo.
echo ====================================================================
echo  Completado exitosamente. Presiona cualquier tecla para cerrar.
echo ====================================================================
pause
