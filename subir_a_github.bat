@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ====================================================================
echo  Sincronizando cambios con GitHub: App_Informes_Offline
echo ====================================================================
echo.
echo [1/4] Agregando archivos unificados y plantillas...
git add .gitignore README.md gui.py main.spec generar_ejecutable.bat subir_a_github.bat procesador.py procesador_actualizacion.py generador_word.py generador_word_actualizacion.py generador_excel.py generador_excel_actualizacion.py generador_oficio.py plantillas
echo.
echo [2/4] Creando commit...
git commit -m "feat: unificacion de modulos Ampliaciones y Actualizacion (De suspendido a regular) con nuevas plantillas"
echo.
echo [3/4] Sincronizando con cambios remotos de GitHub...
git pull --rebase origin main
echo.
echo [4/4] Subiendo cambios a GitHub...
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
