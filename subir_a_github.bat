@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ====================================================================
echo  Sincronizando cambios con GitHub: App_Informes_Offline
echo ====================================================================
echo.
echo [1/5] Estado previo del repositorio:
git status -s
echo.
echo [2/5] Agregando todos los archivos del proyecto (codigo y plantillas)...
git add -A
echo.
echo [3/5] Guardando cambios locales en commit...
git commit -m "feat: unificacion de modulos, actualizacion de dependencias, spec relativo y plantillas"
echo.
echo [4/5] Asegurando arbol limpio y sincronizando con GitHub...
git stash
git pull origin main --rebase
if %ERRORLEVEL% neq 0 (
    echo [Aviso] Rebase encontro conflictos o diferencias. Intentando fusion normal...
    git rebase --abort 2>nul
    git pull origin main --no-rebase --no-edit
)
git stash pop 2>nul
echo.
echo [5/5] Subiendo cambios a GitHub (rama main)...
git push origin main
if %ERRORLEVEL% neq 0 (
    echo.
    echo ====================================================================
    echo  [AVISO IMPORTANTE]
    echo  GitHub rechazo la subida porque tu version remota tiene cambios
    echo  que no coinciden con tu version local actual.
    echo.
    echo  Si tu version local es la definitiva y deseas sobrescribir GitHub:
    echo     git push origin main --force
    echo ====================================================================
) else (
    echo.
    echo ====================================================================
    echo  Completado exitosamente. Sincronizado con GitHub.
    echo ====================================================================
)
pause
