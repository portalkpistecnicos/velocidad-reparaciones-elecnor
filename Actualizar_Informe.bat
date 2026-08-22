@echo off
setlocal
set SCRIPT_DIR=%~dp0
echo Actualizando Informe P31 - Velocidad de Reparaciones...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Actualizar_Informe.ps1"

echo.
if %errorlevel% neq 0 (
    echo Ocurrio un error al actualizar el informe. Revise el mensaje de arriba.
) else (
    echo Proceso terminado. Puede cerrar esta ventana.
)
pause
