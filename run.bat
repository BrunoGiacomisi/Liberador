@echo off
REM Arranca el Generador de Libres en modo desarrollo
REM Requiere haber instalado antes las dependencias: pip install -r requirements.txt

python main.py

if errorlevel 1 (
    echo.
    echo La aplicacion termino con un error.
    pause
)
