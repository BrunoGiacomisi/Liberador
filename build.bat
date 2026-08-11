@echo off
REM Genera el ejecutable GeneradorDeLibres.exe (carpeta dist\)
REM Requiere haber instalado antes las dependencias: pip install -r requirements.txt

pyinstaller --noconfirm --onefile --windowed ^
    --name "GeneradorDeLibres" ^
    --add-data "templates;templates" ^
    --add-data "assets;assets" ^
    main.py

echo.
echo Listo. El ejecutable quedo en la carpeta dist\GeneradorDeLibres.exe
pause
