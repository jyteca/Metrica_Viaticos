@echo off
echo Preparando la compilacion del ejecutable Ligero de Gestion de Viaticos...
echo.
echo Se instalara PyInstaller si es necesario...
pip install pyinstaller

echo.
echo Compilando aplicacion. Esto puede tardar algun minuto...
python -m PyInstaller --noconfirm --onedir --windowed --name="Metrica_Viaticos_Ligero" --add-data "utils;utils" --add-data "assets;assets" app.py

echo.
echo Proceso de compilacion finalizado!
echo Tu programa esta ubicado en formato ejecutable liviano dentro de la carpeta "dist/Metrica_Viaticos_Ligero"
echo No olvides copiar tambien tu carpeta "config" (con el historial y configuracion) al lado de ".exe" para que lea los datos correctamente.
echo.
pause
