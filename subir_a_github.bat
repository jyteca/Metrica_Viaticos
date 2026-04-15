@echo off
echo Actualizando los archivos hacia tu GitHub...
echo.

git add .
git commit -m "Reparando dependencias de libcairo en servidor (packages.txt)"
git push origin main

echo.
echo Codigo sincronizado! 
echo Vuelve a la pantalla de Streamlit Cloud y presiona 'Re-deploy' o 'Reboot app'.
pause
