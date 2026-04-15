@echo off
echo Inicializando repositorio local de Git...
git init
git branch -M main

echo.
echo Agregando archivos al paquete (excluyendo lo pesado segun .gitignore)...
git add .
git commit -m "Solución de Viaticos y Fondos (Streamlit)"

echo.
echo Conectando con tu repositorio en GitHub...
git remote add origin https://github.com/jyteca/Metrica_Viaticos.git

echo.
echo Subiendo proyecto a Internet...
echo (NOTA: Si es tu primera vez, puede abrirse una ventanita de Windows pidiendote iniciar sesion en Github o dar acceso)
git push -u origin main

echo.
echo Proceso de subida finalizado con exito! 
pause
