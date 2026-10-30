@echo off
cd /d "%~dp0"
echo ============================================
echo   CARTERA AUTO - Prototipo COLBEEF
echo ============================================
echo.
echo Instalando dependencias si faltan...
python -m pip install -r requirements.txt -q
echo.
echo Abriendo la aplicacion en el navegador...
echo (Deja esta ventana abierta mientras la uses)
echo.
python -m streamlit run app.py --server.headless true
pause
