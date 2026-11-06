@echo off
rem Abre QuimNexo en el navegador. Corre solo en este equipo.
rem Para cerrarlo: cierre esta ventana negra.
cd /d "%~dp0"
echo Abriendo QuimNexo...
echo (No cierre esta ventana mientras lo use. Para terminar, cierrela.)
start "" /min cmd /c "timeout /t 4 >nul & start http://localhost:8501"
py -m streamlit run "clasificador SGA\app.py" --server.headless true
pause
