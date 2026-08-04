@echo off
REM Doble clic para abrir el Centro de Carga (Windows).
cd /d "%~dp0"
python centro_carga.py
if errorlevel 1 pause
