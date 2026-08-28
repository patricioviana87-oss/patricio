@echo off
cd /d "%~dp0"
py interfaz.py
if errorlevel 1 (
    echo.
    echo Hubo un error. Copia el mensaje de arriba y envialo para revisarlo.
    pause
)
