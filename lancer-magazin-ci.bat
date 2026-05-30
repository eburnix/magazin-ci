@echo off
cd /d "%~dp0"
set /p PIN="Entrez le PIN d'installation : "
if "%PIN%" neq "05535350" (
    echo PIN incorrect. Fermeture.
    pause
    exit /b 1
)
python app.py
pause
