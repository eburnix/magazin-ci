@echo off
REM Build the Magazin-ci executable using PyInstaller.
cd /d "%~dp0"
python -m PyInstaller --noconfirm Magazin-ci.spec
if %ERRORLEVEL% equ 0 (
    echo Build termine. Le binaire se trouve dans dist\Magazin-ci.exe
) else (
    echo Erreur de build.
)
pause
