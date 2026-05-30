@echo off
REM Build the Magazin-ci installer using Inno Setup.
cd /d "%~dp0"
if not exist "dist\Magazin-ci.exe" (
    echo Error: dist\Magazin-ci.exe not found. Build the executable first.
    pause
    exit /b 1
)
set "ISCC_PATH="
if exist "C:\Program Files\Inno Setup 7\ISCC.exe" (
    set "ISCC_PATH=C:\Program Files\Inno Setup 7\ISCC.exe"
) else (
    if exist "C:\Program Files (x86)\Inno Setup 5\ISCC.exe" (
        set "ISCC_PATH=C:\Program Files (x86)\Inno Setup 5\ISCC.exe"
    )
)
if "%ISCC_PATH%"=="" (
    where ISCC >nul 2>&1
    if errorlevel 1 (
        echo Error: ISCC (Inno Setup compiler) not found. Install Inno Setup or add ISCC to PATH.
        pause
        exit /b 1
    ) else (
        set "ISCC_PATH=ISCC"
    )
)
"%ISCC_PATH%" "%~dp0Magazin-ci-installer.iss"
if %ERRORLEVEL% equ 0 (
    echo Installer created in installateur\Setup-Magazin-ci.exe
) else (
    echo Error compiling installer.
)
pause
