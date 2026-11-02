@echo off
setlocal
cd /d "%~dp0"

set "BUILD_ENV=%~dp0.build_venv"
set "PYTHON=%BUILD_ENV%\Scripts\python.exe"

if not exist "%PYTHON%" (
    echo Creating isolated build environment...
    python -m venv "%BUILD_ENV%"
    if errorlevel 1 goto :failed
)

echo Installing build dependencies...
"%PYTHON%" -m pip install -r "%~dp0requirements-build.txt"
if errorlevel 1 goto :failed

echo Building RentPaymentComparer.exe...
"%PYTHON%" -m PyInstaller --noconfirm --clean --onefile --windowed --name RentPaymentComparer --distpath "%~dp0dist" --workpath "%~dp0build\pyinstaller" --specpath "%~dp0build" "%~dp0reconcile_excel.py"
if errorlevel 1 goto :failed

echo.
echo Build complete:
echo %~dp0dist\RentPaymentComparer.exe
pause
exit /b 0

:failed
echo.
echo Build failed. Check the error messages above.
pause
exit /b 1
