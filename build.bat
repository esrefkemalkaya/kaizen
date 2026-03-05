@echo off
cd /d "%~dp0"

echo Installing / upgrading PyInstaller...
python -m pip install --upgrade pyinstaller

echo.
echo Building DrillInvoice.exe ...
python -m PyInstaller drill_invoice.spec --clean

echo.
if exist "dist\DrillInvoice.exe" (
    echo BUILD SUCCESSFUL
    echo Executable: %~dp0dist\DrillInvoice.exe
) else (
    echo BUILD FAILED - check output above for errors
)
pause
