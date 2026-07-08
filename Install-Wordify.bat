title Wordify Installer

echo.
echo ========================================
echo          Wordify Installation
echo ========================================
echo.
echo Please close Microsoft Word before continuing.
echo.
pause

powershell.exe ^
  -NoProfile ^
  -ExecutionPolicy Bypass ^
  -File "%~dp0install-wordify-addin.ps1"

echo.
if errorlevel 1 (
    echo Wordify could not be installed.
    echo Please review the error message above.
) else (
    echo Wordify was installed successfully.
)

echo.
pause
