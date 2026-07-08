@echo off
setlocal

title Wordify Installer

set "INSTALLER=%~dp0install-wordify-addin.ps1"
if not exist "%INSTALLER%" (
    set "INSTALLER=%~dp0scripts\install-wordify-addin.ps1"
)

set "LOCAL_VBA=%~dp0vba"
set "LOCAL_MANIFEST=%~dp0manifest.xml"
set "INSTALL_ARGS="

echo.
echo ========================================
echo          Wordify Installation
echo ========================================
echo.
echo Please close Microsoft Word before continuing.
echo.
pause

if not exist "%INSTALLER%" (
    echo.
    echo Wordify installer script was not found:
    echo %INSTALLER%
    echo.
    pause
    exit /b 1
)

if exist "%LOCAL_VBA%\" (
    if exist "%LOCAL_MANIFEST%" (
        set "INSTALL_ARGS=-SkipSourceBootstrap"
    )
)

powershell.exe ^
  -NoProfile ^
  -ExecutionPolicy Bypass ^
  -File "%INSTALLER%" ^
  %INSTALL_ARGS%

set "INSTALL_EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%INSTALL_EXIT_CODE%"=="0" (
    echo Wordify could not be installed.
    echo Please review the error message above.
) else (
    echo Wordify was installed successfully.
)

echo.
pause
exit /b %INSTALL_EXIT_CODE%
