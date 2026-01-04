@echo off
REM Get version from git tag
for /f "tokens=*" %%i in ('git describe --tags --abbrev^=0 2^>nul') do set VERSION=%%i
if "%VERSION%"=="" set VERSION=v0.0.0

echo Building AudioSwitcher %VERSION%...

REM Create VERSION file for exe
echo %VERSION%> VERSION

REM Build exe
pyinstaller --onefile --noconsole --name AudioSwitcher audio_switcher.py

REM Clean up VERSION file
del VERSION

echo.
echo Build complete: dist\AudioSwitcher.exe
pause
