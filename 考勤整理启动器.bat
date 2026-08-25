@echo off
setlocal
rem ASCII-only batch: UTF-8 Chinese in .bat breaks cmd.exe parsing on Chinese Windows.
cd /d "%~dp0"

set "GUI=%~dp0gui.py"
if not exist "%GUI%" (
  echo [ERROR] gui.py not found:
  echo   %GUI%
  echo.
  pause
  exit /b 1
)

rem Prefer pythonw (no black console). Fall back to python / py launcher.
where pythonw >nul 2>nul
if %ERRORLEVEL% equ 0 (
  start "QingQuan-KaoQin" /D "%~dp0" pythonw "%GUI%"
  exit /b 0
)

where python >nul 2>nul
if %ERRORLEVEL% equ 0 (
  start "QingQuan-KaoQin" /D "%~dp0" python "%GUI%"
  exit /b 0
)

where py >nul 2>nul
if %ERRORLEVEL% equ 0 (
  start "QingQuan-KaoQin" /D "%~dp0" py -3 "%GUI%"
  exit /b 0
)

echo [ERROR] Python not found.
echo Install Python 3 and enable "Add python.exe to PATH".
echo.
pause
exit /b 1
