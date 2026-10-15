@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

REM Clear ELECTRON_RUN_AS_NODE: if set, Electron runs as plain Node
REM and crashes with "Cannot read properties of undefined (reading 'isPackaged')".
set "ELECTRON_RUN_AS_NODE="

title WorkSpace - Development

echo.
echo  ============================================
echo   WorkSpace - Development mode
echo  ============================================
echo.

where node >nul 2>nul
if errorlevel 1 goto no_node

if not exist "node_modules" goto install_deps
goto check_electron

:install_deps
echo  Dependencies missing. Running npm install...
echo.
call npm install
if errorlevel 1 goto install_failed

:check_electron
if exist "node_modules\electron\dist\electron.exe" goto run
echo  Electron binary missing. Downloading...
call node node_modules\electron\install.js
if errorlevel 1 goto install_failed

:run
echo  Starting Vite dev server and Electron...
echo  Keep this window open. Press Ctrl+C or close it to stop.
echo.
call npm run dev
echo.
echo  Application stopped.
pause
endlocal
exit /b 0

:no_node
echo  [ERROR] Node.js not found in PATH.
echo  Install Node.js 20+ from https://nodejs.org
echo.
pause
endlocal
exit /b 1

:install_failed
echo.
echo  [ERROR] Installation failed. See output above.
echo.
pause
endlocal
exit /b 1