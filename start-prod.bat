@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

REM Clear ELECTRON_RUN_AS_NODE: if set, Electron runs as plain Node and crashes.
set "ELECTRON_RUN_AS_NODE="
REM Clear the dev server URL so production mode loads built files from dist/.
set "VITE_DEV_SERVER_URL="

title WorkSpace - Production

echo.
echo  ============================================
echo   WorkSpace - Build and Run (production)
echo  ============================================
echo.

where node >nul 2>nul
if errorlevel 1 goto no_node

if not exist "node_modules" goto install_deps
goto step_build

:install_deps
echo  Dependencies missing. Running npm install...
echo.
call npm install
if errorlevel 1 goto install_failed

:step_build
echo  Step 1 of 2: building project...
call npm run build
if errorlevel 1 goto build_failed

echo.
echo  Step 2 of 2: launching application...
echo.
REM Launch through scripts/run-electron.js: it adds --no-sandbox only when the
REM Chromium sandbox cannot work on this machine (UAC disabled, EnableLUA=0).
call node scripts\run-electron.js dist-electron/main.js

echo.
echo  Application closed.
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
echo  [ERROR] npm install failed. See output above.
echo.
pause
endlocal
exit /b 1

:build_failed
echo.
echo  [ERROR] Build failed. See output above.
echo.
pause
endlocal
exit /b 1