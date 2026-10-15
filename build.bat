@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

set "ELECTRON_RUN_AS_NODE="
REM Clear the dev server URL so production mode loads built files from dist/.
set "VITE_DEV_SERVER_URL="

title WorkSpace - Build

echo.
echo  ============================================
echo   WorkSpace - Build
echo  ============================================
echo.

where node >nul 2>nul
if errorlevel 1 goto no_node

if not exist "node_modules" goto install_deps
goto do_build

:install_deps
echo  Dependencies missing. Running npm install...
echo.
call npm install
if errorlevel 1 goto install_failed

:do_build
call npm run build:portable
if errorlevel 1 goto build_failed

echo.
echo  ============================================
echo   Build completed successfully
echo  ============================================
echo.
echo   Frontend:        dist\
echo   Main process:    dist-electron\
echo   Portable app:    WorkSpace\WorkSpace.exe
echo.
echo   Run with: WorkSpace\WorkSpace.exe
echo   Or:       start-prod.bat
echo   Or:       npx electron dist-electron/main.js
echo.
pause
endlocal
exit /b 0

:no_node
echo  [ERROR] Node.js not found in PATH.
echo.
pause
endlocal
exit /b 1

:install_failed
echo.
echo  [ERROR] npm install failed.
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