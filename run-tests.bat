@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

REM Clear ELECTRON_RUN_AS_NODE - the checks launch a real Electron process.
set "ELECTRON_RUN_AS_NODE="
REM Clear the dev server URL so production mode loads built files from dist/.
set "VITE_DEV_SERVER_URL="

title WorkSpace - Checks

echo.
echo  ============================================
echo   WorkSpace - Test suite
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
if exist "node_modules\electron\dist\electron.exe" goto build_checks
echo  Electron binary missing. Downloading...
call node node_modules\electron\install.js
if errorlevel 1 goto install_failed

:build_checks
echo  Compiling check scripts...
call npm run checks:build
if errorlevel 1 goto compile_failed

echo  Running suites (may take up to 2 minutes)...
echo.
node dist-electron\run-all-checks.js
set "CHECKS_EXIT=%ERRORLEVEL%"

echo.
if not "%CHECKS_EXIT%"=="0" goto has_failures
echo  ============================================
echo   ALL CHECKS PASSED
echo  ============================================
goto report

:has_failures
echo  ============================================
echo   SOME CHECKS FAILED - see output above
echo  ============================================

:report
if not exist "final-report.txt" goto finish
echo.
echo  Detailed report: final-report.txt

:finish
echo.
pause
endlocal
exit /b %CHECKS_EXIT%

:no_node
echo  [ERROR] Node.js not found in PATH.
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

:compile_failed
echo.
echo  [ERROR] Could not compile check scripts.
echo.
pause
endlocal
exit /b 1