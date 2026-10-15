@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title WorkSpace - Diagnostics

echo.
echo  ============================================
echo   WorkSpace - Diagnostics
echo  ============================================
echo.

set "PROBLEMS=0"

REM ---------- Node.js ----------
set "NODE_VER="
for /f "tokens=*" %%v in ('node --version 2^>nul') do set "NODE_VER=%%v"
if defined NODE_VER goto node_ok
echo  [FAIL] Node.js not found in PATH
echo         Install Node.js 20+ from https://nodejs.org
set /a PROBLEMS+=1
goto node_done
:node_ok
echo  [OK]   Node.js: %NODE_VER%
:node_done

REM ---------- Dependencies ----------
if not exist "node_modules" goto no_deps
echo  [OK]   node_modules present
goto deps_done
:no_deps
echo  [FAIL] node_modules missing - run: npm install
set /a PROBLEMS+=1
:deps_done

REM ---------- Electron binary ----------
if not exist "node_modules\electron\dist\electron.exe" goto no_electron
echo  [OK]   Electron binary present
goto electron_done
:no_electron
echo  [FAIL] Electron binary not downloaded
echo         Fix: node node_modules\electron\install.js
set /a PROBLEMS+=1
:electron_done

REM ---------- Blocking environment variable ----------
if not defined ELECTRON_RUN_AS_NODE goto no_envvar
echo  [WARN] ELECTRON_RUN_AS_NODE is set - it blocks Electron startup
echo         The .bat launchers clear it automatically.
echo         For manual runs use: set "ELECTRON_RUN_AS_NODE="
set /a PROBLEMS+=1
goto envvar_done
:no_envvar
echo  [OK]   ELECTRON_RUN_AS_NODE not set
:envvar_done

REM ---------- Secrets file ----------
if not exist "secrets.local.ps1" goto no_secrets
echo  [OK]   secrets.local.ps1 found - import key from Settings
goto secrets_done
:no_secrets
echo  [INFO] secrets.local.ps1 not found - enter key manually in Settings
:secrets_done

REM ---------- Build artifacts ----------
if not exist "dist\index.html" goto no_build
echo  [OK]   Frontend build: dist\index.html
goto fe_done
:no_build
echo  [INFO] Frontend not built - run: npm run build
:fe_done

if not exist "dist-electron\main.js" goto no_main
echo  [OK]   Main process: dist-electron\main.js
goto main_done
:no_main
echo  [INFO] Main process not built - run: npm run build
:main_done

REM ---------- Summary ----------
echo.
echo  --------------------------------------------
if not "%PROBLEMS%"=="0" goto has_problems
echo   No problems found - run start-dev.bat
goto summary_done
:has_problems
echo   Problems found: %PROBLEMS%
:summary_done
echo  --------------------------------------------
echo.
pause
endlocal