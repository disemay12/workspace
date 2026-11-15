@echo off
REM chcp 65001 switched the console codepage and never restored it, so the
REM caller's console stayed in UTF-8 after the script. Remember the previous
REM codepage and put it back before exit.
set "OLDCP="
for /f "tokens=2 delims=:" %%c in ('chcp') do set "OLDCP=%%c"
set "OLDCP=%OLDCP: =%"
if not defined OLDCP set "OLDCP=65001"
chcp 65001 >nul
setlocal
cd /d "%~dp0"
title WorkSpace - Diagnostics
REM Usage: check-env.bat
REM          - interactive run, waits for a keypress at the end
REM        check-env.bat /nopause
REM          - script/CI run, never pauses

echo.
echo  ============================================
echo   WorkSpace - Diagnostics
echo  ============================================
echo.

set "PROBLEMS=0"
REM Warnings do not affect the exit code: the diagnostics stay usable in CI,
REM where ELECTRON_RUN_AS_NODE is set by the environment and is not a defect.
set "WARNINGS=0"

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
set /a WARNINGS+=1
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
if not "%WARNINGS%"=="0" goto has_warnings
echo   No problems found - run start-dev.bat
goto summary_done
:has_problems
echo   Problems found: %PROBLEMS% (warnings: %WARNINGS%)
goto summary_done
:has_warnings
echo   No problems found, warnings: %WARNINGS%
:summary_done
echo  --------------------------------------------
echo.

REM Exit code reflects the result: 1 when problems were found, otherwise 0.
REM pause only for interactive runs: no arguments. With /nopause (or any other
REM argument) the script does not wait for a key and is safe for CI.
set "EXIT_CODE=0"
if not "%PROBLEMS%"=="0" set "EXIT_CODE=1"
if "%~1"=="" pause
REM OLDCP was set before setlocal, so it survives endlocal. The line is expanded
REM with the current values first, then the codepage of the caller is restored.
endlocal & chcp %OLDCP% >nul & exit /b %EXIT_CODE%
