# Build and run the verification suites for the WorkSpace fixes.
# Compiles the app, then runs every check that does not need a GUI window.
# Messages are ASCII-only so Windows PowerShell 5.1 reads the file correctly.

$ErrorActionPreference = 'Continue'
$root = $PSScriptRoot
Set-Location $root

$tsc = Join-Path $root 'node_modules\.bin\tsc.cmd'
$vite = Join-Path $root 'node_modules\.bin\vite.cmd'
$verify = Join-Path $root 'dist-verify'

Write-Host ''
Write-Host ('=' * 72)
Write-Host '  WORKSPACE FIX VERIFICATION'
Write-Host ('=' * 72)

$failures = @()

# --- Type checks ---
Write-Host ''
Write-Host '[1/5] Typecheck frontend...'
& $tsc -p tsconfig.json --noEmit 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) { $failures += 'typecheck-frontend' } else { Write-Host '      OK' }

Write-Host '[2/5] Typecheck electron...'
& $tsc -p tsconfig.electron.json --noEmit 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) { $failures += 'typecheck-electron' } else { Write-Host '      OK' }

# --- Build ---
Write-Host '[3/5] Build frontend and main process...'
& $vite build 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) { $failures += 'vite-build' } else { Write-Host '      frontend OK' }
& $tsc -p tsconfig.electron.json 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) { $failures += 'electron-build' } else { Write-Host '      main OK' }

# --- Compile check suites ---
Write-Host '[4/5] Compile check suites...'
Remove-Item -Recurse -Force $verify -ErrorAction SilentlyContinue
& $tsc --ignoreConfig `
  src/lib/formulaEngine.ts `
  src/components/editors/useFormulaMode.ts `
  electron/checks/formula-check.ts `
  electron/checks/formula-mode-check.ts `
  electron/checks/exit-confirm-check.ts `
  electron/checks/xlsx-save-check.ts `
  electron/checks/ai-formula-check.ts `
  --outDir $verify --rootDir . --module commonjs --target ES2022 `
  --esModuleInterop --skipLibCheck --strict false --jsx react-jsx --types node 2>&1 | Out-Null
if ($LASTEXITCODE -ne 0) {
  $failures += 'compile-checks'
  Write-Host '      COMPILE ERROR'
} else {
  Write-Host '      OK'
}

# --- Run suites ---
Write-Host '[5/5] Run check suites...'
Write-Host ''

$suites = @(
  @{ name = 'Formula engine (108 functions)'; file = 'formula-check.js' },
  @{ name = 'Formula mode (autocomplete, click-to-ref)'; file = 'formula-mode-check.js' },
  @{ name = 'AI formula -> cell round trip'; file = 'ai-formula-check.js' },
  @{ name = 'Exit confirmation'; file = 'exit-confirm-check.js' },
  @{ name = 'XLSX formula persistence'; file = 'xlsx-save-check.js' }
)

foreach ($suite in $suites) {
  $script = Join-Path $verify "electron\checks\$($suite.file)"
  if (-not (Test-Path $script)) {
    $failures += $suite.name
    Write-Host "  [MISS] $($suite.name) - script not built"
    continue
  }
  $output = & node $script 2>&1
  $code = $LASTEXITCODE
  $counts = ($output | Select-String -Pattern 'Passed|projd|Projdeno' | ForEach-Object { $_.Line.Trim() }) -join ' '
  if ($code -eq 0) {
    Write-Host "  [PASS] $($suite.name)"
  } else {
    $failures += $suite.name
    Write-Host "  [FAIL] $($suite.name)"
    $output | Select-String -Pattern '^\s+-' | Select-Object -First 8 | ForEach-Object {
      Write-Host "         $($_.Line.Trim())"
    }
  }
}

# --- Summary ---
Write-Host ''
Write-Host ('-' * 72)
if ($failures.Count -eq 0) {
  Write-Host '  ALL CHECKS PASSED'
  exit 0
} else {
  Write-Host "  FAILED: $($failures -join ', ')"
  exit 1
}
