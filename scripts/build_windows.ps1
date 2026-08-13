$ErrorActionPreference = "Stop"

$ProjectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$VirtualEnvironment = Join-Path $ProjectRoot ".venv-build"
$Python = Join-Path $VirtualEnvironment "Scripts\python.exe"

Set-Location $ProjectRoot

if (-not (Test-Path $Python)) {
    py -m venv $VirtualEnvironment
}

& $Python -m pip install --upgrade pip
& $Python -m pip install -r (Join-Path $ProjectRoot "requirements-dev.txt")
& $Python -m unittest discover -s tests -v
& $Python -m PyInstaller --noconfirm --clean (Join-Path $ProjectRoot "build\sop_reporter_debug.spec")
& $Python -m PyInstaller --noconfirm --clean (Join-Path $ProjectRoot "build\sop_reporter.spec")

Write-Host ""
Write-Host "Build complete:"
Write-Host "  $ProjectRoot\dist\SOPReporter-Debug.exe"
Write-Host "  $ProjectRoot\dist\SOPReporter.exe"

