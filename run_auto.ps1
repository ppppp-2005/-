# paper-formatter auto run (PowerShell)
$ErrorActionPreference = "Continue"
$Root = $PSScriptRoot
Set-Location $Root

$Py = Join-Path $Root ".venv\Scripts\python.exe"

if (-not (Test-Path $Py)) {
    Write-Host "Creating venv..."
    python -m venv (Join-Path $Root ".venv")
}

if (-not (Test-Path $Py)) {
    Write-Host "ERROR: $Py not found"
    exit 1
}

Write-Host "Installing dependencies..."
& $Py -m pip install -r (Join-Path $Root "requirements.txt") -q

Write-Host "Running formatter..."
& $Py (Join-Path $Root "run_auto.py")
$code = $LASTEXITCODE

Write-Host ""
Write-Host "=== run_result.txt ==="
$result = Join-Path $Root "output\run_result.txt"
if (Test-Path $result) {
    Get-Content $result -Encoding UTF8
} else {
    Write-Host "output\run_result.txt not found"
}

if ($code -ne 0) {
    Write-Host "FAILED exit code $code"
} else {
    Write-Host "DONE. Open output\from_template*.docx in Word"
}
exit $code
