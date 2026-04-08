param(
    [string]$Python = "py"
)

$ErrorActionPreference = "Stop"

Set-Location -Path $PSScriptRoot

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    & $Python -m venv .venv
}

. ".\.venv\Scripts\Activate.ps1"

python -m pip install --upgrade pip
python -m pip install -r requirements.txt pyinstaller

pyinstaller `
  --noconfirm `
  --clean `
  --windowed `
  --name TrialBalanceApp `
  --add-data "config.yaml;." `
  --add-data "templates/Trial Balance_template.xlsx;templates" `
  desktop_app.py

Write-Host ""
Write-Host "Build done."
Write-Host "EXE: dist\TrialBalanceApp\TrialBalanceApp.exe"
