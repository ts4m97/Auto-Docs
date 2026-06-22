$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $ProjectRoot

python scripts\create_app_icon.py
python -m PyInstaller --clean --noconfirm AutoDocs.spec

Write-Host ""
Write-Host "Build complete:"
Write-Host "  $ProjectRoot\dist\AutoDocs\AutoDocs.exe"
