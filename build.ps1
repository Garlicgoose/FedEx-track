# FedExMailAlert build script (2026-09-01)
# Usage:  powershell -ExecutionPolicy Bypass -File build.ps1
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

python -m PyInstaller --noconfirm --clean --windowed `
    --name "FedExMailAlert" `
    main.py

Write-Host "DONE: $root\dist\FedExMailAlert"
