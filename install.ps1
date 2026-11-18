# unlock-file installer for Claude Code (Windows)
# Usage: powershell -ExecutionPolicy Bypass -File .\install.ps1

$ErrorActionPreference = "Stop"

$dest = Join-Path $env:USERPROFILE ".claude\skills\unlock-file"

New-Item -ItemType Directory -Force -Path $dest | Out-Null
Copy-Item -Force (Join-Path $PSScriptRoot "SKILL.md") $dest

$scriptDest = Join-Path $dest "scripts"
if (Test-Path $scriptDest) { Remove-Item -Recurse -Force $scriptDest }
New-Item -ItemType Directory -Force -Path $scriptDest | Out-Null
Copy-Item -Force (Join-Path $PSScriptRoot "scripts\unlock_file.py") $scriptDest

Write-Host "Installed to: $dest" -ForegroundColor Green

# --- dependency check ---
$py = Get-Command python -ErrorAction SilentlyContinue
if (-not $py) {
    Write-Warning "Python not found in PATH. Install Python 3.8+ first."
    exit 0
}

$ver = & python --version 2>&1
Write-Host "Found: $ver"

& python -c "import win32com" 2>$null
if ($LASTEXITCODE -ne 0) {
    Write-Host "pywin32 not installed. It is optional (COM route is more precise), install with:" -ForegroundColor Yellow
    Write-Host "  python -m pip install pywin32"
} else {
    Write-Host "pywin32: OK"
}

Write-Host ""
Write-Host "Done. Restart Claude Code, then the 'unlock-file' skill will be available." -ForegroundColor Green