# unlock-file installer for Agent Skills clients (Windows)
# Default target: Claude Code user skills directory.
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File .\install.ps1
#   powershell -ExecutionPolicy Bypass -File .\install.ps1 -Dest "C:\path\to\skills\unlock-file"

param(
    [string]$Dest = (Join-Path $env:USERPROFILE ".claude\skills\unlock-file")
)

$ErrorActionPreference = "Stop"

$Dest = $Dest.TrimEnd('\', '/')
$skillSrc = $PSScriptRoot

if (-not (Test-Path (Join-Path $skillSrc "SKILL.md"))) {
    throw "Cannot find SKILL.md next to install.ps1."
}
if ((Split-Path -Leaf $Dest) -ne "unlock-file") {
    throw "Destination folder must be named 'unlock-file' (it must match the skill name). Got: $Dest"
}

New-Item -ItemType Directory -Force -Path $Dest | Out-Null
Copy-Item -Force (Join-Path $skillSrc "SKILL.md") $Dest

$scriptDest = Join-Path $Dest "scripts"
if (Test-Path $scriptDest) { Remove-Item -Recurse -Force $scriptDest }
New-Item -ItemType Directory -Force -Path $scriptDest | Out-Null
Copy-Item -Force (Join-Path $skillSrc "scripts\unlock_file.py") $scriptDest

Write-Host "Installed to: $Dest" -ForegroundColor Green

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
Write-Host "Done. Restart your agent client (for Claude Code, reopen the session), then the 'unlock-file' skill will be available." -ForegroundColor Green