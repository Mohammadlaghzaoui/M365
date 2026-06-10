<#
  WorkPilot all-in-one server — Windows installer.

  Installs the portal + agent as a Windows service on a domain-joined server,
  so it can create AD users, run Exchange migration batches and serve the UI
  from one machine.

  Run in an elevated PowerShell:
    Set-ExecutionPolicy -Scope Process Bypass
    .\install-windows.ps1 -ApiKey "your-long-random-key" -Port 8787

  Optional Microsoft Graph app-only (for cloud user creation / invitations):
    .\install-windows.ps1 -ApiKey "..." -TenantId "..." -ClientId "..." -ClientSecret "..."
#>
param(
  [Parameter(Mandatory = $true)][string]$ApiKey,
  [int]$Port = 8787,
  [string]$TenantId = "",
  [string]$ClientId = "",
  [string]$ClientSecret = "",
  [string]$AllowedOrigins = "",
  [string]$InstallDir = "C:\WorkPilot"
)

$ErrorActionPreference = "Stop"
Write-Host "== WorkPilot all-in-one server installer ==" -ForegroundColor Cyan

# 1. Node.js check
if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
  throw "Node.js 18+ is required. Install it from https://nodejs.org and re-run."
}
Write-Host "Node.js: $(node --version)"

# 2. Copy project to InstallDir (run this script from the extracted repo root)
$repoRoot = Split-Path -Parent $PSScriptRoot
Write-Host "Building portal + installing dependencies (this can take a few minutes)..."
Push-Location $repoRoot
npm install
npm run build:server
Pop-Location

if (-not (Test-Path $InstallDir)) { New-Item -ItemType Directory -Path $InstallDir | Out-Null }
Copy-Item "$repoRoot\agent\*" $InstallDir -Recurse -Force
Push-Location $InstallDir
npm install --omit=dev
Pop-Location

# 3. Write .env
$envText = @"
PORT=$Port
AGENT_API_KEY=$ApiKey
ALLOWED_ORIGINS=$AllowedOrigins
TENANT_ID=$TenantId
CLIENT_ID=$ClientId
CLIENT_SECRET=$ClientSecret
POWERSHELL_EXE=powershell.exe
PS_MODULES=ActiveDirectory,ExchangeOnlineManagement
ALLOW_RAW_POWERSHELL=false
"@
Set-Content -Path "$InstallDir\.env" -Value $envText -Encoding UTF8
Write-Host ".env written to $InstallDir\.env"

# 4. Install as a Windows service via NSSM (downloads if missing)
$nssm = "$InstallDir\nssm.exe"
if (-not (Test-Path $nssm)) {
  Write-Host "Downloading NSSM (service wrapper)..."
  $zip = "$env:TEMP\nssm.zip"
  Invoke-WebRequest "https://nssm.cc/release/nssm-2.24.zip" -OutFile $zip
  Expand-Archive $zip "$env:TEMP\nssm" -Force
  Copy-Item "$env:TEMP\nssm\nssm-2.24\win64\nssm.exe" $nssm -Force
}

$nodePath = (Get-Command node).Source
& $nssm stop WorkPilot 2>$null
& $nssm remove WorkPilot confirm 2>$null
& $nssm install WorkPilot $nodePath "$InstallDir\src\server.js"
& $nssm set WorkPilot AppDirectory $InstallDir
& $nssm set WorkPilot AppEnvironmentExtra "DOTENV=1"
& $nssm set WorkPilot Start SERVICE_AUTO_START
& $nssm start WorkPilot

# 5. Firewall rule
New-NetFirewallRule -DisplayName "WorkPilot $Port" -Direction Inbound -LocalPort $Port -Protocol TCP -Action Allow -ErrorAction SilentlyContinue | Out-Null

Write-Host ""
Write-Host "Done! WorkPilot is running as a Windows service." -ForegroundColor Green
Write-Host "  Portal + API:  http://$($env:COMPUTERNAME):$Port" -ForegroundColor Green
Write-Host "  Agent API key: $ApiKey"
Write-Host ""
Write-Host "Next: open the portal, log in, and in Settings > Integrations > Migration Agent"
Write-Host "set URL = http://localhost:$Port (same origin) and the API key above."
Write-Host "Put it behind HTTPS (IIS/ARR or a reverse proxy) before exposing it."
