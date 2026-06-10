# =====================================================================
#  WorkPilot All-in-One Server installer (Windows Server, domain-joined)
#  Installs the portal + Migration Agent as a single Windows service.
#  Run as Administrator from the unzipped package root.
# =====================================================================

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$agent = Join-Path $root 'agent'

Write-Host "`n=== WorkPilot All-in-One Server installer ===" -ForegroundColor Cyan

# 1) Node.js check
$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) {
  Write-Host "Node.js not found. Installing via winget..." -ForegroundColor Yellow
  winget install OpenJS.NodeJS.LTS --accept-package-agreements --accept-source-agreements
  $env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path','User')
}
Write-Host "Node.js: $(node --version)"

# 2) Install agent dependencies
Push-Location $agent
if (-not (Test-Path '.env')) {
  Copy-Item '.env.example' '.env'
  # Generate a random API key into the fresh .env
  $key = -join ((48..57)+(97..122) | Get-Random -Count 48 | ForEach-Object {[char]$_})
  (Get-Content '.env') -replace 'AGENT_API_KEY=.*', "AGENT_API_KEY=$key" | Set-Content '.env'
  Write-Host "Generated agent API key (also enter this in the portal Settings): $key" -ForegroundColor Green
}
npm install --omit=dev
Pop-Location

# 3) Verify the portal UI is present (agent/public)
if (-not (Test-Path (Join-Path $agent 'public\index.html'))) {
  Write-Host "agent\public\index.html missing — this package should include the built portal. (Repo dev: npm run build:server)" -ForegroundColor Yellow
}

# 4) Register as a Windows service (NSSM if available, else Scheduled Task)
$serviceName = 'WorkPilotServer'
$nssm = Get-Command nssm -ErrorAction SilentlyContinue
if ($nssm) {
  & nssm stop $serviceName 2>$