@echo off
setlocal
title WorkPilot Agent
cd /d "%~dp0"

echo ============================================
echo   WorkPilot Agent  -  lokaal op deze PC
echo ============================================
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo [FOUT] Node.js is niet gevonden.
  echo Installeer Node.js 18+ van https://nodejs.org en start dit opnieuw.
  echo.
  pause
  exit /b 1
)

if not exist "node_modules" (
  echo Eenmalig: onderdelen installeren... (even wachten^)
  call npm install --omit=dev
  echo.
)

if not exist ".env" (
  echo Eerste keer: .env aanmaken...
  > .env echo PORT=8787
  >> .env echo AGENT_API_KEY=workpilot-local-key
  >> .env echo ALLOWED_ORIGINS=https://sorrento.cloud,https://mohammadlaghzaoui.github.io
  >> .env echo POWERSHELL_EXE=powershell.exe
  >> .env echo PS_MODULES=ActiveDirectory,ExchangeOnlineManagement
  >> .env echo ALLOW_RAW_POWERSHELL=false
  echo.
)

echo ----------------------------------------------------------
echo   Agent draait straks op:  http://localhost:8787
echo   API-sleutel:             workpilot-local-key
echo.
echo   Vul deze in op de portal (sorrento.cloud):
echo      Settings ^> Integrations ^> Migration Agent
echo        URL     = http://localhost:8787
echo        API key = workpilot-local-key
echo.
echo   Laat dit venster open. Sluiten = agent stoppen.
echo ----------------------------------------------------------
echo.

if exist "public\index.html" start "" http://localhost:8787
node src\server.js

pause
