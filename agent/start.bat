@echo off
setlocal
title WorkPilot - lokale server
cd /d "%~dp0"

echo ============================================
echo   WorkPilot  -  lokale server starten
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
  echo Eenmalig: dependencies installeren...
  call npm install --omit=dev
  echo.
)

if not exist ".env" (
  echo Eerste keer: .env aanmaken met een lokale API-sleutel...
  > .env echo PORT=8787
  >> .env echo AGENT_API_KEY=workpilot-local-key
  >> .env echo ALLOWED_ORIGINS=
  >> .env echo POWERSHELL_EXE=powershell.exe
  >> .env echo PS_MODULES=ActiveDirectory,ExchangeOnlineManagement
  >> .env echo ALLOW_RAW_POWERSHELL=false
  echo   API-sleutel = workpilot-local-key  (pas aan in .env indien gewenst)
  echo.
)

echo Server start op http://localhost:8787
echo De browser opent zo automatisch. Sluit dit venster om te stoppen.
echo.
start "" http://localhost:8787
node src\server.js

pause
