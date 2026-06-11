@echo off
setlocal
title WorkPilot Agent
cd /d "%~dp0"

echo ============================================
echo   WorkPilot Agent  -  local on this PC
echo ============================================
echo.

where node >nul 2>nul
if errorlevel 1 (
  echo [ERROR] Node.js was not found.
  echo Install Node.js 18+ from https://nodejs.org and run this again.
  echo.
  pause
  exit /b 1
)

if not exist "node_modules" (
  echo First run: installing components... (please wait^)
  call npm install --omit=dev
  echo.
)

if not exist ".env" (
  echo First run: creating .env...
  > .env echo PORT=8787
  >> .env echo AGENT_API_KEY=workpilot-local-key
  >> .env echo ALLOWED_ORIGINS=https://sorrento.cloud,https://mohammadlaghzaoui.github.io
  >> .env echo POWERSHELL_EXE=powershell.exe
  >> .env echo PS_MODULES=ActiveDirectory,ExchangeOnlineManagement
  >> .env echo ALLOW_RAW_POWERSHELL=false
  echo.
)

echo ----------------------------------------------------------
echo   Agent will run on:   http://localhost:8787
echo   API key:             workpilot-local-key
echo.
echo   Enter these in the portal (sorrento.cloud):
echo      Settings ^> Integrations ^> Migration Agent
echo        URL     = http://localhost:8787
echo        API key = workpilot-local-key
echo.
echo   Keep this window open. Closing it stops the agent.
echo ----------------------------------------------------------
echo.

if exist "public\index.html" start "" http://localhost:8787
node src\server.js

pause
