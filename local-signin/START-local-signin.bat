@echo off
title WorkPilot local sign-in helper
echo Starting the WorkPilot local sign-in helper...
echo (Keep this window open while you connect a tenant.)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0workpilot-local-signin.ps1"
echo.
echo The helper has stopped. You can close this window.
pause
