@echo off
title Merch Management Portal
echo ===================================================
echo   Starting Merch Management Portal...
echo ===================================================
cd /d "%~dp0"

:: Automatically open the website in your default browser
start "" http://localhost:3000

:: Run the server
node server.js
pause
