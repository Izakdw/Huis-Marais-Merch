@echo off
title Huis Marais Public Merch Tunnel
echo ===================================================
echo   Starting Cloudflare Tunnel for Residents...
echo ===================================================
cd /d "%~dp0"
cloudflared.exe tunnel --url http://localhost:3000
pause
