@echo off
cd /d "%~dp0"
echo Starting Otomasi Report...
call npm run electron:dev
pause
