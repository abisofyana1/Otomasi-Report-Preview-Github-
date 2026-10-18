@echo off
cd /d "%~dp0"
echo Building installer...
echo This process may take a few minutes.
call npm run electron:build
if %errorlevel% neq 0 (
    echo Build failed!
    pause
    exit /b %errorlevel%
)
echo Build success!
echo You can find the installer in the "release" folder.
pause
