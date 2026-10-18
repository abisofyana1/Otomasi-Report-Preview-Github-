@echo off
cd /d "%~dp0"
echo Setting local cache to avoid corruption issues...
mkdir .cache 2>nul
set ELECTRON_BUILDER_CACHE=%cd%\.cache

echo Cleaning old build output...
if exist "installer_output" rmdir /s /q "installer_output"

echo Building portable app...
echo This process will download Electron files to the .cache folder.
echo Please ensure you have a stable internet connection.
call npm run electron:build
if %errorlevel% neq 0 (
    echo Build failed!
    echo If you see "zip: not a valid zip file", please try deleting the .cache folder and running this again.
    pause
    exit /b %errorlevel%
)
echo Build success!
echo You can find the portable exe in the "installer_output" folder.
pause
