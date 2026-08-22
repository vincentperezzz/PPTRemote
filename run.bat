@echo off
cd /d "%~dp0"
if exist "PptRemote\bin\Release\net8.0-windows\win-x64\publish\PptRemote.exe" (
  start "" "PptRemote\bin\Release\net8.0-windows\win-x64\publish\PptRemote.exe"
  exit /b 0
)
if exist "publish\PptRemote.exe" (
  start "" "publish\PptRemote.exe"
  exit /b 0
)
echo Build the app first.
pause
