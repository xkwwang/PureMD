@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist "PureMD\PureMD.exe" (
  echo ERROR: PureMD\PureMD.exe not found.
  echo Keep this .cmd file next to the "PureMD" folder!
  pause
  exit /b 1
)
"PureMD\PureMD.exe" --setup
