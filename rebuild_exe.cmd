@echo off
cd /d "%~dp0"
set PYI=C:\Users\15623\.workbuddy\binaries\python\envs\default\Scripts\pyinstaller.exe

echo === PureMD: rebuild exe ===
echo (Close the preview window BEFORE running this)
echo.

if exist "md_preview_app\PureMD" (
  echo Moving old build aside...
  move "md_preview_app\PureMD" "md_preview_app\_old_%RANDOM%"
  if errorlevel 1 (
    echo FAILED to move old build.
    echo The preview window is probably still open - close it and run again.
    pause
    exit /b 1
  )
)

echo Building... (this takes 1-3 minutes, please wait)
"%PYI%" --noconsole --onedir --name PureMD --icon mdpreview\md_preview.ico --add-data "mdpreview;mdpreview" --collect-data webview --collect-submodules webview --hidden-import webview.platforms.edgechromium --collect-data docx --collect-submodules markdown_it --collect-submodules docx --distpath md_preview_app open_md_preview.pyw
if errorlevel 1 (
  echo BUILD FAILED
  pause
  exit /b 1
)

echo Self test...
md_preview_app\PureMD\PureMD.exe --selftest
echo Result:
type "%TEMP%\mdpreview_selftest.txt"
echo.
echo DONE. You can now double-click any .md file.
pause
