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
REM 走 PureMD.spec（精确的 datas/hiddenimports/excludes 清单），
REM 不要再用宽泛的 --collect-submodules，否则体积会膨胀回 35MB。
"%PYI%" --noconfirm --distpath md_preview_app PureMD.spec
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
