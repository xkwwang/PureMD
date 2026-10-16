@echo off
chcp 65001 >nul
reg delete "HKCU\Software\Classes\.md" /f >nul 2>&1
reg delete "HKCU\Software\Classes\PureMD" /f >nul 2>&1
reg delete "HKCU\Software\Classes\MarkdownPreviewMD" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.md\UserChoice" /f >nul 2>&1
reg delete "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.md\OpenWithProgids" /f >nul 2>&1
echo Done. The .md association has been removed.
pause
