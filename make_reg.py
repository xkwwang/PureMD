# -*- coding: utf-8 -*-
"""生成 set_md_preview.reg（UTF-16，可被 regedit 直接导入）。
把 .md 关联到本地纯预览 exe，并设置自定义 M 图标（DefaultIcon 指向 exe,0）。"""
import os as _os
_HERE = _os.path.dirname(_os.path.abspath(__file__))
exe = _os.path.join(_HERE, "md_preview_app", "PureMD", "PureMD.exe")
lines = [
    "Windows Registry Editor Version 5.00",
    "",
    r"[HKEY_CURRENT_USER\Software\Classes\.md]",
    '@="PureMD"',
    "",
    r"[HKEY_CURRENT_USER\Software\Classes\PureMD]",
    '@="PureMD"',
    "",
    r"[HKEY_CURRENT_USER\Software\Classes\PureMD\DefaultIcon]",
    '@="' + '"' + exe + '"' + ',0"',
    "",
    r"[HKEY_CURRENT_USER\Software\Classes\PureMD\shell\open\command]",
    '@="' + '"' + exe + '"' + ' "%1""',
    "",
    r"[HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.md\OpenWithProgids]",
    '"PureMD"=hex(0):',
    "",
    r"[-HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.md\UserChoice]",
]
text = "\r\n".join(lines) + "\r\n"
out = _os.path.join(_HERE, "set_md_preview.reg")
with open(out, "w", encoding="utf-16") as f:
    f.write(text)
print("wrote", out)
