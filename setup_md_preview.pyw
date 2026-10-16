# -*- coding: utf-8 -*-
"""一键设置：把 .md 关联到本地纯预览 exe（带自定义 M 图标）。双击 setup_md_preview.cmd 触发。
相比旧版：更彻底地清除“打开方式”留下的 UserChoice（含递归删除子项），
并写入 OpenWithProgids，确保我们的 ProgID 胜出。"""
import winreg, ctypes, os, traceback

_HERE = os.path.dirname(os.path.abspath(__file__))
EXE = os.path.join(_HERE, "md_preview_app", "PureMD", "PureMD.exe")
CMD = '"{}" "%1"'.format(EXE)
ICON = '"{}",0'.format(EXE)
LOG = os.path.join(_HERE, "setup_log.txt")


def log(msg):
    try:
        with open(LOG, "a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass


def msg(text, err=False):
    ctypes.windll.user32.MessageBoxW(0, str(text), "PureMD 设置", 0x10 if err else 0x40)


def del_key_recursive(root, path):
    """删除注册表项；若有子项则递归。返回是否“已不存在/已删”。"""
    try:
        winreg.DeleteKey(root, path)  # 仅当无子项时成功
        return True
    except FileNotFoundError:
        return True
    except Exception:
        pass
    try:
        h = winreg.OpenKey(root, path, 0, winreg.KEY_READ | winreg.KEY_WRITE)
    except FileNotFoundError:
        return True
    except Exception as e:
        log("open failed %s: %s" % (path, e))
        return False
    try:
        subs = []
        i = 0
        while True:
            try:
                name = winreg.EnumKey(h, i)
            except OSError:
                break
            subs.append(name)
            i += 1
        for s in subs:
            del_key_recursive(root, path + "\\" + s)
        winreg.CloseKey(h)
        winreg.DeleteKey(root, path)
        return True
    except Exception as e:
        log("recursive del failed %s: %s" % (path, e))
        return False


try:
    # 1) 经典关联：.md -> PureMD
    k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\.md")
    winreg.SetValue(k, "", winreg.REG_SZ, "PureMD")
    winreg.CloseKey(k)

    # 2) ProgID：名称 / 图标 / 打开命令
    k2 = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\PureMD")
    winreg.SetValue(k2, "", winreg.REG_SZ, "PureMD")
    kicon = winreg.CreateKey(k2, "DefaultIcon")
    winreg.SetValue(kicon, "", winreg.REG_SZ, ICON)
    winreg.CloseKey(kicon)
    k3 = winreg.CreateKey(k2, r"shell\open\command")
    winreg.SetValue(k3, "", winreg.REG_SZ, CMD)
    winreg.CloseKey(k3)
    winreg.CloseKey(k2)

    # 3) 清除“打开方式”留下的 UserChoice（如 RStudio），让 ProgID 生效
    ok = del_key_recursive(winreg.HKEY_CURRENT_USER,
                           r"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.md\UserChoice")
    log("UserChoice cleared: %s" % ok)

    # 4) 写入 OpenWithProgids，帮助外壳优先选择我们的程序
    try:
        kp = winreg.CreateKey(winreg.HKEY_CURRENT_USER,
                              r"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.md\OpenWithProgids")
        winreg.SetValueEx(kp, "PureMD", 0, winreg.REG_NONE, b"")
        winreg.CloseKey(kp)
    except Exception as e:
        log("OpenWithProgids: %s" % e)

    # 5) 通知外壳刷新关联（图标/默认程序立即生效，不必注销）
    try:
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)  # SHCNE_ASSOCCHANGED
    except Exception:
        pass

    msg("设置成功！\n\n现在双击任意 .md 会用本地纯预览窗口打开，且文件显示自带的 M 图标：\n"
        "· 源文件不会被改动\n· 不生成任何额外文件\n· 多标签统一在一个窗口\n\n"
        "若双击仍走别的程序，请按弹窗后的说明用 PowerShell 强制修复；\n"
        "若图标未立刻变化，重启“Windows 资源管理器”即可。\n\n命令：\n" + CMD)
except Exception as e:
    log(traceback.format_exc())
    msg("设置失败：\n" + str(e), err=True)
