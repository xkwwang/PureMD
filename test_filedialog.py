# -*- coding: utf-8 -*-
r"""自动测试文件选择框：弹窗后用 Win32 消息把它关掉，验证不挂死、不抛异常。

用法：python test_filedialog.py [comdlg|ps|both]
"""
import sys, time, threading, subprocess, importlib.machinery, importlib.util, ctypes
from ctypes import wintypes

PYW = r"D:\00. APP\PureMD\open_md_preview.pyw"
WM_CLOSE = 0x0010
TITLE = "选择 Markdown 文件"


def load_module():
    loader = importlib.machinery.SourceFileLoader("omp", PYW)
    spec = importlib.util.spec_from_loader("omp", loader)
    m = importlib.util.module_from_spec(spec)
    loader.exec_module(m)
    return m


def find_and_close(timeout=8.0):
    """轮询找到标题为 TITLE 的对话框并关闭它，返回是否找到。"""
    u = ctypes.windll.user32
    deadline = time.time() + timeout
    while time.time() < deadline:
        hwnd = u.FindWindowW("#32770", TITLE)
        if not hwnd:
            hwnd = u.FindWindowW(None, TITLE)
        if hwnd:
            time.sleep(0.4)          # 等对话框完成初始化
            u.PostMessageW(hwnd, WM_CLOSE, 0, 0)
            return True
        time.sleep(0.2)
    return False


def run_comdlg(m):
    box = {}
    def target():
        try:
            box["r"] = m._open_dialog_comdlg()
            box["e"] = None
        except Exception as e:
            box["e"] = repr(e)
    t = threading.Thread(target=target, daemon=True)
    t.start()
    found = find_and_close()
    t.join(15)
    err = ctypes.windll.comdlg32.CommDlgExtendedError()
    print("[comdlg32] dialog_found=%s  returned=%r  exception=%r  CommDlgError=%d  thread_alive=%s"
          % (found, box.get("r"), box.get("e"), err, t.is_alive()))
    return found


def run_ps(m):
    box = {}
    def target():
        try:
            box["r"] = m._open_dialog_powershell()
            box["e"] = None
        except Exception as e:
            box["e"] = repr(e)
    t = threading.Thread(target=target, daemon=True)
    t.start()
    found = find_and_close(timeout=15)
    t.join(40)
    print("[powershell] dialog_found=%s  returned=%r  exception=%r  thread_alive=%s"
          % (found, box.get("r"), box.get("e"), t.is_alive()))
    return found


if __name__ == "__main__":
    which = sys.argv[1] if len(sys.argv) > 1 else "both"
    m = load_module()
    print("module loaded ok; func comdlg=%s ps=%s"
          % (hasattr(m, "_open_dialog_comdlg"), hasattr(m, "_open_dialog_powershell")))
    if which in ("comdlg", "both"):
        run_comdlg(m)
    if which in ("ps", "both"):
        run_ps(m)
    print("DONE")
