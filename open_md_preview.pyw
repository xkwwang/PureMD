# -*- coding: utf-8 -*-
"""双击 .md → 单实例多标签纯预览窗口（无编辑器、不写任何文件、不改源文件）。

机制：
- 第一个 .md 启动一个常驻窗口，并占用一个本地 TCP 端口作为单实例锁（优先 54321，若被占用则自动改用系统分配的空闲端口）。
- 之后双击的 .md 作为新进程连上该端口，把渲染好的 HTML 发给常驻窗口，新增一个标签页。
- 全部在内存中传输：源 .md 只读、零改动，且不生成任何额外文件。
- 自愈：若已有实例但已损坏（窗口没建起来 / 不响应），新实例会向旧实例发 kill 指令让其退出，自己接管端口，避免"双击完全没反应"。
- 端口信息记录在 mdpreview/.hub.lock（格式：PID<换行>PORT），新实例据此找到在监听的端口。
"""
import sys, os, base64, json, ctypes, socket, threading, time, atexit, subprocess
try:
    import winreg
except Exception:      # 非 Windows 环境下不影响其余逻辑
    winreg = None
import md2docx          # Markdown → .docx 导出（解析参数与浏览器预览一致）

TEMPLATE_REL = os.path.join("mdpreview", "template.html")
HUB_REL = os.path.join("mdpreview", "hub.html")
OVERRIDES_REL = os.path.join("mdpreview", "wb_overrides.css")
LOCK_REL = os.path.join("mdpreview", ".hub.lock")
PLACEHOLDER = "/*__EMBED_JSON_B64__*/"
PRIMARY_PORT = 54321
RETRIES = 8          # 委托给已有实例时的重试次数（含"WAIT"等待）
RETRY_GAP = 0.3      # 重试间隔(秒)


def script_dir():
    return os.path.dirname(os.path.abspath(__file__))


def lock_path():
    return os.path.join(script_dir(), LOCK_REL)


def show_error(msg):
    try:
        ctypes.windll.user32.MessageBoxW(0, str(msg), "PureMD", 0x10)
    except Exception:
        pass


def _selftest():
    """无界面自检：验证 webview 可导入、资源文件就位、docx 导出链可用
    （markdown-it-py + python-docx + lxml 均随 exe 打包；冻结后用于冒烟测试）。"""
    try:
        import webview
        ver = getattr(webview, "__version__", "?")
        base = script_dir()
        assets = {n: os.path.isfile(os.path.join(base, r))
                  for n, r in (("template", TEMPLATE_REL),
                               ("hub", HUB_REL),
                               ("overrides", OVERRIDES_REL))}
        out = "SELFTEST_OK webview=%s assets=%s\n" % (ver, assets)
    except Exception as e:
        out = "SELFTEST_FAIL %r\n" % e
    try:
        b = md2docx.md_to_docx("# 自检\n\n**粗体** 与 `代码`。\n")
        ok = b.getvalue()[:2] == b"PK" and len(b.getvalue()) > 3000   # zip 魔数 + 非空
        out += "DOCX_EXPORT=%s bytes=%d\n" % ("OK" if ok else "BAD", len(b.getvalue()))
    except Exception as e:
        out += "DOCX_EXPORT=FAIL %r\n" % (e,)
    try:
        with open(os.path.join(os.environ.get("TEMP", script_dir()),
                               "mdpreview_selftest.txt"), "w", encoding="utf-8") as f:
            f.write(out)
    except Exception:
        pass


def _read_text(md_path):
    """只读读取 .md 文本，自动处理编码（UTF-8 / 带 BOM / GBK / GB18030 / Big5 等）。永不写回源文件。"""
    with open(md_path, "rb") as f:
        raw = f.read()
    # 1) 优先 UTF-8（最常用，且中文文献多为 UTF-8）
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        pass
    # 2) 常见中文编码：GBK / GB18030（大陆老文档）、Big5（繁体）、UTF-8-SIG
    for enc in ("gb18030", "gbk", "big5", "utf-8-sig"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    # 3) 仍有 chardet 就交给它探测
    try:
        import chardet
        det = chardet.detect(raw)
        if det and det.get("encoding"):
            try:
                return raw.decode(det["encoding"])
            except (UnicodeDecodeError, LookupError):
                pass
    except Exception:
        pass
    # 4) 兜底：替换无法识别的字符，保证能渲染
    return raw.decode("utf-8", errors="replace")


def render_to_html(md_path):
    tpl_path = os.path.join(script_dir(), TEMPLATE_REL)
    if not os.path.isfile(tpl_path):
        raise FileNotFoundError("找不到模板文件：\n" + tpl_path)
    text = _read_text(md_path)  # 只读，不写回
    name = os.path.basename(md_path)
    obj = {"name": name, "text": text}
    b64 = base64.b64encode(json.dumps(obj, ensure_ascii=False).encode("utf-8")).decode("ascii")
    tpl = open(tpl_path, encoding="utf-8").read()
    if PLACEHOLDER not in tpl:
        raise RuntimeError("模板缺少占位符，可能版本不匹配")
    html = tpl.replace(PLACEHOLDER, b64)
    # 叠加 WorkBuddy 同款排版（字体/标题/代码块/链接等），不改变源 .md 一点内容
    try:
        ov_path = os.path.join(script_dir(), OVERRIDES_REL)
        ov = open(ov_path, encoding="utf-8").read()
        style = "<style>/* WorkBuddy 同款排版 */" + ov + "</style>"
        if "</head>" in html:
            html = html.replace("</head>", style + "</head>", 1)
        elif "</body>" in html:
            html = html.replace("</body>", style + "</body>", 1)
    except Exception:
        pass
    return html


def make_payload(md_path):
    # 带上绝对路径：顶栏的“重新加载 / 自动刷新”需要据此回读源文件
    return {"name": os.path.basename(md_path),
            "path": os.path.abspath(md_path),
            "html": render_to_html(md_path)}


def js_update(payload):
    """把重新渲染好的内容推给已有标签页（用于重新加载 / 自动刷新）。"""
    try:
        import webview
        if not webview.windows:
            return False
        js = "(window.__updateTab?window.__updateTab:function(){})(%s)" % json.dumps(payload)
        webview.windows[0].evaluate_js(js)
        return True
    except Exception:
        return False


# 自动刷新：{绝对路径: 上次记录的文件 mtime}
WATCHED = {}


def _watch_loop():
    """后台轮询已开启自动刷新的文件，改动即重新渲染并推送给前端。"""
    while True:
        time.sleep(0.8)
        for path, last in list(WATCHED.items()):
            try:
                m = os.path.getmtime(path)
            except Exception:
                continue
            if last is None:
                WATCHED[path] = m          # 仅记录基线，不触发刷新
            elif m != last:
                WATCHED[path] = m
                try:
                    js_update(make_payload(path))
                except Exception:
                    pass


DLG_TITLE = "选择 Markdown 文件"
# comdlg32 过滤器格式：描述\0模式\0... 以双 \0 结束（ctypes 会自动补最后一个终止符）
FILTER_C = "Markdown 文件\0*.md;*.markdown;*.txt\0所有文件\0*.*\0"
# PowerShell / WinForms 过滤器格式：描述|模式|描述|模式
FILTER_PS = "Markdown 文件 (*.md;*.markdown;*.txt)|*.md;*.markdown;*.txt|所有文件 (*.*)|*.*"


def _open_dialog_comdlg():
    """用 Win32 原生 GetOpenFileNameW 弹选择框。快、零子进程、无需 COM/STA。返回路径或 None（取消）。"""
    from ctypes import wintypes, byref, c_wchar_p, create_unicode_buffer, windll

    class OPENFILENAMEW(ctypes.Structure):
        _fields_ = [
            ("lStructSize", wintypes.DWORD), ("hwndOwner", wintypes.HWND),
            ("hInstance", wintypes.HINSTANCE), ("lpstrFilter", c_wchar_p),
            ("lpstrCustomFilter", wintypes.LPWSTR), ("nMaxCustFilter", wintypes.DWORD),
            ("nFilterIndex", wintypes.DWORD), ("lpstrFile", wintypes.LPWSTR),
            ("nMaxFile", wintypes.DWORD), ("lpstrFileTitle", wintypes.LPWSTR),
            ("nMaxFileTitle", wintypes.DWORD), ("lpstrInitialDir", c_wchar_p),
            ("lpstrTitle", c_wchar_p), ("Flags", wintypes.DWORD),
            ("nFileOffset", wintypes.WORD), ("nFileExtension", wintypes.WORD),
            ("lpstrDefExt", c_wchar_p), ("lCustData", wintypes.LPARAM),
            ("lpfnHook", ctypes.c_void_p), ("lpTemplateName", c_wchar_p),
            ("pvReserved", ctypes.c_void_p), ("dwReserved", wintypes.DWORD),
            ("FlagsEx", wintypes.DWORD),
        ]

    buf = create_unicode_buffer(4096)
    ofn = OPENFILENAMEW()
    ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
    ofn.lpstrFilter = FILTER_C
    ofn.lpstrFile = ctypes.cast(buf, wintypes.LPWSTR)
    ofn.nMaxFile = 4096
    ofn.lpstrTitle = DLG_TITLE
    # OFN_EXPLORER | OFN_FILEMUSTEXIST | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY
    ofn.Flags = 0x00080000 | 0x00001000 | 0x00000800 | 0x00000004
    if not windll.comdlg32.GetOpenFileNameW(byref(ofn)):
        return None
    return buf.value or None


def _open_dialog_powershell():
    """回退：用系统 Windows PowerShell(-STA) 弹 WinForms 选择框。返回路径或 None（取消）。"""
    script = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "$d = New-Object System.Windows.Forms.OpenFileDialog;"
        "$d.Title = '" + DLG_TITLE + "';"
        "$d.Filter = '" + FILTER_PS + "';"
        "$d.Multiselect = $false;"
        "$d.CheckFileExists = $true;"
        "if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $d.FileName }"
    )
    r = subprocess.run(
        ["powershell", "-NoProfile", "-STA", "-NonInteractive", "-Command", script],
        capture_output=True, text=True, timeout=600,
        creationflags=0x08000000,          # CREATE_NO_WINDOW：不闪黑框
    )
    out = (r.stdout or "").strip()
    return out.splitlines()[0].strip() if out else None


def native_open_dialog():
    """返回 (path, err)：path 为 None 表示用户取消；err 非 None 表示两条路都失败。"""
    try:
        return _open_dialog_comdlg(), None
    except Exception as e1:
        try:
            return _open_dialog_powershell(), None
        except Exception as e2:
            return None, "comdlg32: %r；PowerShell: %r" % (e1, e2)


# ---------------------------------------------------------------- 保存对话框（导出 .docx 用）
SAVE_FILTER_C = "Word 文档 (*.docx)\0*.docx\0所有文件\0*.*\0"
SAVE_FILTER_PS = "Word 文档 (*.docx)|*.docx|所有文件 (*.*)|*.*"
SAVE_TITLE = "导出为 Word 文档"


def _save_dialog_comdlg(default_name):
    """Win32 GetSaveFileNameW：快、零子进程。返回路径或 None（取消）。"""
    from ctypes import wintypes, byref, c_wchar_p, create_unicode_buffer, windll

    class OPENFILENAMEW(ctypes.Structure):
        _fields_ = [
            ("lStructSize", wintypes.DWORD), ("hwndOwner", wintypes.HWND),
            ("hInstance", wintypes.HINSTANCE), ("lpstrFilter", c_wchar_p),
            ("lpstrCustomFilter", wintypes.LPWSTR), ("nMaxCustFilter", wintypes.DWORD),
            ("nFilterIndex", wintypes.DWORD), ("lpstrFile", wintypes.LPWSTR),
            ("nMaxFile", wintypes.DWORD), ("lpstrFileTitle", wintypes.LPWSTR),
            ("nMaxFileTitle", wintypes.DWORD), ("lpstrInitialDir", c_wchar_p),
            ("lpstrTitle", c_wchar_p), ("Flags", wintypes.DWORD),
            ("nFileOffset", wintypes.WORD), ("nFileExtension", wintypes.WORD),
            ("lpstrDefExt", c_wchar_p), ("lCustData", wintypes.LPARAM),
            ("lpfnHook", ctypes.c_void_p), ("lpTemplateName", c_wchar_p),
            ("pvReserved", ctypes.c_void_p), ("dwReserved", wintypes.DWORD),
            ("FlagsEx", wintypes.DWORD),
        ]

    buf = create_unicode_buffer(4096)
    buf.value = default_name or "document.docx"
    ofn = OPENFILENAMEW()
    ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
    ofn.lpstrFilter = SAVE_FILTER_C
    ofn.lpstrFile = ctypes.cast(buf, wintypes.LPWSTR)
    ofn.nMaxFile = 4096
    ofn.lpstrTitle = SAVE_TITLE
    ofn.lpstrDefExt = "docx"
    # OFN_EXPLORER | OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY
    ofn.Flags = 0x00080000 | 0x00000002 | 0x00000800 | 0x00000004
    if not windll.comdlg32.GetSaveFileNameW(byref(ofn)):
        return None
    return buf.value or None


def _save_dialog_powershell(default_name):
    """回退：PowerShell(-STA) WinForms SaveFileDialog。返回路径或 None。"""
    safe = (default_name or "document.docx").replace("'", "''")
    script = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "$d = New-Object System.Windows.Forms.SaveFileDialog;"
        "$d.Title = '" + SAVE_TITLE + "';"
        "$d.Filter = '" + SAVE_FILTER_PS + "';"
        "$d.FileName = '" + safe + "';"
        "$d.DefaultExt = 'docx';"
        "$d.OverwritePrompt = $true;"
        "if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $d.FileName }"
    )
    r = subprocess.run(
        ["powershell", "-NoProfile", "-STA", "-NonInteractive", "-Command", script],
        capture_output=True, text=True, timeout=600,
        creationflags=0x08000000,          # CREATE_NO_WINDOW：不闪黑框
    )
    out = (r.stdout or "").strip()
    return out.splitlines()[0].strip() if out else None


def native_save_dialog(default_name):
    """返回 (path, err)：path 为 None 表示用户取消；err 非 None 表示两条路都失败。"""
    try:
        return _save_dialog_comdlg(default_name), None
    except Exception as e1:
        try:
            return _save_dialog_powershell(default_name), None
        except Exception as e2:
            return None, "comdlg32: %r；PowerShell: %r" % (e1, e2)


def _open_dialog_comdlg_multi():
    """多选用：Win32 GetOpenFileNameW + OFN_ALLOWMULTISELECT|OFN_EXPLORER。返回路径列表（可能为空）。"""
    from ctypes import wintypes, byref, c_wchar_p, create_unicode_buffer, windll

    class OPENFILENAMEW(ctypes.Structure):
        _fields_ = [
            ("lStructSize", wintypes.DWORD), ("hwndOwner", wintypes.HWND),
            ("hInstance", wintypes.HINSTANCE), ("lpstrFilter", c_wchar_p),
            ("lpstrCustomFilter", wintypes.LPWSTR), ("nMaxCustFilter", wintypes.DWORD),
            ("nFilterIndex", wintypes.DWORD), ("lpstrFile", wintypes.LPWSTR),
            ("nMaxFile", wintypes.DWORD), ("lpstrFileTitle", wintypes.LPWSTR),
            ("nMaxFileTitle", wintypes.DWORD), ("lpstrInitialDir", c_wchar_p),
            ("lpstrTitle", c_wchar_p), ("Flags", wintypes.DWORD),
            ("nFileOffset", wintypes.WORD), ("nFileExtension", wintypes.WORD),
            ("lpstrDefExt", c_wchar_p), ("lCustData", wintypes.LPARAM),
            ("lpfnHook", ctypes.c_void_p), ("lpTemplateName", c_wchar_p),
            ("pvReserved", ctypes.c_void_p), ("dwReserved", wintypes.DWORD),
            ("FlagsEx", wintypes.DWORD),
        ]

    buf = create_unicode_buffer(8192)
    ofn = OPENFILENAMEW()
    ofn.lStructSize = ctypes.sizeof(OPENFILENAMEW)
    ofn.lpstrFilter = FILTER_C
    ofn.lpstrFile = ctypes.cast(buf, wintypes.LPWSTR)
    ofn.nMaxFile = 8192
    ofn.lpstrTitle = DLG_TITLE + "（可多选）"
    # OFN_EXPLORER|FILEMUSTEXIST|PATHMUSTEXIST|HIDEREADONLY|ALLOWMULTISELECT
    ofn.Flags = 0x00080000 | 0x00001000 | 0x00000800 | 0x00000004 | 0x00000200
    if not windll.comdlg32.GetOpenFileNameW(byref(ofn)):
        return []
    raw = buf.value or ""
    parts = [p for p in raw.split("\x00") if p]
    if len(parts) <= 1:
        return parts                       # 单文件：完整路径
    folder = parts[0]                      # 多文件：首项是文件夹，其余是文件名
    return [os.path.join(folder, f) for f in parts[1:]]


def _open_dialog_powershell_multi():
    script = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "$d = New-Object System.Windows.Forms.OpenFileDialog;"
        "$d.Title = '" + DLG_TITLE + "';"
        "$d.Filter = '" + FILTER_PS + "';"
        "$d.Multiselect = $true;"
        "$d.CheckFileExists = $true;"
        "if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $d.FileNames | ForEach-Object { Write-Output $_ } }"
    )
    r = subprocess.run(
        ["powershell", "-NoProfile", "-STA", "-NonInteractive", "-Command", script],
        capture_output=True, text=True, timeout=600,
        creationflags=0x08000000,
    )
    out = (r.stdout or "").strip()
    return [ln.strip() for ln in out.splitlines() if ln.strip()] if out else []


def native_open_dialog_multi():
    """返回 (paths, err)：paths 为空列表表示取消；err 非 None 表示两条路都失败。"""
    try:
        return _open_dialog_comdlg_multi(), None
    except Exception as e1:
        try:
            return _open_dialog_powershell_multi(), None
        except Exception as e2:
            return None, "comdlg32: %r；PowerShell: %r" % (e1, e2)


# ---------- 本地配置持久化（主题 / 自动刷新等） ----------
def _config_path():
    base = os.environ.get("APPDATA")
    if not base:
        base = script_dir()
    d = os.path.join(base, "PureMD")
    return os.path.join(d, "config.json"), d


def load_config():
    try:
        p, _ = _config_path()
        if os.path.isfile(p):
            with open(p, encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}


def save_config(obj):
    try:
        p, d = _config_path()
        os.makedirs(d, exist_ok=True)
        cur = load_config()
        if isinstance(obj, dict):
            cur.update(obj)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(cur, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


APP_NAME = "PureMD"
APP_VERSION = "1.0.0"
APP_EMAIL = "xkwwang@126.com"
# 发布到 GitHub 后，把这里填成你的 Releases API 地址（留空则不做更新检查），例如：
#   https://api.github.com/repos/<你的用户名>/<仓库名>/releases/latest
UPDATE_CHECK_URL = ""


class Api:
    """暴露给前端 JS 的接口（window.pywebview.api.*）。"""

    def open_file(self):
        # 注意：不能用 pywebview 的 window.create_file_dialog ——
        # EdgeChromium(WebView2) 平台根本没有实现它，调用必然 AttributeError。
        p, err = native_open_dialog()
        if err:
            return {"error": "打开文件对话框失败：" + err}
        if not p:
            return None
        try:
            return make_payload(p)
        except Exception as e:
            return {"error": "渲染失败：%r" % (e,)}

    def reload(self, path):
        if not path or not os.path.isfile(path):
            return {"error": "文件不存在：" + str(path)}
        try:
            return make_payload(path)
        except Exception as e:
            return {"error": "重新渲染失败：%r" % (e,)}

    def export_docx(self, path):
        """把当前 .md 导出为 .docx：读源文件（只读）→ 转换 → 弹保存框 → 写盘。

        返回 {"ok": True, "path": ...}；用户取消返回 None；失败返回 {"error": ...}。
        源 .md 全程只读，不产生任何中间文件。
        """
        if not path or not os.path.isfile(path):
            return {"error": "文件不存在：" + str(path)}
        try:
            text = _read_text(path)                     # 只读，编码自动识别
            buf = md2docx.md_to_docx(text, base_dir=os.path.dirname(os.path.abspath(path)))
        except Exception as e:
            return {"error": "转换失败：%r" % (e,)}
        default = os.path.splitext(os.path.basename(path))[0] + ".docx"
        target, err = native_save_dialog(default)
        if err:
            return {"error": "无法打开保存对话框：" + err}
        if not target:
            return None                                  # 用户取消
        if not target.lower().endswith(".docx"):
            target += ".docx"
        try:
            with open(target, "wb") as f:
                f.write(buf.getvalue())
        except Exception as e:
            return {"error": "写入失败：%r" % (e,)}
        return {"ok": True, "path": target}

    def set_watch(self, path, on):
        try:
            if on:
                if not os.path.isfile(path):
                    return False
                WATCHED[path] = os.path.getmtime(path)   # 先记基线，不立即刷新
            else:
                WATCHED.pop(path, None)
            return True
        except Exception:
            return False

    def feedback(self):
        """仅返回反馈邮箱，由前端弹窗展示；不自动拉起邮件客户端。"""
        return {"ok": True, "email": APP_EMAIL}

    def open_files(self):
        """打开文件对话框（支持多选），返回 payload 列表；取消返回 None。"""
        paths, err = native_open_dialog_multi()
        if err:
            return {"error": "打开文件对话框失败：" + err}
        if not paths:
            return None
        items = []
        for p in paths:
            try:
                items.append(make_payload(p))
            except Exception as e:
                items.append({"error": "渲染失败：%r" % (e,)})
        return items

    def get_config(self):
        return load_config()

    def set_config(self, obj):
        save_config(obj or {})
        return True

    def check_update(self):
        """查询 GitHub Releases 是否有新版本。未配置 UPDATE_CHECK_URL 时返回 None。"""
        if not UPDATE_CHECK_URL:
            return None
        try:
            import urllib.request, json as _json
            req = urllib.request.Request(UPDATE_CHECK_URL, headers={"User-Agent": APP_NAME})
            with urllib.request.urlopen(req, timeout=6) as resp:
                data = _json.loads(resp.read().decode("utf-8"))
            tag = data.get("tag_name") or ""
            ver = tag.lstrip("vV")
            if ver and ver != APP_VERSION:
                return {"latest": ver, "url": data.get("html_url", "")}
        except Exception:
            pass
        return None


def raise_window():
    """双击第二个 .md 时，已有预览窗口可能在后台或最小化；把它恢复到前台。"""
    # 1) pywebview 自带方法（内部用 WinForms Invoke 抛到 GUI 线程，线程安全）
    try:
        import webview
        if webview.windows:
            w = webview.windows[0]
            try:
                w.restore()   # 若被最小化，先恢复为普通窗口
            except Exception:
                pass
            try:
                w.show()      # WinForms: Show() + Activate()，把窗口提到最前
            except Exception:
                pass
    except Exception:
        pass
    # 2) Win32 兜底：按本进程 PID 找到顶层窗口，强制恢复并置顶（应对 Windows 前台锁）
    try:
        u = ctypes.windll.user32
        k = ctypes.windll.kernel32
        EnumWindows = u.EnumWindows
        EnumWindows.restype = ctypes.c_bool
        EnumWindows.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        GetWinThreadProcId = u.GetWindowThreadProcessId
        GetWinThreadProcId.restype = ctypes.c_ulong
        GetWinThreadProcId.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
        IsIconic = u.IsIconic
        IsIconic.restype = ctypes.c_bool
        IsIconic.argtypes = [ctypes.c_void_p]
        ShowWindow = u.ShowWindow
        ShowWindow.restype = ctypes.c_bool
        ShowWindow.argtypes = [ctypes.c_void_p, ctypes.c_int]
        SetWindowPos = u.SetWindowPos
        SetWindowPos.restype = ctypes.c_bool
        SetWindowPos.argtypes = [ctypes.c_void_p, ctypes.c_void_p,
                                 ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                 ctypes.c_int, ctypes.c_uint]
        AllowFG = u.AllowSetForegroundWindow
        AllowFG.restype = ctypes.c_bool
        AllowFG.argtypes = [ctypes.c_int]
        SetFG = u.SetForegroundWindow
        SetFG.restype = ctypes.c_bool
        SetFG.argtypes = [ctypes.c_void_p]

        SW_RESTORE = 9
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        pid = k.GetCurrentProcessId()
        found = []

        @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        def _enum(hwnd, lparam):
            tp = ctypes.c_ulong(0)
            GetWinThreadProcId(hwnd, ctypes.byref(tp))
            if tp.value == pid:
                found.append(hwnd)
            return True

        EnumWindows(_enum, 0)
        for hwnd in found:
            try:
                if IsIconic(hwnd):
                    ShowWindow(hwnd, SW_RESTORE)
            except Exception:
                pass
            try:
                # TOPMOST 再 NOTOPMOST：把窗口顶到 Z 序最前（绕过前台锁）
                SetWindowPos(hwnd, ctypes.c_void_p(-1), 0, 0, 0, 0,
                             SWP_NOMOVE | SWP_NOSIZE)
                SetWindowPos(hwnd, ctypes.c_void_p(-2), 0, 0, 0, 0,
                             SWP_NOMOVE | SWP_NOSIZE)
            except Exception:
                pass
            try:
                AllowFG(-1)
                SetFG(hwnd)
            except Exception:
                pass
    except Exception:
        pass


def js_inject(payload):
    try:
        import webview
        if not webview.windows:
            return False
        js = ("(window.__addTab?window.__addTab"
              ":function(p){(window.__pending=window.__pending||[]).push(p)})(%s)") % json.dumps(payload)
        webview.windows[0].evaluate_js(js)
        raise_window()   # 新增标签页后，把窗口从后台/最小化状态提到最前
        return True
    except Exception:
        return False


def pid_alive(pid):
    """用 Windows API 判断某 PID 是否仍在运行（避免僵尸锁）。"""
    try:
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        h = ctypes.windll.kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
        if h == 0:
            return False
        code = ctypes.c_uint()
        ctypes.windll.kernel32.GetExitCodeProcess(h, ctypes.byref(code))
        ctypes.windll.kernel32.CloseHandle(h)
        return code.value == 259  # STILL_ACTIVE
    except Exception:
        return False


def write_lock(port):
    try:
        with open(lock_path(), "w", encoding="utf-8") as f:
            f.write("%d\n%d\n" % (os.getpid(), port))
    except Exception:
        pass


def read_lock():
    """返回 (pid, port) 或 (None, None)。"""
    try:
        with open(lock_path(), "r", encoding="utf-8") as f:
            lines = f.read().split()
        if not lines:
            return (None, None)
        pid = int(lines[0])
        port = int(lines[1]) if len(lines) > 1 else PRIMARY_PORT
        return (pid, port)
    except Exception:
        return (None, None)


def clear_lock():
    try:
        os.remove(lock_path())
    except Exception:
        pass


# srv 在 start_hub 内赋值，被 server_loop / atexit 捕获
srv = None


def handle_client(c):
    """处理一个连入的连接：kill 指令 / 正常渲染委托，并回 ACK。"""
    try:
        c.settimeout(3)
        data = b""
        while True:
            chunk = c.recv(1 << 20)
            if not chunk:
                break
            data += chunk
    except Exception:
        return
    try:
        payload = json.loads(data.decode("utf-8"))
    except Exception:
        return
    if isinstance(payload, dict) and payload.get("__ctl") == "kill":
        try:
            c.sendall(b"BYE\n")
        except Exception:
            pass
        try:
            c.close()
        except Exception:
            pass
        os._exit(0)  # 立即退出，释放端口，由新实例接管
    if not isinstance(payload, dict) or "__ctl" in payload:
        return
    ok = js_inject(payload)
    try:
        c.sendall(b"OK\n" if ok else b"WAIT\n")
    except Exception:
        pass


def server_loop():
    while True:
        try:
            c, _ = srv.accept()
        except Exception:
            continue
        threading.Thread(target=handle_client, args=(c,), daemon=True).start()


def send_to_hub(payload, port, retries=RETRIES, gap=RETRY_GAP):
    """委托给已有实例。返回 "OK"(已渲染) / "WAIT"(实例还在初始化) / ""(无响应/失败)。"""
    for _ in range(retries):
        try:
            s = socket.socket()
            s.settimeout(2)
            s.connect(("127.0.0.1", port))
            s.sendall(json.dumps(payload).encode("utf-8"))
            s.shutdown(socket.SHUT_WR)
            resp = b""
            try:
                while True:
                    c = s.recv(1024)
                    if not c:
                        break
                    resp += c
            except Exception:
                pass
            s.close()
            r = resp.strip().decode("ascii", "ignore")
            if r == "OK":
                return "OK"
            if r == "WAIT":
                time.sleep(gap)
                continue
            if r:
                return r
        except Exception:
            time.sleep(gap)
    return ""


def kill_hub(port):
    """向已有（损坏的）实例发送 kill 指令，让其退出并释放端口。"""
    try:
        s = socket.socket()
        s.settimeout(2)
        s.connect(("127.0.0.1", port))
        s.sendall(json.dumps({"__ctl": "kill"}).encode("utf-8"))
        s.close()
    except Exception:
        pass


def bind_hub():
    """优先 PRIMARY_PORT；若被占用（如某系统代理/安全 agent）则改用系统分配的空闲端口。

    注意：不能用 SO_REUSEADDR（Windows 上它反而允许两个 socket 共用同一端口，
    会被 snac_agent 这类已占用者“抢走”连接）。改用 SO_EXCLUSIVEADDRUSE，
    端口被占用时 bind 直接失败，从而正确回退到空闲端口。
    """
    global srv
    EXCLUSIVE = 0x00000002  # SO_EXCLUSIVEADDRUSE (Windows)
    def make_sock():
        s = socket.socket()
        try:
            s.setsockopt(socket.SOL_SOCKET, EXCLUSIVE, 1)
        except Exception:
            pass
        return s
    s = make_sock()
    try:
        s.bind(("127.0.0.1", PRIMARY_PORT))
        return s, PRIMARY_PORT
    except Exception:
        pass
    s = make_sock()
    s.bind(("127.0.0.1", 0))
    return s, s.getsockname()[1]


def start_hub(first_payload):
    global srv
    # 若已有实例（但本函数被调用说明 main 已判定无存活实例或需要接管），
    # 这里不再主动杀，直接选端口启动；端口冲突已由 bind_hub 的回退解决。
    srv, used_port = bind_hub()
    srv.listen(8)
    write_lock(used_port)
    atexit.register(clear_lock)
    atexit.register(lambda: srv.close())

    threading.Thread(target=server_loop, daemon=True).start()
    threading.Thread(target=_watch_loop, daemon=True).start()   # 自动刷新轮询

    import webview
    hub_path = os.path.join(script_dir(), HUB_REL)
    try:
        hub_html = open(hub_path, encoding="utf-8").read()
    except Exception as e:
        clear_lock()
        show_error("找不到 hub 界面文件：\n" + str(e))
        return
    # 把上次保存的主题直接注入初始 HTML：首帧即正确（避免先亮后暗的闪动），
    # 也作为“下次打开仍是上次关闭时主题”的强保证。内容主题同样由全局 theme 驱动。
    # 同一份配置里还注入：版本号（帮助→关于弹窗）、上次会话的标签（历史恢复）。
    try:
        _cfg = load_config()
    except Exception:
        _cfg = {}
    try:
        _saved = _cfg.get("theme") if isinstance(_cfg, dict) else None
        if _saved in ("light", "dark", "beige"):
            hub_html = hub_html.replace('var theme = "light";', 'var theme = "%s";' % _saved, 1)
            hub_html = hub_html.replace('var userToggled = false;', 'var userToggled = true;', 1)
    except Exception:
        pass
    # 注入真实版本号（前端“关于”弹窗显示 PureMD vX.Y.Z）
    try:
        hub_html = hub_html.replace('var APP_VERSION = "dev";', 'var APP_VERSION = "%s";' % APP_VERSION, 1)
    except Exception:
        pass
    # 会话恢复：把上次关闭时开着的标签按保存顺序注入 __pending（激活项排最后，
    # 恢复后即为当前标签）；前端 DOM 就绪后会自动 addTab 重建它们。
    # JSON 必须转义 "</" 为 "<\/"，否则内容里的 </script> 会提前截断脚本。
    try:
        _sess = _cfg.get("session") if isinstance(_cfg, dict) else None
        if isinstance(_sess, dict):
            _paths = []
            for _p in (_sess.get("files") or []):
                if isinstance(_p, str) and os.path.isfile(_p) and _p not in _paths:
                    _paths.append(_p)
            _active = _sess.get("active") or ""
            if _active in _paths:
                _paths.remove(_active)
                _paths.append(_active)   # 激活项排最后 → 恢复后即当前标签
            _payloads = []
            for _p in _paths:
                try:
                    _payloads.append(make_payload(_p))
                except Exception:
                    pass
            if _payloads:
                _json = json.dumps(_payloads, ensure_ascii=False).replace("</", "<\\/")
                hub_html = hub_html.replace(
                    "window.__pending = window.__pending || [];",
                    "window.__pending = %s;" % _json, 1)
    except Exception:
        pass
    webview.create_window("PureMD", html=hub_html, width=1000, height=740, js_api=Api())
    # 关键：args 必须是元组，否则 pywebview 会把 dict 当可迭代对象传给
    # threading.Thread(args=dict) 而抛 TypeError，导致窗口永远不出现。
    webview.start(_init_hub, (first_payload,))


def _init_hub(first_payload):
    time.sleep(0.4)
    js_inject(first_payload)


def hub_running():
    pid, _port = read_lock()
    return bool(pid and pid_alive(pid))


# ---------------------------------------------------------------- 文件关联安装
# 说明：本函数被 --setup 调用，目的是让"发给朋友"的版本完全自包含——
# 不再依赖任何本机 Python 路径，exe 自己就能写 HKCU 注册表。
PROG_ID = "PureMD"


def app_open_command():
    """返回注册表中 shell\\open\\command 的命令串。冻结后用 exe 自身；未冻结用 pythonw + 脚本。"""
    if getattr(sys, "frozen", False):
        exe = os.path.abspath(sys.executable)
        return '"%s" "%%1"' % exe
    exe = os.path.abspath(sys.executable)          # pythonw.exe
    scr = os.path.abspath(sys.argv[0])             # open_md_preview.pyw
    return '"%s" "%s" "%%1"' % (exe, scr)


def app_icon_spec():
    """图标来源：冻结后取 exe 内嵌图标（第 0 个）；未冻结退化为解释器图标。"""
    if getattr(sys, "frozen", False):
        return '"%s",0' % os.path.abspath(sys.executable)
    return '"%s",0' % os.path.abspath(sys.executable)


def _del_key_recursive(root, path):
    """删除注册表项（含子项递归删除）。返回 True 表示"已不存在或已删除"。"""
    try:
        winreg.DeleteKey(root, path)
        return True
    except FileNotFoundError:
        return True
    except Exception:
        pass
    try:
        h = winreg.OpenKey(root, path, 0, winreg.KEY_READ | winreg.KEY_WRITE)
    except FileNotFoundError:
        return True
    except Exception:
        return False
    try:
        subs = []
        i = 0
        while True:
            try:
                subs.append(winreg.EnumKey(h, i))
            except OSError:
                break
            i += 1
        for s in subs:
            _del_key_recursive(root, path + "\\" + s)
        winreg.CloseKey(h)
        winreg.DeleteKey(root, path)
        return True
    except Exception:
        return False


def setup_association(silent=False):
    """把 .md 关联到本程序（仅写 HKCU，不需要管理员权限）。返回 (ok, 说明文本)。"""
    if winreg is None:
        return False, "仅支持 Windows"
    try:
        cmd = app_open_command()
        icon = app_icon_spec()

        # 1) 扩展名 -> ProgID
        k = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\.md")
        winreg.SetValue(k, "", winreg.REG_SZ, PROG_ID)
        winreg.CloseKey(k)

        # 2) ProgID：显示名 / 图标 / 打开命令
        k2 = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\%s" % PROG_ID)
        winreg.SetValue(k2, "", winreg.REG_SZ, "PureMD")
        ki = winreg.CreateKey(k2, "DefaultIcon")
        winreg.SetValue(ki, "", winreg.REG_SZ, icon)
        winreg.CloseKey(ki)
        k3 = winreg.CreateKey(k2, r"shell\open\command")
        winreg.SetValue(k3, "", winreg.REG_SZ, cmd)
        winreg.CloseKey(k3)
        winreg.CloseKey(k2)

        # 3) 清掉"打开方式"遗留的 UserChoice（例如被 RStudio 抢走的情况）
        _del_key_recursive(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.md\UserChoice")

        # 4) OpenWithProgids：帮助外壳优先选中我们
        try:
            kp = winreg.CreateKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.md\OpenWithProgids")
            winreg.SetValueEx(kp, PROG_ID, 0, winreg.REG_NONE, b"")
            winreg.CloseKey(kp)
        except Exception:
            pass

        # 5) 通知资源管理器刷新（图标/默认程序立刻生效，无需注销）
        try:
            ctypes.windll.shell32.SHChangeNotify(0x08000000, 0x0000, None, None)
        except Exception:
            pass

        tip = ("已设置成功！\n\n现在双击任意 .md 都会用本地纯预览窗口打开：\n"
               "· 源文件只读，绝不会被改动\n· 不产生任何额外文件\n· 多个 .md 集中在同一窗口的标签页里\n\n"
               "如果图标没立刻变，重启一次\"Windows 资源管理器\"即可。\n\n打开命令：\n" + cmd)
        if not silent:
            try:
                ctypes.windll.user32.MessageBoxW(0, tip, "PureMD · 关联设置", 0x40)
            except Exception:
                pass
        return True, tip
    except Exception as e:
        err = "设置失败：%r" % (e,)
        if not silent:
            try:
                ctypes.windll.user32.MessageBoxW(0, err, "PureMD · 关联设置", 0x10)
            except Exception:
                pass
        return False, err


def main():
    # --setup：安装/修复 .md 文件关联（便携版发给别人后，对方跑一次即可）
    if len(sys.argv) >= 2 and sys.argv[1] == "--setup":
        ok, text = setup_association()
        try:
            sys.stdout.write(text + "\n")
        except Exception:
            pass
        return
    if len(sys.argv) >= 2 and sys.argv[1] == "--selftest":
        _selftest()
        return
    def _prompt_setup():
        """没给 .md 参数（例如直接双击 exe）时，询问是否关联 .md。"""
        try:
            r = ctypes.windll.user32.MessageBoxW(
                0,
                "这是一个 PureMD 纯预览工具：\n"
                "· 只读打开，绝不改动你的 .md 源文件\n"
                "· 不生成任何临时文件\n"
                "· 多个 .md 会集中在同一个窗口的标签页里\n\n"
                "现在要把 .md 文件关联到它吗？关联后双击任意 .md 即可直接预览。\n\n"
                "选「否」也行：把 .md 拖到本程序图标上同样能打开。",
                "PureMD", 0x24 | 0x40)      # MB_YESNO | MB_ICONINFORMATION
        except Exception:
            r = 7
        if r == 6:                                 # IDYES
            setup_association()
        return

    if len(sys.argv) < 2:
        _prompt_setup()
        return
    md_path = sys.argv[1]
    if not os.path.isfile(md_path):
        show_error("找不到文件：\n" + md_path)
        return
    try:
        payload = make_payload(md_path)
    except Exception as e:
        show_error("生成预览失败：\n" + str(e))
        return

    # 已有且活着的实例 → 尝试委托；若实例已损坏则不委托，杀掉它自己接管
    pid, port = read_lock()
    if pid and pid_alive(pid):
        status = send_to_hub(payload, port)
        if status == "OK":
            return
        kill_hub(port)
        time.sleep(0.6)
    clear_lock()
    start_hub(payload)


if __name__ == "__main__":
    main()
