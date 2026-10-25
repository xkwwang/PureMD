# -*- coding: utf-8 -*-
"""双击 .md → 单实例多标签预览窗口（默认只读预览，不写任何文件、不改源文件）。

机制：
- 第一个 .md 启动一个常驻窗口，并占用一个本地 TCP 端口作为单实例锁（优先 54321，若被占用则自动改用系统分配的空闲端口）。
- 之后双击的 .md 作为新进程连上该端口，把渲染好的 HTML 发给常驻窗口，新增一个标签页。
- 全部在内存中传输：源 .md 只读、零改动，且不生成任何额外文件。
- 需要改文字时，点工具栏「编辑」进入编辑模式，改完点「保存」才写回源文件
  （写前自动复制一份 .bak，按原编码与原换行风格写回）——除此之外任何时候都不写盘。
- 自愈：若已有实例但已损坏（窗口没建起来 / 不响应），新实例会向旧实例发 kill 指令让其退出，自己接管端口，避免"双击完全没反应"。
- 端口信息记录在 mdpreview/.hub.lock（格式：PID<换行>PORT），新实例据此找到在监听的端口。
"""
import sys, os, base64, json, ctypes, socket, threading, time, atexit, subprocess, re, shutil
import urllib.request, urllib.parse
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


def _read_text_meta(md_path):
    """读取 .md 并返回 (文本, 编码, 换行符)。

    保存回写时要用"同一个编码 + 同一个换行风格"，否则 GBK 文档会变乱码、
    LF 文档会被改成 CRLF（整篇 diff 全红）。默认读取路径仍是只读的。
    """
    with open(md_path, "rb") as f:
        raw = f.read()
    nl = "\r\n" if b"\r\n" in raw else "\n"
    enc = "utf-8"
    text = None
    if raw[:3] == b"\xef\xbb\xbf":
        enc = "utf-8-sig"
        try:
            text = raw.decode(enc)
        except UnicodeDecodeError:
            text = None
    if text is None:
        try:
            text = raw.decode("utf-8")
            enc = "utf-8"
        except UnicodeDecodeError:
            for cand in ("gb18030", "gbk", "big5", "utf-8-sig"):
                try:
                    text = raw.decode(cand)
                    enc = cand
                    break
                except UnicodeDecodeError:
                    continue
    if text is None:
        try:
            import chardet
            det = chardet.detect(raw)
            if det and det.get("encoding"):
                text = raw.decode(det["encoding"])
                enc = det["encoding"]
        except Exception:
            pass
    if text is None:
        text = raw.decode("utf-8", errors="replace")
        enc = "utf-8"
    return text, enc, nl


def _read_text(md_path):
    """只读读取 .md 文本，自动处理编码（UTF-8 / 带 BOM / GBK / GB18030 / Big5 等）。永不写回源文件。"""
    return _read_text_meta(md_path)[0]


# 本进程内已经备份过的文件（同一文件只备份一次，避免反复保存把原始稿冲掉）
_BACKED_UP = set()


def _backup_once(path):
    """保存前留一份备份，返回备份路径。

    - 同一次运行里，同一个文件只在第一次保存时备份；
    - 若 xxx.md.bak 已存在（以前改过），不覆盖它，改用 xxx.md.bak-时间戳，
      保证最早那一版永远还在。
    """
    key = os.path.abspath(path).lower()
    if key in _BACKED_UP:
        return ""
    target = path + ".bak"
    if os.path.exists(target):
        target = path + ".bak-" + time.strftime("%Y%m%d-%H%M%S")
        n = 1
        while os.path.exists(target):
            target = path + ".bak-" + time.strftime("%Y%m%d-%H%M%S") + "-%d" % n
            n += 1
    shutil.copy2(path, target)
    _BACKED_UP.add(key)
    return target


# 本地图片内联渲染：把 md 引用的本地图片读成 base64 data URI 随 payload 下发，
# 前端渲染时把 <img src> 换成 data URI（原始路径保存到 data-pm-src，编辑模式
# 保存回写时仍写原相对路径）——源 .md 全程只读、零改动。
# 思路与 ErgeMD（Tauri + Rust read_image_as_data_url）一致：
#   WebView2 里 iframe 是不透明来源，相对路径/file:// 都解析不到，base64 是正解。
_IMAGE_MIME = {
    ".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
    ".gif": "image/gif", ".webp": "image/webp", ".svg": "image/svg+xml",
    ".bmp": "image/bmp", ".ico": "image/x-icon", ".avif": "image/avif",
    ".tif": "image/tiff", ".tiff": "image/tiff",
}
MAX_IMAGE_BYTES = 10 * 1024 * 1024      # 单张超过 10MB 不内联（避免 payload 暴涨）
MAX_INLINE_TOTAL = 48 * 1024 * 1024     # 全部内联图片的原始字节总和上限
_MD_IMG_RE = re.compile(
    r"!\[[^\]]*\]\(\s*(<[^>]*>|[^\s)]+)(?:\s+(?:\"[^\"]*\"|'[^']*'|\([^)]*\)))?\s*\)")
_HTML_IMG_RE = re.compile(
    r"<img\b[^>]*?\bsrc\s*=\s*(?:\"([^\"]*)\"|'([^']*)'|([^\s>]+))", re.I)
_URL_SCHEME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9+.-]*:")


def _resolve_local_image(md_path, src):
    """把 md 里的图片 src 解析成本地文件路径；远程/内嵌/不存在返回 ""。"""
    s = (src or "").strip()
    if not s or s.startswith("#") or s.startswith("//"):
        return ""
    # http/data/blob/mailto 等带协议的一律不碰（Windows 盘符 C:/ 例外）
    if _URL_SCHEME_RE.match(s) and not re.match(r"^[A-Za-z]:[\\/]", s):
        return ""
    variants = []
    for base in (s, urllib.parse.unquote(s)):
        base = base.replace("\\", "/")
        if base not in variants:
            variants.append(base)
    md_dir = os.path.dirname(os.path.abspath(md_path))
    for v in variants:
        p = v if re.match(r"^[A-Za-z]:/", v) else os.path.join(md_dir, v)
        p = os.path.normpath(p)
        if os.path.isfile(p):
            return p
    return ""


def _inline_local_images(text, md_path):
    """收集 md 引用的本地图片 → {原始src: dataURI}。只读，绝不写任何文件。"""
    mapping = {}
    srcs = []
    for m in _MD_IMG_RE.finditer(text):
        u = m.group(1)
        if u.startswith("<") and u.endswith(">"):
            u = u[1:-1]
        srcs.append(u)
    for m in _HTML_IMG_RE.finditer(text):
        srcs.append(m.group(1) or m.group(2) or m.group(3) or "")
    total = 0
    for src in srcs:
        if not src or src in mapping:
            continue
        p = _resolve_local_image(md_path, src)
        if not p:
            continue
        mime = _IMAGE_MIME.get(os.path.splitext(p)[1].lower())
        if not mime:
            continue
        try:
            size = os.path.getsize(p)
            if size > MAX_IMAGE_BYTES or total + size > MAX_INLINE_TOTAL:
                continue
            with open(p, "rb") as f:
                data = f.read()
        except OSError:
            continue
        mapping[src] = "data:%s;base64,%s" % (mime, base64.b64encode(data).decode("ascii"))
        total += size
    return mapping


def render_to_html(md_path):
    tpl_path = os.path.join(script_dir(), TEMPLATE_REL)
    if not os.path.isfile(tpl_path):
        raise FileNotFoundError("找不到模板文件：\n" + tpl_path)
    text = _read_text(md_path)  # 只读，不写回
    name = os.path.basename(md_path)
    obj = {"name": name, "text": text}
    images = _inline_local_images(text, md_path)
    if images:
        obj["images"] = images
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


# ---------------------------------------------------------------- 文件夹选择框（侧栏“打开文件夹”用）
FOLDER_TITLE = "选择要浏览的文件夹"


def _pick_folder_comdlg():
    """Win32 SHBrowseForFolderW：快、零子进程。返回路径或 None（取消）。"""
    from ctypes import wintypes, byref, create_unicode_buffer, windll

    class BROWSEINFO(ctypes.Structure):
        _fields_ = [
            ("hwndOwner", wintypes.HWND), ("pidlRoot", ctypes.c_void_p),
            ("pszDisplayName", wintypes.LPWSTR), ("lpszTitle", wintypes.LPCWSTR),
            ("ulFlags", ctypes.c_uint), ("lpfn", ctypes.c_void_p),
            ("lParam", ctypes.c_void_p), ("iImage", ctypes.c_int),
        ]

    disp = create_unicode_buffer(1024)
    bi = BROWSEINFO()
    bi.hwndOwner = 0
    bi.pszDisplayName = ctypes.cast(disp, wintypes.LPWSTR)
    bi.lpszTitle = FOLDER_TITLE
    bi.ulFlags = 0x00000040     # BIF_RETURNONLYFSDIRS：只返回真实文件夹
    sh = windll.shell32
    sh.SHBrowseForFolderW.argtypes = [ctypes.POINTER(BROWSEINFO)]
    sh.SHBrowseForFolderW.restype = ctypes.c_void_p
    sh.SHGetPathFromIDListW.argtypes = [ctypes.c_void_p, wintypes.LPWSTR]
    sh.SHGetPathFromIDListW.restype = ctypes.c_int
    windll.ole32.CoTaskMemFree.argtypes = [ctypes.c_void_p]
    windll.ole32.CoTaskMemFree.restype = None
    pidl = sh.SHBrowseForFolderW(byref(bi))
    if not pidl:
        return None
    out = create_unicode_buffer(1024)
    ok = sh.SHGetPathFromIDListW(pidl, out)
    windll.ole32.CoTaskMemFree(pidl)
    if not ok:
        return None
    return out.value or None


def _pick_folder_powershell():
    """回退：PowerShell(-STA) WinForms FolderBrowserDialog。返回路径或 None。"""
    safe = FOLDER_TITLE.replace("'", "''")
    script = (
        "Add-Type -AssemblyName System.Windows.Forms;"
        "$d = New-Object System.Windows.Forms.FolderBrowserDialog;"
        "$d.Description = '" + safe + "';"
        "$d.ShowNewFolderButton = $false;"
        "if ($d.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Write-Output $d.SelectedPath }"
    )
    r = subprocess.run(
        ["powershell", "-NoProfile", "-STA", "-NonInteractive", "-Command", script],
        capture_output=True, text=True, timeout=600,
        creationflags=0x08000000,
    )
    out = (r.stdout or "").strip()
    return out.splitlines()[0].strip() if out else None


def native_pick_folder():
    """返回 (path, err)：path 为 None 表示用户取消；err 非 None 表示两条路都失败。"""
    try:
        return _pick_folder_comdlg(), None
    except Exception as e1:
        try:
            return _pick_folder_powershell(), None
        except Exception as e2:
            return None, "SHBrowseForFolder: %r；PowerShell: %r" % (e1, e2)


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


_CFG_LOCK = threading.Lock()


def save_config(obj):
    """把 obj 合并进 config.json。

    pywebview 会把每个 js_api 调用放到独立线程，前端的「改颜色」「存高亮」这类
    连续调用可能并发进来；不串行化的话会出现"读到旧内容 → 互相覆盖"（曾经实测
    丢过一次高亮）。这里用锁把「读 → 改 → 写」包成原子操作，并先写临时文件再
    os.replace，避免写一半被关窗口留下半个 JSON。
    """
    try:
        p, d = _config_path()
        os.makedirs(d, exist_ok=True)
        with _CFG_LOCK:
            cur = load_config()
            if isinstance(obj, dict):
                cur.update(obj)
            tmp = p + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(cur, f, ensure_ascii=False, indent=2)
            os.replace(tmp, p)
    except Exception:
        pass


APP_NAME = "PureMD"
APP_VERSION = "1.5.8"
# 发布到 GitHub 后，把这里填成你的 Releases API 地址（留空则不做更新检查），例如：
#   https://api.github.com/repos/<你的用户名>/<仓库名>/releases/latest
UPDATE_CHECK_URL = "https://api.github.com/repos/xkwwang/PureMD/releases/latest"


def _ver_key(s):
    """把 "v1.2.3-beta.1" 之类的版本号转成可比较的元组 (1,2,3)；非数字段按 0，忽略预发行后缀。"""
    s = (s or "").strip().lstrip("vV")
    core = re.split(r"[-+_ ]", s, maxsplit=1)[0]
    nums = []
    for p in core.split("."):
        d = "".join(ch for ch in p if ch.isdigit())
        nums.append(int(d) if d else 0)
    while len(nums) < 3:
        nums.append(0)
    return tuple(nums[:4])


def _ver_gt(a, b):
    """a 是否比 b 新（语义化比较；相等或无法解析时返回 False）。"""
    try:
        return _ver_key(a) > _ver_key(b)
    except Exception:
        return False


class _RECT(ctypes.Structure):
    _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long),
                ("right", ctypes.c_long), ("bottom", ctypes.c_long)]


class _MONITORINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.c_ulong), ("rcMonitor", _RECT),
                ("rcWork", _RECT), ("dwFlags", ctypes.c_ulong)]


def _work_area(hwnd):
    """窗口所在显示器的工作区矩形（不含任务栏 / 自动隐藏的任务栏保留其占位）。

    无边框窗口若走系统 maximize()，会盖到整块屏幕（含任务栏）——WinForms 对
    FormBorderStyle=None 的窗口就是这么处理的。所以最大化改为手动把窗口铺到
    工作区，任务栏始终露在外面。多显示器环境下取窗口当前所在显示器。
    """
    user32 = ctypes.windll.user32
    try:
        hmon = user32.MonitorFromWindow(hwnd, 2)      # MONITOR_DEFAULTTONEAREST
        if hmon:
            mi = _MONITORINFO()
            mi.cbSize = ctypes.sizeof(_MONITORINFO)
            if user32.GetMonitorInfoW(hmon, ctypes.byref(mi)):
                return (mi.rcWork.left, mi.rcWork.top, mi.rcWork.right, mi.rcWork.bottom)
    except Exception:
        pass
    try:                                              # 兜底：主显示器工作区
        rc = _RECT()
        if user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rc), 0):   # SPI_GETWORKAREA
            return (rc.left, rc.top, rc.right, rc.bottom)
    except Exception:
        pass
    return None


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

    # ---- 左侧文件夹侧栏 ----
    def pick_folder(self):
        """弹出文件夹选择框，返回选中路径字符串；取消返回 None；失败返回 {"error": ...}。"""
        p, err = native_pick_folder()
        if err:
            return {"error": "无法打开文件夹选择框：" + err}
        return p

    def list_dir(self, path):
        """列出某文件夹下的子文件夹与 .md/.markdown 文件（隐藏项跳过），按字母序。
        返回 {"folders": [...], "files": [...]}；失败返回 {"error": ...}。"""
        if not path or not os.path.isdir(path):
            return {"error": "无效的文件夹路径"}
        try:
            folders, files = [], []
            for name in os.listdir(path):
                if name.startswith("."):
                    continue
                full = os.path.join(path, name)
                try:
                    if os.path.isdir(full):
                        folders.append(name)
                    elif os.path.isfile(full) and name.lower().endswith((".md", ".markdown")):
                        files.append(name)
                except Exception:
                    pass
            folders.sort(key=lambda s: s.lower())
            files.sort(key=lambda s: s.lower())
            return {"folders": folders, "files": files}
        except Exception as e:
            return {"error": "读取文件夹失败：%r" % (e,)}

    def open_md_path(self, path):
        """把指定 .md 渲染成 payload（与双击/打开文件走同一条渲染通道）。"""
        if not path or not os.path.isfile(path):
            return {"error": "文件不存在：" + str(path)}
        try:
            return make_payload(path)
        except Exception as e:
            return {"error": "渲染失败：%r" % (e,)}

    # ---- 编辑模式：把改好的文本写回源文件（必须由用户显式点「保存」触发）----
    def save_md(self, path, text):
        """把 text 写回 path。

        安全约定（用户要求的底线：不改则已，改则留痕）：
          1. 只在这个接口被显式调用时才写盘，预览 / 高亮本身永不自动写；
          2. 写之前先把原文件复制一份 xxx.md.bak（已有备份则不覆盖，避免把最原始
             的版本冲掉）；
          3. 按原文件的编码与换行风格写回（GBK 存 GBK、CRLF 存 CRLF）。
        返回 {ok, bytes, backup, mtime} 或 {error}。
        """
        enc = None
        try:
            if not path or not os.path.isfile(path):
                return {"error": "文件不存在：" + str(path)}
            if not isinstance(text, str):
                return {"error": "内容格式不对"}
            old_text, enc, nl = _read_text_meta(path)
            if text == old_text:
                return {"ok": True, "bytes": 0, "backup": "", "unchanged": True}
            # 统一成原文件的换行风格，避免整篇 diff 全红
            body = text.replace("\r\n", "\n")
            if nl == "\r\n":
                body = body.replace("\n", "\r\n")
            data = body.encode(enc)
            backup = _backup_once(path)
            with open(path, "wb") as f:
                f.write(data)
            st = os.stat(path)
            # 自己刚写的，别让「自动刷新」误判成外部修改又重渲染一遍
            try:
                WATCHED[path] = st.st_mtime
                WATCHED[os.path.abspath(path)] = st.st_mtime
            except Exception:
                pass
            return {"ok": True, "bytes": len(data), "backup": backup or "",
                    "mtime": st.st_mtime, "encoding": enc}
        except UnicodeEncodeError as e:
            return {"error": "原文件是 %s 编码，新内容里有它存不下的字符（%s）。"
                             "可以先把文件另存为 UTF-8 再编辑。" % (enc or "未知", e.reason)}
        except PermissionError:
            return {"error": "文件被占用或设成了只读，写不进去。请关掉占用它的程序，或取消只读属性。"}
        except Exception as e:
            return {"error": "保存失败：%r" % (e,)}

    def get_config(self):
        return load_config()

    def set_config(self, obj):
        save_config(obj or {})
        return True

    def check_update(self, manual=False):
        """查询 GitHub Releases 是否有新版本。

        返回 {ok, has_update, current, latest, url, name, notes, error}：
        - manual=True：用户主动点「检查更新」——忽略自动开关与"跳过此版本"，失败时带错误详情
        - manual=False：启动时自动检查——尊重 auto_update_check 开关与 skipped_version，失败静默
        网络访问走 urllib 默认链路（环境变量 / 系统代理均自动生效）。
        """
        result = {"ok": False, "has_update": False, "current": APP_VERSION, "latest": ""}
        if not UPDATE_CHECK_URL:
            result["error"] = "未配置更新检查地址"
            return result
        cfg = load_config()
        if not manual and cfg.get("auto_update_check") is False:
            return result                      # 用户关掉了自动检查：静默跳过
        try:
            req = urllib.request.Request(UPDATE_CHECK_URL, headers={
                "User-Agent": "%s/%s" % (APP_NAME, APP_VERSION),
                "Accept": "application/vnd.github+json",
            })
            with urllib.request.urlopen(req, timeout=8) as resp:
                data = json.loads(resp.read().decode("utf-8"))
            latest = (data.get("tag_name") or "").strip().lstrip("vV")
            has = bool(latest) and _ver_gt(latest, APP_VERSION)
            if not manual and latest and latest == (cfg.get("skipped_version") or ""):
                has = False                    # 用户此前选择跳过这个版本
            result.update({
                "ok": True,
                "has_update": has,
                "latest": latest,
                "url": data.get("html_url") or "",
                "name": data.get("name") or data.get("tag_name") or "",
                "notes": (data.get("body") or "")[:600],
            })
        except Exception as e:
            result["error"] = ("%s: %s" % (type(e).__name__, e)) if manual else "网络不可达或超时"
        return result

    def open_url(self, url):
        """用系统默认浏览器打开 URL（打开发布页下载新版本）。"""
        if not url:
            return False
        try:
            os.startfile(url)                  # Windows 首选
            return True
        except Exception:
            pass
        try:                                    # 兜底：rundll32（避开 WebView2 工作线程直调 ShellExecute 的坑）
            subprocess.Popen(["rundll32.exe", "url.dll,FileProtocolHandler", url],
                             creationflags=0x08000000)
            return True
        except Exception:
            return False

    # ---- 窗口控制：优先 pywebview，失败则回退 Windows 原生消息 ----
    # 回退很关键：一旦 pywebview 侧调用失效，窗口会「关不掉 / 最小不了」。
    _drag_origin = None
    _drag_start_xy = (None, None)   # 最近一次 drag_start 的鼠标坐标（window_move 兜底复用）
    # 手动最大化状态（见 window_maximize）：无边框窗口走系统 maximize() 会连
    # 任务栏一起盖住，所以改为 SetWindowPos 铺满「工作区」，状态在此自管。
    _maximized = False
    _max_rect = None        # 最大化前的窗口矩形 (L, T, R, B)，还原用
    _max_wa = None          # 最大化时铺的工作区矩形 (L, T, R, B)，拖动摆位按它做比例映射

    def _hwnd(self):
        """取主窗口句柄（选择规则见 _own_main_hwnd 的注释）。"""
        user32 = ctypes.windll.user32

        # 1) pywebview 的 native（WinForms BrowserForm）最可靠
        try:
            native = getattr(webview.windows[0], "native", None)
            if native is not None:
                h = int(native.Handle.ToInt64())
                if h and user32.IsWindow(h):
                    return h
        except Exception:
            pass

        # 2) 本进程面积最大的顶层窗口
        h = _own_main_hwnd()
        if h:
            return h

        # 3) 兜底：按标题找，但必须属于本进程；否则取本进程的前台窗口
        try:
            import ctypes.wintypes as wt
            mypid = os.getpid()
            pid = wt.DWORD()
            h = user32.FindWindowW(None, "PureMD")
            if h:
                user32.GetWindowThreadProcessId(h, ctypes.byref(pid))
                if pid.value == mypid:
                    return h
            fg = user32.GetForegroundWindow()
            if fg:
                user32.GetWindowThreadProcessId(fg, ctypes.byref(pid))
                if pid.value == mypid:
                    return fg
        except Exception:
            pass
        return 0

    def window_minimize(self):
        try:
            webview.windows[0].minimize()
            return True
        except Exception:
            pass
        try:                                    # WM_SYSCOMMAND / SC_MINIMIZE
            h = self._hwnd()
            if h:
                ctypes.windll.user32.PostMessageW(h, 0x0112, 0xF020, 0)
                return True
        except Exception:
            pass
        return False

    def window_maximize(self):
        """最大化到「工作区」：桌面任务栏始终露在外面（无边框窗口的系统
        maximize() 会连任务栏一起盖住，故手动 SetWindowPos 到工作区矩形）。
        幂等：已最大化时直接返回。"""
        h = self._hwnd()
        if not h:
            return False
        if Api._maximized:
            return True
        try:
            rect = _RECT()
            ctypes.windll.user32.GetWindowRect(h, ctypes.byref(rect))
            Api._max_rect = (int(rect.left), int(rect.top), int(rect.right), int(rect.bottom))
            wa = _work_area(h)
            if not wa:
                return False
            # 先 SW_RESTORE：清掉可能存在的系统最大化/最小化状态，再铺到工作区
            ctypes.windll.user32.ShowWindow(h, 9)   # SW_RESTORE
            ctypes.windll.user32.SetWindowPos(
                h, None, int(wa[0]), int(wa[1]),
                int(wa[2] - wa[0]), int(wa[3] - wa[1]),
                0x0004)                               # SWP_NOZORDER
            Api._maximized = True
            Api._max_wa = (int(wa[0]), int(wa[1]), int(wa[2]), int(wa[3]))
            _push_max_state(True)
            return True
        except Exception:
            Api._maximized = False
            Api._max_rect = None
            return False

    def window_restore(self):
        """从最大化还原成原来的窗口矩形；未最大化时兜底走系统还原。"""
        h = self._hwnd()
        if not h:
            return False
        try:
            if Api._maximized and Api._max_rect:
                L, T, R, B = Api._max_rect
                ctypes.windll.user32.SetWindowPos(
                    h, None, int(L), int(T), int(R - L), int(B - T), 0x0004)
                Api._maximized = False
                Api._max_rect = None
                Api._max_wa = None
                _push_max_state(False)
                return True
        except Exception:
            pass
        try:
            webview.windows[0].restore()
            return True
        except Exception:
            pass
        try:                                    # WM_SYSCOMMAND / SC_RESTORE
            if h:
                ctypes.windll.user32.PostMessageW(h, 0x0112, 0xF120, 0)
                return True
        except Exception:
            pass
        return False

    def window_close(self):
        try:
            webview.windows[0].destroy()
            return True
        except Exception:
            pass
        try:                                    # WM_CLOSE 兜底，保证一定能关掉
            h = self._hwnd()
            if h:
                ctypes.windll.user32.PostMessageW(h, 0x0010, 0, 0)
                return True
        except Exception:
            pass
        return False

    # ---- 自绘标题栏拖动：绕开 pywebview 内置 easy_drag（实测在 Edge 内核不可靠） ----
    # JS 侧累计相对位移，Python 侧叠加到拖动开始时的窗口原点上。
    def window_drag_start(self, mx=None, my=None):
        """开始拖动标题栏。mx/my 为 mousedown 时的屏幕坐标（最大化还原要用）。
        最大化状态下：先还原成原尺寸，再把窗口摆到鼠标下方（抓取点按比例
        映射到还原后的标题栏），随后 window_move 即可无缝跟随——与原生
        Windows「拖最大化窗口先还原再跟手」一致。"""
        try:
            h = self._hwnd()
            if not h:
                Api._drag_origin = None
                return False
            user32 = ctypes.windll.user32
            Api._drag_start_xy = (mx, my)
            if Api._maximized and mx is not None and my is not None and Api._max_rect:
                # 1) 先把要用的数据取出来——window_restore() 会把 _max_rect /
                #    _maximized / _max_wa 全部清掉，之后再解包就是 None（曾经
                #    正是如此：TypeError 被外层 except 吞掉，摆位整段成了死代码，
                #    窗口只是原地还原、光标不在标题栏上，大幅拖动时观感很差）。
                L0, T0, R0, B0 = Api._max_rect
                wa_max = Api._max_wa or _work_area(h) or (L0, T0, R0, B0)
                # 2) 还原到原尺寸（同时清掉 _maximized 标志）
                if not self.window_restore():
                    return False
                # 3) 把窗口摆到鼠标下方：抓取点在最大化窗口里的相对位置，按
                #    宽/高比例映射到还原后的窗口；y 钳在标题栏高度内，保证
                #    光标始终落在标题栏上（原生 Windows 也是这个手感）。
                mrect = _RECT()
                user32.GetWindowRect(h, ctypes.byref(mrect))
                rw = max(1, mrect.right - mrect.left)
                rh = max(1, mrect.bottom - mrect.top)
                mw = max(1, int(wa_max[2]) - int(wa_max[0]))
                mh = max(1, int(wa_max[3]) - int(wa_max[1]))
                fx = (float(mx) - float(wa_max[0])) / mw
                fy = (float(my) - float(wa_max[1])) / mh
                fx = min(max(fx, 0.0), 1.0)
                offx = int(round(fx * rw))
                offy = int(round(fy * rh))
                offy = max(0, min(offy, 40))          # 钳在标题栏高度内，防止乱跳
                newL = int(mx) - offx
                newT = int(my) - offy
                # 别把窗口拖出工作区（顶部至少露出来，左右留 80px 可抓）
                wa = _work_area(h) or wa_max
                newT = max(int(wa[1]), min(newT, int(wa[3]) - 40))
                newL = max(int(wa[0]) - rw + 80,
                           min(newL, int(wa[2]) - 80))
                user32.SetWindowPos(h, None, newL, newT, 0, 0, 0x0001 | 0x0004)
                Api._drag_origin = (newL, newT)
                return True
            # 系统最大化（Win+方向键等路径）时同样不允许直接拖：先还原再拖更自然
            if user32.IsZoomed(h):
                Api._drag_origin = None
                return False
            rect = (ctypes.c_long * 4)()        # 等价于 RECT：left, top, right, bottom
            user32.GetWindowRect(h, ctypes.byref(rect))
            Api._drag_origin = (rect[0], rect[1])
            return True
        except Exception:
            Api._drag_origin = None
            return False

    def window_move(self, dx, dy):
        try:
            org = Api._drag_origin
            if not org and not self.window_drag_start(*(Api._drag_start_xy or (None, None))):
                return False
            org = Api._drag_origin
            h = self._hwnd()
            if not h:
                return False
            # SWP_NOSIZE(0x0001) | SWP_NOZORDER(0x0004)
            # 注意：绝不能带 SWP_NOMOVE(0x0002)，否则 SetWindowPos 会忽略上面传入的
            # x/y 坐标，窗口就完全拖不动（本 bug 曾导致拖动失效）。
            ctypes.windll.user32.SetWindowPos(
                h, None, int(org[0]) + int(dx), int(org[1]) + int(dy),
                0, 0, 0x0001 | 0x0004,
            )
            return True
        except Exception:
            return False

    # ---- 自由缩放 ----
    # 两个坑：
    #   ① JS 侧的 API 调用是异步的，若在 mousedown 里先 await 取窗口矩形再挂监听，
    #     这段延迟内的 mousemove 会全部丢失（表现为「第一次拖边缘没反应」）。
    #   ② pywebview 的几个 API 调用之间也不保证先来后到（各自派发），
    #     所以不能靠「先调 start 记起点、再调 to 应用位移」两次调用来配合。
    # 解法：每次拖拽带一个自增的手势号 gid；Python 发现 gid 变了就自己用当前矩形
    # 作为起点（此刻还没被本手势改过，必然准确），之后只应用「累计位移」。
    # 累计位移的好处：即使调用排队、丢帧，最后落地的那次一定对应正确的最终几何。
    _resize_gid = -1
    _resize_origin = None      # (L, T, R, B) 物理像素
    _resize_base = (0.0, 0.0)  # 位移基准（最大化还原那一帧的 dx/dy）
    _resize_scale = 1.0        # 物理像素 / CSS 像素（兼容 DPI 缩放）

    def window_resize_to(self, gid, edge, dx, dy, css_w=None, css_h=None, mx=None, my=None):
        """应用「相对拖拽起点的累计位移」（dx/dy 为 CSS 像素）。

        mx / my 是鼠标当前的屏幕坐标：只在「最大化 / 贴靠」状态下拖边缘时用得上——
        此时系统不让窗口跟着鼠标变尺寸（表现为拖了没反应），先还原成普通窗口，
        再把被拖的那条边直接对齐到鼠标位置，从这一帧起继续跟随，手感与原生一致。
        """
        try:
            h = self._hwnd() or _find_puremd_hwnd()
            if not h:
                return False
            try:
                gid = int(gid)
            except Exception:
                gid = Api._resize_gid
            if Api._maximized or ctypes.windll.user32.IsZoomed(h):
                # 最大化：先还原（还原后的矩形才是可自由缩放的状态）。
                # _maximized 是手动工作区最大化，走 window_restore 回到原矩形；
                # IsZoomed 是系统路径（Win+↑ 等），用 SW_RESTORE。
                try:
                    if Api._maximized:
                        self.window_restore()          # 内部已推送 False
                    else:
                        ctypes.windll.user32.ShowWindow(h, 9)   # SW_RESTORE
                        _push_max_state(False)         # 系统最大化路径，图标也要翻回去
                except Exception:
                    return False
                rect = (ctypes.c_long * 4)()
                ctypes.windll.user32.GetWindowRect(h, ctypes.byref(rect))
                L, T, R, B = int(rect[0]), int(rect[1]), int(rect[2]), int(rect[3])
                # 把被拖动的边/角对齐到鼠标当前位置，避免"还原后窗口尺寸突跳"
                edge_s = str(edge or "")
                try:
                    if "e" in edge_s and mx is not None: R = int(mx)
                    if "w" in edge_s and mx is not None: L = int(mx)
                    if "s" in edge_s and my is not None: B = int(my)
                    if "n" in edge_s and my is not None: T = int(my)
                except Exception:
                    pass
                if R - L < _MIN_W:
                    R = L + _MIN_W if "w" not in edge_s else L + _MIN_W
                if B - T < _MIN_H:
                    B = T + _MIN_H if "n" not in edge_s else T + _MIN_H
                try:
                    ctypes.windll.user32.SetWindowPos(h, None, L, T, R - L, B - T, 0x0004)
                except Exception:
                    pass
                Api._resize_origin = (L, T, R, B)
                Api._resize_base = (float(dx or 0), float(dy or 0))   # 之后按"相对这一帧"的位移算
                Api._resize_gid = gid
                Api._resize_scale = _dpi_scale(h)
                return True
            if gid != Api._resize_gid:
                # 新手势：记录起点矩形，并按 CSS 宽高反推「每 CSS 像素多少物理像素」
                rect = (ctypes.c_long * 4)()
                ctypes.windll.user32.GetWindowRect(h, ctypes.byref(rect))
                L0, T0, R0, B0 = int(rect[0]), int(rect[1]), int(rect[2]), int(rect[3])
                Api._resize_origin = (L0, T0, R0, B0)
                Api._resize_base = (0.0, 0.0)
                Api._resize_gid = gid
                k = 1.0
                try:
                    cw = float(css_w or 0)
                    if cw > 10:
                        k = (R0 - L0) / cw
                        if not (0.2 < k < 8):
                            k = 1.0
                except Exception:
                    k = 1.0
                if abs(k - 1.0) < 1e-6:
                    k = _dpi_scale(h)      # CSS 宽高不可信时，退回系统 DPI 比例
                Api._resize_scale = k
            org = Api._resize_origin
            if not org:
                return False
            k = Api._resize_scale or 1.0
            bx, by = Api._resize_base or (0.0, 0.0)
            DX = int(round((float(dx) - bx) * k))
            DY = int(round((float(dy) - by) * k))
            L, T, R, B = org
            edge = str(edge or "")
            if "w" in edge:
                L += DX
            if "e" in edge:
                R += DX
            if "n" in edge:
                T += DY
            if "s" in edge:
                B += DY
            if R - L < _MIN_W:
                if "w" in edge:
                    L = R - _MIN_W
                else:
                    R = L + _MIN_W
            if B - T < _MIN_H:
                if "n" in edge:
                    T = B - _MIN_H
                else:
                    B = T + _MIN_H
            # 只改位置与尺寸，保持 Z 序：SWP_NOZORDER(0x0004)
            ctypes.windll.user32.SetWindowPos(
                h, None, L, T, R - L, B - T, 0x0004)
            return True
        except Exception:
            return False

    def apply_window_features(self):
        """前端加载完成后调用：补上任务栏最小化所需样式。返回是否成功。"""
        return bool(enable_window_features())


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


# ---------------------------------------------------------------- 窗口特性：任务栏最小化 + 自由缩放
# 根因：frameless=True 时 WinForms 把 Form 设为 FormBorderStyle.None，于是窗口样式里
#   既没有 WS_MINIMIZEBOX 也没有 WS_SYSMENU，导致「点任务栏图标无法最小化」。
# 做法（轻量、不子类化 WndProc —— 子类化 WinForms 窗口过程会破坏 WebView2 消息链，风险高）：
#   · 窗口建好后补上 WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_SYSMENU（任务栏/系统可最小化/最大化），
#     并补 WS_EX_APPWINDOW 保证任务栏有按钮。这几个位不会在 FormBorderStyle.None 上画可见标题栏，
#     仅影响 shell 行为，外观仍干净。
#   · 自由缩放改用 HTML 自绘的 8 个边缘/角手柄（见 hub.html），拖拽时调用 window_get_rect /
#     window_set_rect，由 Python 用 SetWindowPos 直接改窗口尺寸，完全可控、不依赖系统 NCHITTEST。
#   · 最小尺寸在 window_set_rect 内强制。
_user32 = ctypes.windll.user32

GWL_STYLE = -16
GWL_EXSTYLE = -20
WS_MINIMIZEBOX = 0x20000    # 任务栏点击可最小化（依赖 WS_SYSMENU）
WS_MAXIMIZEBOX = 0x10000
WS_SYSMENU = 0x80000
WS_EX_APPWINDOW = 0x40000   # 保证任务栏有按钮
SWP_FRAMECHANGED = 0x0020
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_NOZORDER = 0x0004
_MIN_W, _MIN_H = 480, 340


def _dpi_scale(hwnd):
    """窗口所在显示器的缩放比例（物理像素 / CSS 像素）；拿不到就按 1.0。"""
    try:
        u = ctypes.windll.user32
        get_dpi = getattr(u, "GetDpiForWindow", None)
        if get_dpi is not None:
            d = int(get_dpi(ctypes.c_void_p(hwnd)))
            if 48 < d < 480:
                return d / 96.0
    except Exception:
        pass
    return 1.0


def _own_main_hwnd():
    """返回本进程面积最大的顶层窗口（主窗口）。

    必须按「进程 + 面积」挑，不能只按标题：Win11 任务栏缩略图会创建一个
    Windows.Internal.Shell.TabProxyWindow，标题同样是 "PureMD"（实测 136x39），
    FindWindowW 会先命中它，后续 SetWindowPos / 样式位就全打到代理窗口上了。
    """
    try:
        import ctypes.wintypes as wt
        mypid = os.getpid()
        cand = []

        @ctypes.WINFUNCTYPE(wt.BOOL, wt.HWND, wt.LPARAM)
        def _cb(h, _l):
            pid = wt.DWORD()
            _user32.GetWindowThreadProcessId(h, ctypes.byref(pid))
            if pid.value != mypid:
                return True
            r = wt.RECT()
            _user32.GetWindowRect(h, ctypes.byref(r))
            ww, hh = r.right - r.left, r.bottom - r.top
            if ww < 200 or hh < 150:
                return True
            score = ww * hh
            cls = ctypes.create_unicode_buffer(256)
            _user32.GetClassNameW(h, cls, 256)
            if cls.value.startswith("WindowsForms10"):
                score += 1 << 40          # 主窗体（WinForms）
            if _user32.IsWindowVisible(h):
                score += 1 << 30
            cand.append((score, h))
            return True

        _user32.EnumWindows(_cb, 0)
        if cand:
            cand.sort()
            return cand[-1][1]
    except Exception:
        pass
    return 0


def _find_puremd_hwnd(retries=60):
    """等主窗口建好（每次 0.1s，最多 retries 次）；找不到再按标题兜底（需属于本进程）。"""
    for _ in range(retries):
        h = _own_main_hwnd()
        if h:
            return h
        time.sleep(0.1)
    try:
        import ctypes.wintypes as wt
        h = _user32.FindWindowW(None, "PureMD")
        if h:
            pid = wt.DWORD()
            _user32.GetWindowThreadProcessId(h, ctypes.byref(pid))
            if pid.value == os.getpid():
                return h
    except Exception:
        pass
    return 0


def enable_window_features():
    """给 frameless 窗口补上最小化/最大化框与系统菜单（任务栏最小化所需），不子类化 WndProc。"""
    try:
        hwnd = _find_puremd_hwnd()
        if not hwnd:
            return False
        _user32.GetWindowLongPtrW.restype = ctypes.c_int64
        _user32.GetWindowLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int]
        _user32.SetWindowLongPtrW.restype = ctypes.c_int64
        _user32.SetWindowLongPtrW.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_int64]
        _user32.SetWindowPos.restype = ctypes.c_int
        _user32.SetWindowPos.argtypes = [ctypes.c_void_p, ctypes.c_void_p,
                                         ctypes.c_int, ctypes.c_int, ctypes.c_int,
                                         ctypes.c_int, ctypes.c_uint]

        style = _user32.GetWindowLongPtrW(hwnd, GWL_STYLE)
        _user32.SetWindowLongPtrW(
            hwnd, GWL_STYLE,
            style | WS_MINIMIZEBOX | WS_MAXIMIZEBOX | WS_SYSMENU)
        ex = _user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
        _user32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, ex | WS_EX_APPWINDOW)
        # 强制重算，让新样式立刻生效
        _user32.SetWindowPos(hwnd, None, 0, 0, 0, 0,
                             SWP_FRAMECHANGED | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER)
        return True
    except Exception:
        return False


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


def _push_max_state(maxed):
    """把窗口最大化状态推给前端，同步 maxBtn 图标（□ ⇄ ⧉）。

    除按钮点击外，「拖标题栏还原」「拖边缘还原」这两条路径也会改变最大化
    状态，若不通告前端，按钮会停留在「还原」图标上，与窗口实际状态相反。
    """
    try:
        import webview
        if not webview.windows:
            return False
        js = ("if(window.__pmMaxState)window.__pmMaxState(%s)"
              % ("true" if maxed else "false"))
        webview.windows[0].evaluate_js(js)
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
    # 会话恢复：把上次关闭时开着的标签按保存顺序重建（激活项排最后，恢复后即
    # 当前标签）。重要：这些 payload 累计渲染后 HTML 常达 1~2MB+，绝不能像旧版
    # 那样内联进 hub_html 再走 create_window(html=...) 注入——pywebview 的 Edge
    # 后端用 CoreWebView2.NavigateToString 喂字符串，而微软明确限定其总大小
    # 不得超过 2MB，超限时导航静默失败、整页白屏（这正是之前 7 文件会话必白屏的
    # 真正根因，与 <body> 标签污染无关）。故此处只“收集”payload，待窗口就绪后改
    # 走与“运行时双击新标签”完全相同的通道（js_inject → evaluate_js，即
    # ExecuteScriptAsync，无 2MB 限制）逐条 addTab，从根上避开该限制。
    _session_payloads = []
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
            # 上次关闭时每个文件的阅读位置（滚动像素），按绝对路径索引
            _scroll = _sess.get("scroll") if isinstance(_sess.get("scroll"), dict) else {}
            _scroll_abs = {}
            for _k, _v in _scroll.items():
                try:
                    _scroll_abs[os.path.abspath(_k)] = int(_v or 0)
                except Exception:
                    pass
            _first_path = (first_payload or {}).get("path")
            for _p in _paths:
                if _first_path and os.path.abspath(_p) == os.path.abspath(_first_path):
                    continue   # 避免与“本次双击启动的文件”重复成两个标签
                try:
                    _pay = make_payload(_p)
                    # 带上上次读到的位置，前端加载完成后自动滚回原位（不是从头开始）
                    _pay["scrollY"] = _scroll_abs.get(os.path.abspath(_p), 0)
                    _session_payloads.append(_pay)
                except Exception:
                    pass
            # 本次双击启动的文件同样适用：若上次会话记过它的位置，就从原位继续
            try:
                if _first_path and isinstance(first_payload, dict):
                    _y = _scroll_abs.get(os.path.abspath(_first_path))
                    if _y:
                        first_payload["scrollY"] = _y
            except Exception:
                pass
    except Exception:
        pass
    # 仅允许标题栏（hub.html 里的 .pm-drag-region）拖动窗口，避免点工具栏/标签误拖窗口。
    # 注意：标记类名不可使用 .pywebview-drag-region——那是 pywebview 内置 easy_drag 的
    # 选择器（body mousedown 监听无条件注册，easy_drag=False 也关不掉），复用它会激活
    # pywebview 那条坏掉的 move 链路并在后台持续抛错。详见 hub.html 拖动处的注释。
    try:
        webview.settings['DRAG_REGION_DIRECT_TARGET_ONLY'] = True
    except Exception:
        pass
    # easy_drag=False：不用 pywebview 内置拖动（它依赖 window.move，在 Edge 内核上不可靠），
    # 改由 hub.html 调用 window_drag_start / window_move 自己实现拖动。
    webview.create_window("PureMD", html=hub_html, width=1000, height=740,
                          js_api=Api(), frameless=True, easy_drag=False)
    # 关键：args 必须是元组，否则 pywebview 会把 dict 当可迭代对象传给
    # threading.Thread(args=dict) 而抛 TypeError，导致窗口永远不出现。
    webview.start(_init_hub, (first_payload, _session_payloads))


def _disable_native_zoom(retries=24):
    """关掉 WebView2 自带的「整页缩放」（Ctrl+滚轮 / Ctrl+加减 / 双指捏合）。

    前端把 Ctrl+滚轮 改成了「只缩放正文」：zoom 只加在每条标签页 iframe 内的
    <html> 上，父页面（标签栏、工具条、侧栏）一个像素都不缩放。如果内核还保留
    自己的整页缩放，两套会叠加——界面被放大、正文又被放大一次。这里把内核缩放
    直接关掉，缩放权完全交给前端 JS。

    必须在 UI 线程访问 CoreWebView2（跨线程读 Settings 会抛
    InvalidOperationException: "CoreWebView2 can only be accessed from the UI
    thread"），所以取值和赋值都塞进 Form.Invoke 里做；窗口/内核可能还没就绪，
    失败就重试几次（调用方在后台线程里跑，不阻塞首屏）。
    """
    for _ in range(max(1, retries)):
        try:
            import webview
            from System import Action

            win = webview.windows[0] if webview.windows else None
            form = getattr(win, "native", None)
            ctl = getattr(form, "webview", None)
            if ctl is None:
                time.sleep(0.15)
                continue
            state = {}

            def _apply():
                try:
                    core = ctl.CoreWebView2
                    if core is None:
                        state["err"] = "core not ready"
                        return
                    core.Settings.IsZoomControlEnabled = False
                    state["ok"] = (core.Settings.IsZoomControlEnabled is False)
                except Exception as e:
                    state["err"] = repr(e)

            try:
                form.Invoke(Action(_apply))
            except Exception as e:
                state["err"] = "invoke:" + repr(e)
            if state.get("ok"):
                return True
        except Exception:
            pass
        time.sleep(0.15)
    return False


def _init_hub(first_payload, session_payloads=None):
    time.sleep(0.4)
    # 补上最小化框 + 缩放框，并接管 NCCALCSIZE/NCHITTEST（任务栏最小化、自由缩放）
    try:
        enable_window_features()
    except Exception:
        pass
    # 关掉内核整页缩放（后台线程，失败了也不影响使用：前端的 preventDefault 仍在）
    try:
        threading.Thread(target=_disable_native_zoom, daemon=True).start()
    except Exception:
        pass
    # 先恢复上次会话的标签：与“运行时双击新标签”走同一通道（js_inject →
    # evaluate_js，即 ExecuteScriptAsync），彻底避开 NavigateToString 的 2MB 限制。
    for _p in (session_payloads or []):
        try:
            js_inject(_p)
        except Exception:
            pass
    if first_payload:
        js_inject(first_payload)


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


def _show_version():
    """只报版本号，不开窗口。控制台版直接打印；窗口版（--noconsole）弹个小框。"""
    msg = "PureMD %s" % APP_VERSION
    try:
        if sys.stdout is not None:
            sys.stdout.write(msg + "\n")
            sys.stdout.flush()
            return
    except Exception:
        pass
    try:
        ctypes.windll.user32.MessageBoxW(
            0, msg + "\n\n（程序里点「帮助 → 关于」也能看到版本号）",
            "PureMD", 0x40)          # MB_ICONINFORMATION
    except Exception:
        pass


def main():
    # --version / -v：报版本后直接退出，绝不把它当成文件路径
    try:
        _args = sys.argv[1:]
    except Exception:
        _args = []
    if any(a in ("--version", "-v", "/version") for a in _args):
        _show_version()
        return
    if any(a in ("--help", "-h", "/?") for a in _args):
        text = ("PureMD %s — 只读 .md 预览器\n\n"
                "用法：PureMD.exe [选项] <文件.md>\n\n"
                "  --version, -v     显示版本号\n"
                "  --setup           安装/修复 .md 文件关联\n"
                "  --selftest        自检（资源、导出）\n" % APP_VERSION)
        try:
            if sys.stdout is not None:
                sys.stdout.write(text)
                sys.stdout.flush()
                return
        except Exception:
            pass
        try:
            ctypes.windll.user32.MessageBoxW(0, text, "PureMD", 0x40)
        except Exception:
            pass
        return
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
                "· 默认只读、不主动改动你的 .md 源文件\n"
                "（只有编辑模式里手动点保存才写回，且只改动过的段落）\n"
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
