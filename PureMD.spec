# -*- mode: python ; coding: utf-8 -*-
# PureMD 打包配置（精简版）
#
# 原则：「只装运行时真正会用到的东西」——把原先宽泛的 collect_submodules /
# 整个 mdpreview 目录打包，换成精确清单 + 显式排除，在保证功能不变的前提下
# 把安装包压到最小。运行时真正需要的资源只有 mdpreview 下这 3 个文件：
#   hub.html         —— 多标签宿主界面
#   template.html    —— Markdown 渲染模板（含 GitHub 风格样式）
#   wb_overrides.css —— WorkBuddy 同款排版覆盖层
# 其余（wb_markdown.css 是抽取 WorkBuddy 样式时的分析底稿、md_preview.ico 是
# 构建期图标）都不参与运行，因此不再打进包里。
from PyInstaller.utils.hooks import collect_data_files

datas = [
    ('mdpreview/hub.html', 'mdpreview'),
    ('mdpreview/template.html', 'mdpreview'),
    ('mdpreview/wb_overrides.css', 'mdpreview'),
]
# pywebview 的 js/、lib/（WebView2 / WinForms 运行时）是运行时必需的
datas += collect_data_files('webview')
# python-docx 的 default.docx 模板（Document() 会用到）
datas += collect_data_files('docx')

# 动态 import 的后端必须显式声明，否则 PyInstaller 静态分析抓不到：
#   winforms     —— Windows 上的窗口承载层（WinForms + WebView2）
#   edgechromium —— 浏览器控件后端
hiddenimports = [
    'webview.platforms.winforms',
    'webview.platforms.edgechromium',
]

# 精简排除项（逐个核对过，运行时确实用不到）：
#   lxml.*        —— python-docx 只用 lxml.etree（+ 其依赖的 _elementpath），
#                    objectify / sax / builder / html / isoschematron 都是独立
#                    编译扩展，合计约 3MB，剔除后 docx 导出功能不受影响。
#   setuptools / pkg_resources / pip / wheel —— PyInstaller 常见的元数据冗余。
#   tkinter / pytest / pydoc / doctest —— GUI 用的是 WinForms，不是 Tk；也不做测试。
excludes = [
    'lxml.objectify',
    'lxml.sax',
    'lxml.builder',
    'lxml.html',
    'lxml.isoschematron',
    'setuptools',
    'pkg_resources',
    'pip',
    'wheel',
    'tkinter',
    'pytest',
    'pydoc',
    'doctest',
]

a = Analysis(
    ['open_md_preview.pyw'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=2,          # 剔除 docstring 与 assert（已核对源码无 assert / 无 __doc__ 依赖）
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='PureMD',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['mdpreview/md_preview.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='PureMD',
)
