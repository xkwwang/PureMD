# PureMD — Markdown 纯预览（双击即看，绝不改写源文件）

一个为「只读预览」而生的轻量 Markdown 查看器。核心保证：**双击 `.md` 即渲染，全程只读，绝不改动你的源文件，也不在磁盘上生成任何额外文件**（预览内容全部在内存里）。

## 特性

- **零改写**：源 `.md` 始终只读，写论文 / 读文献时绝不担心被后台编辑软件改坏。
- **三主题**：亮色 / 暗色 / 米色（护眼纸感），关闭后记住你的选择。
- **多标签 + 拖拽排序**：一行最多 6 个标签，多了出现左右箭头；标签可自由拖动排序；`☰` 按钮垂直列出全部标签。
- **自动刷新**：开启后，源文件被外部编辑保存时自动重新渲染当前页。
- **打开 / 重新加载 / 导出 PDF**：可一次打开多个文件；导出走系统打印对话框，选「另存为 PDF」即可存档。
- **GBK / 中文编码自动识别**：老文档常用 GBK/GB18030 也能正常显示，不乱码。
- **反馈**：点一下弹出作者邮箱，可一键复制后手动发邮件（不自动打开邮件客户端）。
- **完全离线、单文件分发**：打包后只有一个文件夹，发给朋友解压即用，无需安装、无需联网。

## 下载与安装

1. 到 [Releases](../../releases) 下载 `PureMD.zip`，解压到任意位置（如 `D:\Tools\PureMD`）。
   > 解压后**不要拆散 `PureMD` 文件夹**，它里面的 `_internal` 是运行依赖。
2. 双击 `关联 md 文件.cmd` → 看到「已设置成功」。
3. 之后双击任意 `.md` 即可预览；也可以直接把 `.md` 拖到 `PureMD.exe` 图标上打开（无需关联）。

### 系统依赖

需要 **Microsoft Edge WebView2 Runtime**。Win11 与较新的 Win10 已自带；若打开时提示「WebView2 启动失败」，请安装
[WebView2 Runtime](https://developer.microsoft.com/microsoft-edge/webview2/)（免费）。

## 从源码构建

```bash
# 1) 准备 Python 环境（建议 3.10+，并用虚拟环境）
python -m venv venv && venv\Scripts\activate
pip install pyinstaller pywebview

# 2) 打包成单目录 exe（产物在 md_preview_app/PureMD/）
pyinstaller --noconsole --onedir --name PureMD ^
  --icon mdpreview/md_preview.ico ^
  --add-data "mdpreview;mdpreview" ^
  --collect-data webview --collect-submodules webview ^
  --hidden-import webview.platforms.edgechromium ^
  open_md_preview.pyw
```

也可以直接双击仓库里的 `rebuild_exe.cmd`（会跑上面的命令并做自检）。

## 目录结构

```
PureMD/
├── open_md_preview.pyw      # 主程序（启动器 + 单实例端口锁 + 文件对话框 + 配置/反馈接口）
├── mdpreview/
│   ├── hub.html             # 多标签宿主界面（顶栏 / 标签栏 / 主题 / 反馈弹窗）
│   ├── template.html        # Markdown 渲染模板
│   ├── wb_overrides.css      # WorkBuddy 同款排版覆盖层
│   └── md_preview.ico        # 应用图标
├── rebuild_exe.cmd          # 一键重建 exe
├── make_zip.py              # 打包分发 zip
├── setup_md_preview.pyw     # 注册表关联脚本（exe 自带 --setup，一般不用直接跑）
├── LICENSE
├── .gitignore
└── README.md
```

## 配置

用户配置（主题、是否默认开启自动刷新）保存在：

```
%APPDATA%/PureMD/config.json
```

## 协议

[MIT](LICENSE) © 2026 xkwwang@126.com
