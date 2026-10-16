# PureMD — Markdown 纯预览（双击即看，绝不改写源文件）

一个为「只读预览」而生的轻量 Markdown 查看器。核心保证：**双击 `.md` 即渲染，全程只读，绝不改动你的源文件，也不在磁盘上生成任何额外文件**（预览内容全部在内存里）。

## 特性

- **零改写**：源 `.md` 始终只读，写论文 / 读文献时绝不担心被后台编辑软件改坏。
- **三主题**：亮色 / 暗色 / 米色（护眼纸感），关闭后记住你的选择（首帧即正确，不闪屏）。
- **多标签 + 拖拽排序**：一行最多 6 个标签，多了出现左右箭头；标签可自由拖动排序；`☰` 按钮垂直列出全部标签，点空白处自动收起。
- **历史恢复**：关闭时开着的标签会被记住，下次打开自动恢复上次的页面与激活项。
- **自动刷新**：开启后，源文件被外部编辑保存时自动重新渲染当前页。
- **打开 / 重新加载 / 导出 PDF**：可一次打开多个文件；导出走系统打印对话框，选「另存为 PDF」即可存档。
- **导出 Word（.docx）**：一键把当前文档转成可编辑的 .docx——标题/列表/表格/代码块/粗斜体/删除线/链接/图片全部保留，解析规则与预览完全一致（同样的 markdown 引擎）。
- **GBK / 中文编码自动识别**：老文档常用 GBK/GB18030 也能正常显示，不乱码。
- **帮助菜单**：「帮助」下拉含 *反馈*（弹作者邮箱、一键复制，不自动打开邮件客户端）与 *关于*（显示版本号）。
- **完全离线、单文件分发**：打包后只有一个文件夹，发给朋友解压即用，无需安装、无需联网。

## 下载与安装

1. 到 [Releases](../../releases) 下载 `PureMD.zip`，解压到任意位置（如 `D:\Tools\PureMD`）。
   > 解压后**不要拆散 `PureMD` 文件夹**，它里面的 `_internal` 是运行依赖。
2. 双击 `关联 md 文件.cmd` → 看到「已设置成功」。
3. 之后双击任意 `.md` 即可预览；也可以直接把 `.md` 拖到 `PureMD.exe` 图标上打开（无需关联）。

### 系统依赖

需要 **Microsoft Edge WebView2 Runtime**。Win11 与较新的 Win10 已自带；若打开时提示「WebView2 启动失败」，请安装
[WebView2 Runtime](https://developer.microsoft.com/microsoft-edge/webview2/)（免费）。

## 验证 Release 真伪（Sigstore 防伪）

每个 Release 附带三个文件：`PureMD.zip`（本体）、`PureMD.zip.sig`（签名）、`PureMD.zip.pem`（证书）。
下载后三者放同一目录，安装 [cosign](https://github.com/sigstore/cosign/releases) 后执行
（把 `v1.0.0` 换成你要验证的版本号）：

```bash
# Windows PowerShell / CMD 可去掉续行符 \ 写成一行
cosign verify-blob \
  --certificate PureMD.zip.pem \
  --certificate-identity "https://github.com/xkwwang/PureMD/.github/workflows/build.yml@refs/tags/v1.0.0" \
  --certificate-oidc-issuer "https://token.actions.githubusercontent.com" \
  --signature PureMD.zip.sig \
  PureMD.zip
```

输出 `Verified OK` 即证明：该 zip 确实由本仓库的 GitHub Actions 构建产出，且发布后未被篡改。
（注意：`--certificate` 必须提供；只用 `--certificate-identity` 而不给证书文件，cosign 会报
"provide a key with --key or --sk, a certificate to verify against with --certificate ..." 错误。）

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
├── open_md_preview.pyw      # 主程序（启动器 + 单实例端口锁 + 文件对话框 + 配置/帮助接口 + 会话恢复注入）
├── md2docx.py               # Markdown → .docx 转换器（markdown-it-py + python-docx，参数与预览一致）
├── mdpreview/
│   ├── hub.html             # 多标签宿主界面（顶栏 / 标签栏 / 主题 / 帮助菜单 / 会话保存）
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

用户配置（主题、是否默认开启自动刷新、上次关闭时的标签页会话）保存在：

```
%APPDATA%/PureMD/config.json
```

其中 `session.files` / `session.active` 记录上次关闭时打开的标签，下次启动时自动恢复；
删掉这两项（或整个配置文件）即可清空历史，不影响其它功能。

## 协议

[MIT](LICENSE) © 2026 xkwwang@126.com
