# -*- coding: utf-8 -*-
r"""把 dist_share\PureMD 打包成 zip（Compress-Archive 不稳定时的替代方案）。

路径全部相对本脚本所在仓库根，CI 也能用：
  - 默认源目录：<仓库>/dist_share/PureMD
  - 默认输出：  <仓库>/dist_share/PureMD.zip
  - 可用环境变量覆盖：PUREMD_ZIP_BASE / PUREMD_ZIP_SRC / PUREMD_ZIP_OUT
打包前会把 release-assets/ 里的发布源文件（启动器 .cmd、使用说明.txt）
自动复制进源目录，保证本地与 CI 产出的 zip 内容一致。
"""
import os, zipfile, sys, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
BASE = os.environ.get("PUREMD_ZIP_BASE", os.path.join(HERE, "dist_share"))
ROOT = os.environ.get("PUREMD_ZIP_SRC", os.path.join(BASE, "PureMD"))
OUT = os.environ.get("PUREMD_ZIP_OUT", os.path.join(BASE, "PureMD.zip"))
ASSETS = os.path.join(HERE, "release-assets")

# 把发布源文件折叠进待打包目录（单一事实来源 = release-assets/）
if os.path.isdir(ASSETS):
    os.makedirs(ROOT, exist_ok=True)
    for name in os.listdir(ASSETS):
        src = os.path.join(ASSETS, name)
        if os.path.isfile(src):
            shutil.copyfile(src, os.path.join(ROOT, name))

if not os.path.isdir(ROOT):
    print("NO_SRC_DIR: %s" % ROOT)
    sys.exit(1)
if os.path.exists(OUT):
    os.remove(OUT)

n = 0
skipped = []
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
    for dirpath, dirs, files in os.walk(ROOT):
        # 跳过备份/旧构建目录与运行时锁文件
        dirs[:] = [d for d in dirs if not d.startswith("_old")]
        for fn in files:
            if fn == ".hub.lock" or fn.startswith("_old"):
                continue
            full = os.path.join(dirpath, fn)
            arc = os.path.relpath(full, BASE)
            try:
                z.write(full, arc)
                n += 1
            except Exception as e:
                skipped.append("%s -> %r" % (arc, e))

print("ENTRIES=%d" % n)
print("SIZE_MB=%.1f" % (os.path.getsize(OUT) / 1048576.0))
if skipped:
    print("SKIPPED=%d" % len(skipped))
    for s in skipped[:10]:
        print("  " + s)
