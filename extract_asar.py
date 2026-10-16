import sys, json, math

asar = r"D:\00. APP\workbuddy\resources\app.asar"
with open(asar, "rb") as f:
    head = f.read(16)
    # offset 4: uint32 headerSize (小端)
    header_size = int.from_bytes(head[4:8], "little")
    header_bytes = f.read(header_size)
    text = header_bytes.decode("utf-8")
    header, _ = json.JSONDecoder().raw_decode(text)
data_start = math.ceil((16 + header_size) / 4096) * 4096

paths = []
def walk(node, p):
    for name, info in node.items():
        full = (p + "/" + name) if p else name
        if isinstance(info, dict) and "files" in info:
            walk(info["files"], full)
        else:
            paths.append(full)

walk(header.get("files", {}), "")

# 筛选预览/渲染相关 css/js
import re
pat = re.compile(r"(markdown|preview|highlight|marked|markdownit).*\.(css|js)$", re.I)
hits = [p for p in paths if pat.search(p)]
print("TOTAL_ASAR_FILES", len(paths))
print("HITS", len(hits))
for h in hits:
    print(h)

if len(sys.argv) >= 3:
    target = sys.argv[1]
    out = sys.argv[2]
    # 在 header.files 树里定位
    cur = header["files"]
    info = None
    for part in target.split("/"):
        if part not in cur:
            print("NOT_FOUND", target); sys.exit(0)
        info = cur[part]
        if "files" in info:
            cur = info["files"]
    off = data_start + int(info.get("offset", 0))
    size = int(info.get("size", 0))
    with open(asar, "rb") as f:
        f.seek(off)
        data = f.read(size)
    with open(out, "wb") as o:
        o.write(data)
    print("EXTRACTED", target, "->", out, size)
