# -*- coding: utf-8 -*-
"""生成类似 WorkBuddy 的 Markdown 'M' 图标：圆角蓝青渐变底 + 白色文档 + 蓝色 M。"""
import math
from PIL import Image, ImageDraw, ImageFont

FONT = r"C:\Windows\Fonts\arialbd.ttf"
OUT = r"D:\00. APP\notepad++\mdpreview\md_preview.ico"

TOP = (59, 130, 246, 255)    # #3B82F6
BOT = (6, 182, 212, 255)     # #06B6D4
DOC = (255, 255, 255, 255)
MCOL = (37, 99, 235, 255)    # #2563EB
FOLD = (226, 232, 240, 255)  # 折角浅灰
SHADOW = (0, 0, 0, 55)


def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(4))


def rr(draw, box, radius, fill):
    draw.rounded_rectangle(box, radius=radius, fill=fill)


def draw_icon(size):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # 渐变背景
    for y in range(size):
        d.line([(0, y), (size, y)], fill=lerp(TOP, BOT, y / size))
    # 圆角遮罩：只保留圆角方形的渐变
    bg = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    dbg = ImageDraw.Draw(bg)
    m = max(1, int(size * 0.06))
    rr(dbg, [m, m, size - m, size - m], int(size * 0.20), (255, 255, 255, 255))
    img = Image.composite(img, Image.new("RGBA", (size, size), (0, 0, 0, 0)), bg)
    d = ImageDraw.Draw(img)
    # 文档阴影
    sh = max(1, int(size * 0.02))
    dm = int(size * 0.21)
    rr(d, [dm + sh, dm + sh, size - dm + sh, size - dm + sh], int(size * 0.07), SHADOW)
    # 白色文档
    doc_box = [dm, dm, size - dm, size - dm]
    rr(d, doc_box, int(size * 0.07), DOC)
    # 右上折角
    if size >= 32:
        f = int(size * 0.12)
        d.polygon([(size - dm - f, dm), (size - dm, dm), (size - dm, dm + f)], fill=FOLD)
    # M 字母居中
    fs = int(size * 0.50)
    font = ImageFont.truetype(FONT, fs)
    bbox = d.textbbox((0, 0), "M", font=font)
    w = bbox[2] - bbox[0]
    h = bbox[3] - bbox[1]
    x = (size - w) / 2 - bbox[0]
    y = (size - h) / 2 - bbox[1] + int(size * 0.01)
    d.text((x, y), "M", font=font, fill=MCOL)
    return img


sizes = [256, 128, 96, 64, 48, 40, 32, 24, 16]
imgs = [draw_icon(s) for s in sizes]
imgs[0].save(OUT, format="ICO", sizes=[(s, s) for s in sizes],
             append_images=imgs[1:])
print("saved", OUT)

from PIL import Image as I
with I.open(OUT) as t:
    print("frames:", t.n_frames)
