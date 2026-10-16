# -*- coding: utf-8 -*-
r"""PureMD 专用：Markdown → .docx（Word）转换器。

设计原则：**解析行为与浏览器端预览完全一致**。
预览用的是 markdown-it（JS）默认预设 + {html:true, linkify:true, typographer:true}；
这里用 markdown-it-py（同一引擎的 Python 移植）配同样参数，保证
GFM 表格 / 删除线 / 自动链接 / 智能标点 的处理结果和预览一致。
（任务列表、数学公式 预览端也未启用插件，这里同样按纯文本处理。）

依赖：markdown-it-py[linkify]、python-docx（均随 PureMD venv 打包）。
"""
import base64
import io
import os
import re
from html import unescape

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from markdown_it import MarkdownIt

# ---------------------------------------------------------------- 常量
CODE_FONT = "Consolas"                 # 代码字体（等宽，不随正文方案）
LATIN_FONT = "Times New Roman"         # 西文/数字统一字体
HEADING_EA_FONT = "黑体"               # 中文标题字体
BODY_EA_FONT = "宋体"                  # 中文正文字体
INLINE_CODE_FILL = "F0F1F3"            # 行内代码底色（GitHub 风）
CODE_BLOCK_FILL = "F6F8FA"             # 代码块底色
TABLE_HEAD_FILL = "EDEFF2"             # 表头底色
HR_COLOR = "D0D7DE"                    # 分割线颜色
LINK_HEX = "0969DA"                    # 链接颜色
QUOTE_GRAY = RGBColor(0x57, 0x60, 0x6A)
PLACEHOLDER_GRAY = RGBColor(0x8B, 0x94, 0x9E)
IMG_MAX_CM = 15.5                      # 图片最大宽度（A4 减去页边距）
CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")
TAG_RE = re.compile(r"<[^>]+>")
DATA_URI_RE = re.compile(r"^data:[^;,]*(;base64)?,(.*)$", re.S)
SCHEME_RE = re.compile(r"^[a-zA-Z][a-zA-Z0-9+.\-]*:")


# ---------------------------------------------------------------- 底层 XML 助手
def _set_run_fonts(run, ascii_font, ea_font=None):
    """设置 run 的西文字体（ascii/hAnsi）与中文字体（eastAsia）。"""
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), ascii_font)
    rFonts.set(qn("w:hAnsi"), ascii_font)
    if ea_font:
        rFonts.set(qn("w:eastAsia"), ea_font)


def _set_style_fonts(style, latin_font, ea_font):
    """给样式显式设置西文+中文字体。

    必须先删除 rFonts 上的主题属性（asciiTheme/hAnsiTheme/eastAsiaTheme/cstheme）：
    OOXML 规范里主题 属性会 supersede 同名显式属性，不删的话设了也白设。
    """
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)
    for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if rFonts.get(qn(attr)) is not None:
            del rFonts.attrib[qn(attr)]
    rFonts.set(qn("w:ascii"), latin_font)
    rFonts.set(qn("w:hAnsi"), latin_font)
    rFonts.set(qn("w:cs"), latin_font)
    rFonts.set(qn("w:eastAsia"), ea_font)


def _shade_run(run, fill):
    rPr = run._r.get_or_add_rPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    rPr.append(shd)


def _shade_paragraph(p, fill):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def _shade_cell(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def _bottom_border(p, color=HR_COLOR, sz="6"):
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), sz)
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), color)
    pBdr.append(bottom)
    pPr.append(pBdr)


def _strip_tags(s):
    return unescape(TAG_RE.sub("", s or ""))


def _clean(s):
    """去掉 docx 不允许的控制字符。"""
    return CTRL_RE.sub("", s or "")


def _add_hyperlink(paragraph, url, text, bold=False, italic=False, strike=False):
    """python-docx 没有超链接 API，手工拼 w:hyperlink（外链关系 + 蓝色下划线）。"""
    part = paragraph.part
    r_id = part.relate_to(url, RT.HYPERLINK, is_external=True)
    hl = OxmlElement("w:hyperlink")
    hl.set(qn("r:id"), r_id)
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    c = OxmlElement("w:color")
    c.set(qn("w:val"), LINK_HEX)
    rPr.append(c)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), "single")
    rPr.append(u)
    if bold:
        rPr.append(OxmlElement("w:b"))
    if italic:
        rPr.append(OxmlElement("w:i"))
    if strike:
        rPr.append(OxmlElement("w:strike"))
    r.append(rPr)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = _clean(text)
    r.append(t)
    hl.append(r)
    paragraph._p.append(hl)


# ---------------------------------------------------------------- 文档骨架
def _setup_document(doc):
    """A4 + 统一字体规范 + 中性色标题。

    字体（中文公文惯例）：
      - 中文标题 → 黑体；中文正文 → 宋体
      - 英文与数字一律 Times New Roman（标题、正文、表格、超链接均走此规则）
      - 代码块 / 行内代码例外：保持 Consolas 等宽，不随正文方案
    """
    normal = doc.styles["Normal"]
    normal.font.size = Pt(11)
    _set_style_fonts(normal, LATIN_FONT, BODY_EA_FONT)
    for lvl in range(1, 7):
        try:
            st = doc.styles["Heading %d" % lvl]
            st.font.color.rgb = RGBColor(0x1F, 0x23, 0x28)
            _set_style_fonts(st, LATIN_FONT, HEADING_EA_FONT)
        except KeyError:
            pass
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.5)
    sec.top_margin = sec.bottom_margin = Cm(2.5)


def _make_md():
    """与浏览器端 template.html 的 markdownit(...) 参数保持一致。"""
    md = MarkdownIt("default")
    md.options["html"] = True
    md.options["linkify"] = True
    md.options["typographer"] = True
    md.options["breaks"] = False
    return md


# ---------------------------------------------------------------- 转换器
class _Converter:
    def __init__(self, doc, base_dir=None):
        self.doc = doc
        self.base_dir = base_dir or ""
        # 块级状态
        self.para = None            # 当前段落 {style, segs, prefix, hanging, quote}
        self.pending_style = None   # heading_open 记下，等 paragraph_open 消费
        self.list_stack = []        # ["ul"/"ol", ...]
        self.ord_counters = {}      # {depth: 当前序号}
        self.quote_depth = 0
        self.item_prefix = None     # list_item_open 产生，paragraph_open 消费
        # 行内状态
        self.style = {"bold": False, "italic": False, "strike": False}
        self.link_stack = []
        # 表格状态
        self.table = None
        self.cell = None            # 当前单元格：[[seg,...], ...]（多个段落）
        self.row = None
        self.in_head = False

    # ---------------- 入口 ----------------
    def walk(self, tokens):
        for tok in tokens:
            getattr(self, "_t_" + tok.type, self._t_noop)(tok)

    def _t_noop(self, tok):
        pass

    # ---------------- 标题 / 段落 ----------------
    def _t_heading_open(self, tok):
        self._flush()                         # 防御：上一个块没关干净
        try:
            self.pending_style = "Heading %d" % int(tok.tag[1:])
        except Exception:
            self.pending_style = "Heading 1"

    def _t_heading_close(self, tok):
        # 注意：markdown-it 的标题内容是裸 inline（heading_open, inline, heading_close），
        # 没有 paragraph_open/close 包裹——必须在这里主动落笔，否则标题会与
        # 后续块（表格/引用/下一个标题）串段。
        self._flush()
        self.pending_style = None

    def _t_paragraph_open(self, tok):
        if self.cell is not None:            # 表格单元格内的段落
            self.cell.append([])
            return
        if self.para is not None:            # 防御：上一个没正常关闭
            self._flush()
        prefix = self.item_prefix
        self.item_prefix = None
        self.para = {
            "style": self.pending_style,
            "segs": [],
            "prefix": prefix,
            "hanging": prefix is not None,
            "quote": self.quote_depth > 0,
        }
        self.pending_style = None

    def _t_paragraph_close(self, tok):
        if self.cell is not None:
            return
        self._flush()

    def _flush(self):
        if not self.para:
            return
        info = self.para
        self.para = None
        p = self.doc.add_paragraph(style=info["style"])
        pf = p.paragraph_format
        left = 0.6 * len(self.list_stack) + 0.5 * self.quote_depth
        if info["hanging"]:
            pf.left_indent = Cm(left + 0.45)
            pf.first_line_indent = Cm(-0.45)
        elif left > 0:
            pf.left_indent = Cm(left)
        if info["prefix"]:
            p.add_run(info["prefix"])
        for seg in info["segs"]:
            self._emit(p, seg)
        if info["quote"]:
            for r in p.runs:
                r.font.italic = True
                r.font.color.rgb = QUOTE_GRAY

    # ---------------- 行内 ----------------
    def _t_inline(self, tok):
        if self.cell is not None:
            if not self.cell:
                self.cell.append([])
            target = self.cell[-1]
        else:
            if self.para is None:            # 行内容不在段落里（罕见）：兜底建段
                self._t_paragraph_open(None)
            target = self.para["segs"]
        for ch in tok.children or []:
            t = ch.type
            if t in ("text", "text_special"):
                target.append(dict(kind="text", text=_clean(ch.content),
                                   href=(self.link_stack[-1] if self.link_stack else None),
                                   **self.style))
            elif t == "code_inline":
                target.append(dict(kind="text", text=_clean(ch.content), code=True,
                                   href=None, **self.style))
            elif t in ("softbreak", "hardbreak"):
                target.append({"kind": "break"})
            elif t == "image":
                attrs = ch.attrs or {}
                target.append({"kind": "picture", "src": attrs.get("src", ""),
                               "alt": _clean(ch.content or "")})
            elif t == "strong_open":
                self.style["bold"] = True
            elif t == "strong_close":
                self.style["bold"] = False
            elif t == "em_open":
                self.style["italic"] = True
            elif t == "em_close":
                self.style["italic"] = False
            elif t == "s_open":
                self.style["strike"] = True
            elif t == "s_close":
                self.style["strike"] = False
            elif t == "link_open":
                self.link_stack.append((ch.attrs or {}).get("href", ""))
            elif t == "link_close":
                if self.link_stack:
                    self.link_stack.pop()
            elif t == "html_inline":
                txt = _strip_tags(ch.content)
                if txt:
                    target.append(dict(kind="text", text=_clean(txt),
                                       href=(self.link_stack[-1] if self.link_stack else None),
                                       **self.style))

    # ---------------- 列表 ----------------
    def _t_bullet_list_open(self, tok):
        self.item_prefix = None
        self.list_stack.append("ul")

    def _t_ordered_list_open(self, tok):
        self.item_prefix = None
        self.list_stack.append("ol")
        self.ord_counters[len(self.list_stack)] = 0

    def _t_bullet_list_close(self, tok):
        if self.list_stack:
            self.list_stack.pop()

    def _t_ordered_list_close(self, tok):
        if self.list_stack:
            self.ord_counters.pop(len(self.list_stack), None)
            self.list_stack.pop()

    def _t_list_item_open(self, tok):
        if self.list_stack and self.list_stack[-1] == "ol":
            d = len(self.list_stack)
            self.ord_counters[d] = self.ord_counters.get(d, 0) + 1
            self.item_prefix = "%d. " % self.ord_counters[d]
        else:
            self.item_prefix = "\u2022 "        # •

    def _t_list_item_close(self, tok):
        pass

    # ---------------- 引用 ----------------
    def _t_blockquote_open(self, tok):
        self.quote_depth += 1

    def _t_blockquote_close(self, tok):
        self.quote_depth = max(0, self.quote_depth - 1)

    # ---------------- 代码块 / 分割线 / 裸 HTML ----------------
    def _t_fence(self, tok):
        self._add_code_block(tok.content)

    _t_code_block = _t_fence

    def _add_code_block(self, content):
        if self.para is not None:
            self._flush()
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        left = 0.6 * len(self.list_stack) + 0.5 * self.quote_depth
        if left > 0:
            pf.left_indent = Cm(left)
        pf.space_before = Pt(2)
        pf.space_after = Pt(2)
        pf.line_spacing = 1.0
        _shade_paragraph(p, CODE_BLOCK_FILL)
        lines = content.split("\n")
        while lines and not lines[-1].strip():
            lines.pop()
        run = p.add_run()
        _set_run_fonts(run, CODE_FONT, BODY_EA_FONT)
        run.font.size = Pt(9)
        for i, ln in enumerate(lines):
            if i:
                run.add_break()
            run.add_text(_clean(ln))

    def _t_hr(self, tok):
        if self.para is not None:
            self._flush()
        p = self.doc.add_paragraph()
        _bottom_border(p)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)

    def _t_html_block(self, tok):
        if self.para is not None:
            self._flush()
        txt = _strip_tags(tok.content).strip()
        if txt:
            self.doc.add_paragraph(_clean(txt))

    # ---------------- 表格 ----------------
    def _t_table_open(self, tok):
        self._flush()                         # 表格前可能挂着未关的标题/段落
        self.table = {"rows": [], "header": False}
        self.in_head = False

    def _t_table_close(self, tok):
        self._build_table()
        self.table = None

    def _t_thead_open(self, tok):
        self.in_head = True
        if self.table is not None:
            self.table["header"] = True

    def _t_thead_close(self, tok):
        self.in_head = False

    _t_tbody_open = _t_noop
    _t_tbody_close = _t_noop

    def _t_tr_open(self, tok):
        self.row = []

    def _t_tr_close(self, tok):
        if self.table is not None and self.row is not None:
            self.table["rows"].append(self.row)
        self.row = None

    def _t_th_open(self, tok):
        self.cell = [[]]
        self._cell_align = self._parse_align(tok)

    def _t_td_open(self, tok):
        self.cell = [[]]
        self._cell_align = self._parse_align(tok)

    def _t_th_close(self, tok):
        self._close_cell(bold=True)

    def _t_td_close(self, tok):
        self._close_cell(bold=False)

    def _parse_align(self, tok):
        style = (tok.attrs or {}).get("style", "") or ""
        m = re.search(r"text-align:\s*(left|center|right)", style)
        return {"left": WD_ALIGN_PARAGRAPH.LEFT,
                "center": WD_ALIGN_PARAGRAPH.CENTER,
                "right": WD_ALIGN_PARAGRAPH.RIGHT}.get(m.group(1)) if m else None

    def _close_cell(self, bold):
        if self.row is None:
            self.cell = None
            return
        self.row.append({"paras": self.cell or [[]], "bold": bold,
                         "align": getattr(self, "_cell_align", None)})
        self.cell = None

    def _build_table(self):
        if not self.table or not self.table["rows"]:
            return
        rows = self.table["rows"]
        ncols = max(len(r) for r in rows)
        if ncols <= 0:
            return
        t = self.doc.add_table(rows=len(rows), cols=ncols)
        try:
            t.style = "Table Grid"
        except Exception:
            pass
        t.autofit = True
        for ri, row in enumerate(rows):
            for ci in range(ncols):
                cell = t.cell(ri, ci)
                data = row[ci] if ci < len(row) else {"paras": [[]], "bold": False, "align": None}
                first = True
                for cpara in data["paras"] or [[]]:
                    p = cell.paragraphs[0] if first else cell.add_paragraph()
                    first = False
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.space_before = Pt(0)
                    if data.get("align"):
                        p.alignment = data["align"]
                    for seg in cpara:
                        self._emit(p, seg)
                if ri == 0 and self.table["header"]:
                    _shade_cell(cell, TABLE_HEAD_FILL)
                if data.get("bold") or (ri == 0 and self.table["header"]):
                    for p in cell.paragraphs:
                        for r in p.runs:
                            r.font.bold = True

    # ---------------- 段落到 docx 的落笔 ----------------
    def _emit(self, p, seg):
        kind = seg.get("kind")
        if kind == "break":
            p.add_run().add_break()
            return
        if kind == "picture":
            self._emit_picture(p, seg)
            return
        text = seg.get("text", "")
        if not text:
            return
        if seg.get("href"):
            _add_hyperlink(p, seg["href"], text,
                           bold=seg.get("bold"), italic=seg.get("italic"),
                           strike=seg.get("strike"))
            return
        r = p.add_run(text)
        if seg.get("bold"):
            r.font.bold = True
        if seg.get("italic"):
            r.font.italic = True
        if seg.get("strike"):
            r.font.strike = True
        if seg.get("code"):
            _set_run_fonts(r, CODE_FONT, BODY_EA_FONT)
            r.font.size = Pt(10)
            _shade_run(r, INLINE_CODE_FILL)

    def _emit_picture(self, p, seg):
        src = seg.get("src") or ""
        alt = seg.get("alt") or ""
        data = self._load_image(src)
        if data:
            try:
                run = p.add_run()
                run.add_picture(io.BytesIO(data), width=Cm(IMG_MAX_CM))
                return
            except Exception:
                pass
        r = p.add_run(("[图片] " + alt) if alt else "[图片]")
        r.font.italic = True
        r.font.color.rgb = PLACEHOLDER_GRAY

    def _load_image(self, src):
        """把图片 src 变成字节：支持 data:URI、本地相对/绝对路径、file: URL。
        远程 http(s) 一律不下载（本工具完全离线）。"""
        try:
            m = DATA_URI_RE.match(src)
            if m:
                if m.group(1):
                    return base64.b64decode(m.group(2))
                from urllib.parse import unquote
                return unquote(m.group(2)).encode("utf-8")
            if SCHEME_RE.match(src):
                low = src.lower()
                if low.startswith("file:"):
                    from urllib.parse import unquote, urlparse
                    path = unquote(urlparse(src).path)
                    if path.startswith("/") and len(path) > 2 and path[2] == ":":
                        path = path[1:]           # /C:/x → C:/x
                    if os.path.isfile(path):
                        with open(path, "rb") as f:
                            return f.read()
                return None                          # http(s) 等：离线不下载
            cand = src if os.path.isabs(src) else os.path.join(self.base_dir, src)
            if os.path.isfile(cand):
                with open(cand, "rb") as f:
                    return f.read()
        except Exception:
            return None
        return None


# ---------------------------------------------------------------- 对外 API
def md_to_docx(text, base_dir=None):
    """把 Markdown 文本转成 .docx，返回 BytesIO（调用方负责落盘）。"""
    md = _make_md()
    tokens = md.parse(text or "")
    doc = Document()
    _setup_document(doc)
    conv = _Converter(doc, base_dir)
    conv.walk(tokens)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def convert_file(md_path, out_path=None):
    """命令行/调试用：直接转一个文件。返回输出路径。"""
    with open(md_path, "rb") as f:
        raw = f.read()
    for enc in ("utf-8", "gb18030", "big5"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = raw.decode("utf-8", errors="replace")
    buf = md_to_docx(text, base_dir=os.path.dirname(os.path.abspath(md_path)))
    if not out_path:
        out_path = os.path.splitext(md_path)[0] + ".docx"
    with open(out_path, "wb") as f:
        f.write(buf.getvalue())
    return out_path


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("usage: python md2docx.py <input.md> [output.docx]")
        sys.exit(1)
    out = convert_file(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
    print("saved:", out)
