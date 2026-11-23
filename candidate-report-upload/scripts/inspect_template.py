#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从一份现成的 docx 报告里扒出版式，生成 report.py 能吃的 config 片段。

用户公司已经有一套报告模板时用这个：把他们的 docx 丢进来，产出的
"layout" 段直接塞进 config.json，之后 report.py 就按他们的版式出报告。

用法::

    python inspect_template.py --input 客户模板.docx --out my-config.json
    python inspect_template.py --input 客户模板.docx --logo-dir . --out my-config.json

会尽量扒：页面尺寸、页边距、页眉页脚距离、字体与字号、条目符号、
制表位、页脚保密声明、页眉图片。扒不到的留默认值，不硬猜。

依赖: python-docx
"""
from __future__ import annotations

import argparse
import json
import sys
import zipfile
from collections import Counter
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

# 常见条目符号，按出现次数挑最多的那个
BULLETS = "■□▪▫◆◇●○•·▶▸-–—*"
TEXT_EXT = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".emf", ".wmf"}


def emu2cm(v):
    return None if v is None else round(v / 360000.0, 2)


def style_font(doc):
    st = doc.styles["Normal"]
    east = ascii_ = None
    size = None
    try:
        rpr = st.element.get_or_add_rPr()
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is not None:
            east = fonts.get(qn("w:eastAsia"))
            ascii_ = fonts.get(qn("w:ascii")) or fonts.get(qn("w:hAnsi"))
    except Exception:
        pass
    if st.font.size is not None:
        size = round(st.font.size.pt, 1)
    return east, ascii_, size


def scan_runs(doc, limit=400):
    """正文里真实用到的字体、字号、条目符号。"""
    east = Counter()
    ascii_ = Counter()
    size = Counter()
    bullets = Counter()
    for i, p in enumerate(doc.paragraphs):
        if i > limit:
            break
        text = (p.text or "").strip()
        if text and text[0] in BULLETS:
            bullets[text[0]] += 1
        for r in p.runs:
            if not (r.text or "").strip():
                continue
            try:
                rpr = r._element.rPr
                if rpr is not None:
                    fonts = rpr.find(qn("w:rFonts"))
                    if fonts is not None:
                        if fonts.get(qn("w:eastAsia")):
                            east[fonts.get(qn("w:eastAsia"))] += 1
                        if fonts.get(qn("w:ascii")):
                            ascii_[fonts.get(qn("w:ascii"))] += 1
            except Exception:
                pass
            if r.font.size is not None:
                size[round(r.font.size.pt, 1)] += 1
    return east, ascii_, size, bullets


def scan_tabs(doc, limit=200):
    """段落里用到的制表位（cm），取最常用的三个。"""
    pos = Counter()
    for i, p in enumerate(doc.paragraphs):
        if i > limit:
            break
        try:
            for ts in p.paragraph_format.tab_stops:
                pos[emu2cm(ts.position)] += 1
        except Exception:
            continue
    return [x for x, _ in pos.most_common(3) if x]


def scan_indent(doc, limit=200):
    """悬挂缩进：left_indent 为正、first_line_indent 为负的段落。"""
    hang = Counter()
    for i, p in enumerate(doc.paragraphs):
        if i > limit:
            break
        pf = p.paragraph_format
        try:
            left = pf.left_indent
            first = pf.first_line_indent
        except Exception:
            continue
        if left and first and first < 0:
            hang[emu2cm(-first)] += 1
    return hang.most_common(1)[0][0] if hang else None


def footer_text(doc):
    texts = []
    for sec in doc.sections:
        try:
            root = sec.footer._element
        except Exception:
            continue
        for t in root.iter(qn("w:t")):
            if t.text:
                texts.append(t.text)
    return "".join(texts).strip()


def export_logo(src: Path, out_dir: Path):
    """页眉里的图片导出来，返回相对路径。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src) as z:
        names = [n for n in z.namelist() if n.startswith("word/media/")]
        medias = [n for n in names if Path(n).suffix.lower() in TEXT_EXT]
        if not medias:
            return ""
        picks = sorted(medias, key=lambda n: z.getinfo(n).file_size, reverse=True)
        pick = picks[0]
        dst = out_dir / ("logo" + Path(pick).suffix.lower())
        with z.open(pick) as f, open(dst, "wb") as o:
            o.write(f.read())
    return str(dst)


def main(argv=None):
    ap = argparse.ArgumentParser(description="从现成 docx 扒版式，生成 config 片段")
    ap.add_argument("--input", required=True, help="用户公司的报告模板 docx")
    ap.add_argument("--out", help="写出 config json 的路径；不填就打印到屏幕")
    ap.add_argument("--logo-dir", help="把页眉图片导到这个目录")
    args = ap.parse_args(argv)

    src = Path(args.input)
    if not src.exists():
        raise SystemExit(f"找不到文件：{src}")

    doc = Document(str(src))
    sec = doc.sections[0]

    st_east, st_ascii, st_size = style_font(doc)
    east_c, ascii_c, size_c, bullets = scan_runs(doc)
    tabs = scan_tabs(doc)
    hang = scan_indent(doc)

    layout = {
        "page_w_cm": emu2cm(sec.page_width),
        "page_h_cm": emu2cm(sec.page_height),
        "margin_l_cm": emu2cm(sec.left_margin),
        "margin_r_cm": emu2cm(sec.right_margin),
        "margin_t_cm": emu2cm(sec.top_margin),
        "margin_b_cm": emu2cm(sec.bottom_margin),
        "header_dist_cm": emu2cm(sec.header_distance),
        "footer_dist_cm": emu2cm(sec.footer_distance),
        "font_east": east_c.most_common(1)[0][0] if east_c else st_east,
        "font_ascii": ascii_c.most_common(1)[0][0] if ascii_c else st_ascii,
        "font_size_pt": size_c.most_common(1)[0][0] if size_c else st_size,
        "bullet": bullets.most_common(1)[0][0] if bullets else None,
        # 制表位在多数模板里是 Word 默认制表位（XML 里查不到），扒不到就沿用默认值
        "tab_field_cm": tabs[0] if len(tabs) > 0 else 5.3,
        "tab_edu_cm": tabs[1] if len(tabs) > 1 else 4.46,
        "tab_job_cm": tabs[2] if len(tabs) > 2 else 5.47,
        "hang_cm": hang,
    }

    cfg = {
        "company": "",
        "logo": export_logo(src, Path(args.logo_dir)) if args.logo_dir else "",
        "logo_width_cm": 3.2,
        "footer": footer_text(doc),
        "layout": {k: v for k, v in layout.items() if v not in (None, "")},
    }

    text = json.dumps(cfg, ensure_ascii=False, indent=2)
    if args.out:
        Path(args.out).write_text(text + "\n", encoding="utf-8")
        print(f"已写出 {args.out}", file=sys.stderr)
        miss = [k for k, v in layout.items() if v in (None, "")]
        if miss:
            print("这几项没扒到，走默认值：" + "、".join(miss), file=sys.stderr)
        if not tabs:
            print("制表位没扒到（模板用的是 Word 默认制表位），先按默认值给；对不齐就手调 tab_*_cm。",
                  file=sys.stderr)
    else:
        print(text)


if __name__ == "__main__":
    main()
