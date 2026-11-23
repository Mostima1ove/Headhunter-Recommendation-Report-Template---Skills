#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""猎头候选人推荐报告生成器。

从零构建 docx，不依赖任何二进制模板文件。换公司只需要改 config 里的
公司名 / logo / 页脚声明，其余版式由本脚本固定。

用法::

    python report.py --config config.json --data data.json --out 张三-XX基金-量化研究员.docx

依赖: python-docx
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# ---------------------------------------------------------------------------
# 版式默认值。config 里的 "layout" 可以逐项覆盖（见 inspect_template.py 的产出）。
# ---------------------------------------------------------------------------
DEFAULT_LAYOUT = {
    "page_w_cm": 21.0,          # A4 宽
    "page_h_cm": 29.7,          # A4 高
    "margin_l_cm": 3.17,
    "margin_r_cm": 3.01,
    "margin_t_cm": 2.54,
    "margin_b_cm": 2.54,
    "header_dist_cm": 1.9,
    "footer_dist_cm": 1.9,
    "font_east": "宋体",
    "font_ascii": "Calibri",
    "font_size_pt": 10.5,       # 五号
    "line_pt": None,            # None = 默认单倍行距；填固定值会压页
    "bullet": "■",              # 条目符号，勿用 ⚫（多数中文字体没有，会掉成黑块）
    "tab_field_cm": 5.3,        # 字段名 -> 值
    "tab_edu_cm": 4.46,         # 日期 -> 学校
    "tab_job_cm": 5.47,         # 日期 -> 公司
    "hang_cm": 0.74,            # 悬挂缩进
}

# 下列模块级常量由 apply_layout() 按 config 覆盖
PAGE_W, PAGE_H = Cm(21.0), Cm(29.7)
MARGIN_L, MARGIN_R = Cm(3.17), Cm(3.01)
MARGIN_T, MARGIN_B = Cm(2.54), Cm(2.54)
HEADER_DIST, FOOTER_DIST = Cm(1.9), Cm(1.9)

BODY_SIZE = 10.5
LINE_PT = None
EAST_FONT, ASCII_FONT = "宋体", "Calibri"
BULLET = "■"

TAB_FIELD = Cm(5.3)
TAB_EDU = Cm(4.46)
TAB_JOB = Cm(5.47)
HANG = Cm(0.74)


def apply_layout(cfg):
    """把 config["layout"] 里给的项盖到默认值上，并写回模块级常量。"""
    lay = dict(DEFAULT_LAYOUT)
    for k, v in (cfg.get("layout") or {}).items():
        if k in lay and v not in (None, ""):
            lay[k] = v
    globals().update({
        "PAGE_W": Cm(lay["page_w_cm"]),
        "PAGE_H": Cm(lay["page_h_cm"]),
        "MARGIN_L": Cm(lay["margin_l_cm"]),
        "MARGIN_R": Cm(lay["margin_r_cm"]),
        "MARGIN_T": Cm(lay["margin_t_cm"]),
        "MARGIN_B": Cm(lay["margin_b_cm"]),
        "HEADER_DIST": Cm(lay["header_dist_cm"]),
        "FOOTER_DIST": Cm(lay["footer_dist_cm"]),
        "BODY_SIZE": float(lay["font_size_pt"]),
        "LINE_PT": lay["line_pt"],
        "EAST_FONT": lay["font_east"],
        "ASCII_FONT": lay["font_ascii"],
        "BULLET": lay["bullet"],
        "TAB_FIELD": Cm(lay["tab_field_cm"]),
        "TAB_EDU": Cm(lay["tab_edu_cm"]),
        "TAB_JOB": Cm(lay["tab_job_cm"]),
        "HANG": Cm(lay["hang_cm"]),
    })
    return lay

DASH = "—"


# ---------------------------------------------------------------------------
# 基础工具
# ---------------------------------------------------------------------------
def _style_run(run, bold: bool = False, size: float = BODY_SIZE):
    run.font.name = ASCII_FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.append(fonts)
    fonts.set(qn("w:ascii"), ASCII_FONT)
    fonts.set(qn("w:hAnsi"), ASCII_FONT)
    fonts.set(qn("w:eastAsia"), EAST_FONT)
    return run


def add_para(doc, text="", *, bold=False, size=BODY_SIZE, left=None,
             first=None, tab=None, line=LINE_PT, keep_together=False,
             keep_with_next=False, space_before=None):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    if left is not None:
        pf.left_indent = left
    if first is not None:
        pf.first_line_indent = first
    if line:
        pf.line_spacing = Pt(line)
    if space_before is not None:
        pf.space_before = Pt(space_before)
    pf.space_after = Pt(0)
    if tab is not None:
        pf.tab_stops.add_tab_stop(tab, WD_TAB_ALIGNMENT.LEFT)
    pf.keep_together = keep_together
    pf.keep_with_next = keep_with_next
    if text:
        _style_run(p.add_run(text), bold=bold, size=size)
    return p


def blank(doc, count=1):
    for _ in range(count):
        add_para(doc, "")


def section_title(doc, text):
    return add_para(doc, text, bold=True, keep_with_next=True)


def field(doc, label, value):
    """基本信息那一行：字段名 + 制表位 + 值。"""
    p = add_para(doc, "", tab=TAB_FIELD, keep_together=True)
    _style_run(p.add_run(f"{label}："), bold=True)
    _style_run(p.add_run("\t"), bold=True)
    _style_run(p.add_run(value if value else DASH), bold=True)
    return p


def bullet(doc, text):
    """条目标题 + 正文，悬挂缩进，整条不跨页。"""
    p = add_para(doc, "", left=HANG, first=-HANG, keep_together=True)
    _style_run(p.add_run(f"{BULLET} {text}"))
    return p


def reason(doc, text):
    p = add_para(doc, "", left=HANG, first=-HANG, keep_together=True)
    _style_run(p.add_run(f"{BULLET} {text}"))
    return p


# ---------------------------------------------------------------------------
# 页眉页脚
# ---------------------------------------------------------------------------
def build_header(doc, cfg):
    header = doc.sections[0].header
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    logo = cfg.get("logo")
    if logo and Path(logo).exists():
        p.add_run().add_picture(logo, width=Cm(cfg.get("logo_width_cm", 3.2)))
    elif cfg.get("company"):
        _style_run(p.add_run(cfg["company"]), bold=True, size=12)


def build_footer(doc, cfg):
    footer = doc.sections[0].footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.space_before = Pt(4)
    pf.space_after = Pt(4)

    ppr = p._p.get_or_add_pPr()
    borders = OxmlElement("w:pBdr")
    for edge in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "4")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    ppr.append(borders)

    text = cfg.get("footer")
    if text:
        _style_run(p.add_run(text), size=10.5)


# ---------------------------------------------------------------------------
# 正文
# ---------------------------------------------------------------------------
def build(doc, cfg, data):
    apply_layout(cfg)
    sec = doc.sections[0]
    sec.page_width, sec.page_height = PAGE_W, PAGE_H
    sec.left_margin, sec.right_margin = MARGIN_L, MARGIN_R
    sec.top_margin, sec.bottom_margin = MARGIN_T, MARGIN_B
    sec.header_distance, sec.footer_distance = HEADER_DIST, FOOTER_DIST

    normal = doc.styles["Normal"]
    normal.font.name = ASCII_FONT
    normal.font.size = Pt(BODY_SIZE)
    rpr = normal.element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.append(fonts)
    fonts.set(qn("w:eastAsia"), EAST_FONT)

    build_header(doc, cfg)
    build_footer(doc, cfg)

    # 抬头四行
    head = [
        ("客户名称：", data.get("client")),
        ("推荐职位：", data.get("position")),
        ("工作地点：", data.get("location")),
        ("推荐日期：", data.get("date")),
    ]
    for label, value in head:
        p = add_para(doc, "", tab=TAB_FIELD, keep_together=True)
        _style_run(p.add_run(label), bold=True)
        _style_run(p.add_run("\t" + (value or DASH)), bold=True)
    blank(doc)

    # 推荐理由
    section_title(doc, "推荐理由")
    for r in data.get("reasons", [])[:5]:
        reason(doc, r)
    blank(doc, 2)

    # 基本信息
    section_title(doc, "基本信息")
    basic = data.get("basic", {})
    for label, key in [
        ("姓    名", "name"), ("性    别", "sex"), ("出生年月", "birth"),
        ("家庭地址", "address"), ("希望工作地", "city"), ("国    籍", "nation"),
        ("学    历", "education"), ("专业资质", "cert"), ("手    机", "phone"),
        ("邮    箱", "email"),
    ]:
        field(doc, label, basic.get(key))
    blank(doc, 2)

    # 教育背景
    section_title(doc, "教育背景")
    for edu in data.get("education", []):
        p = add_para(doc, "", tab=TAB_EDU, keep_together=True, keep_with_next=True)
        _style_run(p.add_run(edu["date"]), bold=True)
        _style_run(p.add_run("\t" + edu["school"]), bold=True)
        p2 = add_para(doc, "", left=TAB_EDU, keep_together=True)
        _style_run(p2.add_run(edu["degree"]), bold=True)
    blank(doc)

    # 技能
    section_title(doc, "专业资质与技能")
    for s in data.get("skills", []):
        reason(doc, s)
    blank(doc)

    # 工作经历
    section_title(doc, "工作经验")
    for job in data.get("experience", []):
        p = add_para(doc, "", tab=TAB_JOB, keep_together=True, keep_with_next=True)
        _style_run(p.add_run(job["date"]), bold=True)
        _style_run(p.add_run("\t" + job["org"]), bold=True)
        add_para(doc, "主要职责与业绩：", bold=True, keep_with_next=True)
        for item in job.get("items", []):
            bullet(doc, item)
        if job is not data.get("experience", [])[-1]:
            blank(doc)

    # 项目经历（应届生、项目型候选人；不填就不出这一节）
    projects = data.get("projects") or []
    if projects:
        blank(doc)
        section_title(doc, "项目经历")
        for pr in projects:
            p = add_para(doc, "", tab=TAB_JOB, keep_together=True, keep_with_next=True)
            _style_run(p.add_run(pr["date"]), bold=True)
            _style_run(p.add_run("\t" + pr["name"]), bold=True)
            for item in pr.get("items", []):
                bullet(doc, item)
            if pr is not projects[-1]:
                blank(doc)

    return doc


def check_jd(data):
    """有 JD 就对着写；没有就按简历综合总结，别卡住流程。"""
    if not (data.get("position") or "").strip():
        raise SystemExit(
            "缺少职位名（position）——不知道推什么岗，报告写不了。先向用户要职位全称。"
        )
    if not (data.get("jd") or "").strip():
        print(
            "提示：这次没有岗位介绍（jd），推荐理由按简历综合总结。"
            "能拿到 JD 的话，对着职责逐条写会准得多。",
            file=sys.stderr,
        )


def main(argv=None):
    ap = argparse.ArgumentParser(description="猎头候选人推荐报告生成器")
    ap.add_argument("--config", required=True, help="公司配置 json")
    ap.add_argument("--data", required=True, help="候选人数据 json")
    ap.add_argument("--out", required=True, help="输出 docx 路径")
    args = ap.parse_args(argv)

    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
    data = json.loads(Path(args.data).read_text(encoding="utf-8"))

    check_jd(data)

    doc = Document()
    build(doc, cfg, data)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    try:
        doc.save(out)
    except PermissionError:
        alt = out.with_name(out.stem + "-v2" + out.suffix)
        doc.save(alt)
        print(f"目标文件被占用，改存 {alt}", file=sys.stderr)
        out = alt
    print(out)


if __name__ == "__main__":
    main()
