#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
挑战杯大学生课外学术科技作品竞赛申报书 docx 生成器

格式标准：A4，页边距上下 2.54cm 左右 2.5cm；正文宋体小四 1.5 倍行距首行缩进 2 字符；
一级标题黑体三号居中；二级标题黑体小三左对齐；三级标题宋体四号加粗；表格宋体五号居中；
参考文献宋体五号；创新点表格 4 列（序号/类型/创新点描述/对比优势）。

三大类别：natural_science 自然科学类学术论文 / social_science 哲学社会科学类 / tech_invention 科技发明制作。

使用方式：
    python build.py --data data.json --out output.docx
    python build.py --demo --out demo.docx

JSON 字段定义详见 SKILL.md 第十章。
"""

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor


# ============================================================
# 字体与格式常量
# ============================================================

FONT_SONG = "宋体"
FONT_HEI = "黑体"
FONT_FANGSONG = "仿宋"
FONT_KAI = "楷体"
FONT_TIMES = "Times New Roman"

SIZE_ER = Pt(22)
SIZE_XIAO_ER = Pt(18)
SIZE_SAN = Pt(16)
SIZE_XIAO_SAN = Pt(15)
SIZE_SI = Pt(14)
SIZE_XIAO_SI = Pt(12)
SIZE_WU = Pt(10.5)
SIZE_XIAO_WU = Pt(9)

PAGE_WIDTH_CM = 21.0
PAGE_HEIGHT_CM = 29.7
MARGIN_TOP_BOTTOM_CM = 2.54
MARGIN_LEFT_RIGHT_CM = 2.5

CATEGORY_LABEL = {
    "natural_science": "自然科学类学术论文",
    "social_science": "哲学社会科学类社会调查报告和学术论文",
    "tech_invention": "科技发明制作",
}
APPLICANT_LABEL = {"individual": "个人项目", "collective": "集体项目"}


# ============================================================
# 工具函数
# ============================================================

def set_run_font(run, font_name: str = FONT_SONG,
                 font_size=SIZE_XIAO_SI, bold: bool = False,
                 color: Optional[RGBColor] = None) -> None:
    """设置 run 字体（中英文同步设置 eastAsia/ascii/hAnsi）"""
    run.font.name = font_name
    run.font.size = font_size
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)


def set_cell_font(cell, font_name: str = FONT_SONG,
                  font_size=SIZE_WU, bold: bool = False,
                  alignment=WD_ALIGN_PARAGRAPH.CENTER) -> None:
    """设置单元格内所有文字字体与对齐"""
    for paragraph in cell.paragraphs:
        paragraph.alignment = alignment
        pf = paragraph.paragraph_format
        pf.line_spacing = 1.25
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        for run in paragraph.runs:
            set_run_font(run, font_name=font_name, font_size=font_size, bold=bold)


def set_cell_text(cell, text: str, font_name: str = FONT_SONG,
                  font_size=SIZE_WU, bold: bool = False,
                  alignment=WD_ALIGN_PARAGRAPH.CENTER) -> None:
    """清空单元格并写入文字（含字体设置）"""
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = alignment
    pf = p.paragraph_format
    pf.line_spacing = 1.25
    run = p.add_run(str(text))
    set_run_font(run, font_name=font_name, font_size=font_size, bold=bold)


def add_paragraph_with_format(
    doc,
    text: str,
    font_name: str = FONT_SONG,
    font_size=SIZE_XIAO_SI,
    bold: bool = False,
    alignment=WD_ALIGN_PARAGRAPH.LEFT,
    first_line_indent: bool = True,
    line_spacing: float = 1.5,
    space_before: float = 0,
    space_after: float = 0,
):
    """添加带格式段落，可控制字体/字号/对齐/缩进/行距/段前后"""
    p = doc.add_paragraph()
    p.alignment = alignment
    pf = p.paragraph_format
    if first_line_indent:
        pf.first_line_indent = Pt(font_size.pt * 2)
    pf.line_spacing = line_spacing
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    run = p.add_run(text)
    set_run_font(run, font_name=font_name, font_size=font_size, bold=bold)
    return p


def add_heading_level1(doc, text: str):
    """一级标题：黑体三号，居中，段前段后 12pt"""
    return add_paragraph_with_format(
        doc, text,
        font_name=FONT_HEI, font_size=SIZE_SAN, bold=True,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_indent=False,
        space_before=12, space_after=12,
    )


def add_heading_level2(doc, text: str):
    """二级标题：黑体小三，左对齐，段前段后 6pt"""
    return add_paragraph_with_format(
        doc, text,
        font_name=FONT_HEI, font_size=SIZE_XIAO_SAN, bold=True,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=False,
        space_before=6, space_after=6,
    )


def add_heading_level3(doc, text: str):
    """三级标题：宋体四号加粗，左对齐"""
    return add_paragraph_with_format(
        doc, text,
        font_name=FONT_SONG, font_size=SIZE_SI, bold=True,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=False,
        space_before=6, space_after=3,
    )


def add_body_paragraph(doc, text: str, indent: bool = True):
    """正文段落：宋体小四，1.5 倍行距，首行缩进 2 字符"""
    return add_paragraph_with_format(
        doc, text,
        font_name=FONT_SONG, font_size=SIZE_XIAO_SI,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_indent=indent,
        line_spacing=1.5,
    )


def add_reference_paragraph(doc, text: str):
    """参考文献段落：宋体五号，单倍行距，悬挂缩进"""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf = p.paragraph_format
    pf.line_spacing = 1.25
    pf.left_indent = Pt(15)
    pf.first_line_indent = Pt(-15)
    pf.space_before = Pt(0)
    pf.space_after = Pt(2)
    run = p.add_run(text)
    set_run_font(run, font_name=FONT_SONG, font_size=SIZE_WU)
    return p


def add_table_from_data(doc, headers: List[str], rows: List[List[str]],
                        col_widths: Optional[List[float]] = None,
                        first_col_bold: bool = False):
    """从数据创建表格，自动应用规范格式（表头加粗居中、数据居中）"""
    n_cols = len(headers)
    n_rows = 1 + len(rows)
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        set_cell_text(hdr_cells[i], h, font_name=FONT_SONG,
                      font_size=SIZE_WU, bold=True)
        hdr_cells[i].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    for i, row in enumerate(rows, start=1):
        cells = table.rows[i].cells
        for j, val in enumerate(row):
            is_bold = first_col_bold and j == 0
            set_cell_text(cells[j], val, font_name=FONT_SONG,
                          font_size=SIZE_WU, bold=is_bold)
            cells[j].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(w)

    doc.add_paragraph()
    return table


def setup_page(doc):
    """设置 A4 页面与页边距（上下 2.54cm，左右 2.5cm）"""
    section = doc.sections[0]
    section.page_width = Cm(PAGE_WIDTH_CM)
    section.page_height = Cm(PAGE_HEIGHT_CM)
    section.top_margin = Cm(MARGIN_TOP_BOTTOM_CM)
    section.bottom_margin = Cm(MARGIN_TOP_BOTTOM_CM)
    section.left_margin = Cm(MARGIN_LEFT_RIGHT_CM)
    section.right_margin = Cm(MARGIN_LEFT_RIGHT_CM)


def add_page_number(doc):
    """页脚添加居中页码（宋体五号）"""
    section = doc.sections[0]
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)
    set_run_font(run, font_name=FONT_SONG, font_size=SIZE_WU)


def add_page_break(doc):
    """插入分页符"""
    p = doc.add_paragraph()
    run = p.add_run()
    run.add_break(WD_BREAK.PAGE)


def count_chinese_chars(text: str) -> int:
    """统计字数（中文字符+数字+英文字母+标点，不计空白；用于字数校验）"""
    return sum(1 for c in text if not c.isspace())


def safe_str(value, default=""):
    """安全转字符串，None 返回默认值"""
    if value is None:
        return default
    return str(value)


def join_paragraphs(paragraphs) -> str:
    """将段落列表合并为单个字符串（用于字数统计）"""
    if isinstance(paragraphs, str):
        return paragraphs
    if isinstance(paragraphs, list):
        return "".join(str(p) for p in paragraphs)
    return ""


def format_innovation_row(idx: int, inv: dict) -> List[str]:
    """格式化创新点表格行（序号/类型/创新点描述/对比优势）"""
    if not isinstance(inv, dict):
        return [str(idx), "", str(inv), ""]
    return [
        str(idx),
        inv.get("type", ""),
        inv.get("description", ""),
        inv.get("advantage", ""),
    ]


def format_reference(idx: int, ref: dict) -> str:
    """格式化参考文献条目（GB/T 7714）"""
    if not isinstance(ref, dict):
        return f"[{idx}] {ref}"
    authors, title = ref.get("authors", ""), ref.get("title", "")
    ref_type = ref.get("type", "J")
    source, year = ref.get("source", ""), ref.get("year", "")
    volume, issue = ref.get("volume", ""), ref.get("issue", "")
    pages, publisher = ref.get("pages", ""), ref.get("publisher", "")
    city = ref.get("city", "")

    if ref_type in ("J", "C"):
        vol_iss = f"{volume}({issue})" if volume and issue else (volume or issue)
        page_part = f": {pages}" if pages else ""
        return f"[{idx}] {authors}. {title}[{ref_type}]. {source}, {year}{', ' + vol_iss if vol_iss else ''}{page_part}."
    elif ref_type == "M":
        page_part = f": {pages}" if pages else ""
        return f"[{idx}] {authors}. {title}[M]. {city}: {publisher}, {year}{page_part}."
    elif ref_type == "D":
        return f"[{idx}] {authors}. {title}[D]. {city}: {source}, {year}."
    elif ref_type == "S":
        return f"[{idx}] {authors}. {title}: {source}[S]. {city}: {publisher}, {year}."
    elif ref_type == "EB/OL":
        url = ref.get("url", "")
        access = ref.get("access_date", "")
        return f"[{idx}] {authors}. {title}[EB/OL]. ({year})[{access}]. {url}."
    else:
        return f"[{idx}] {authors}. {title}[{ref_type}]. {source}, {year}."


# ============================================================
# ApplicationDocBuilder 主类
# ============================================================

class ApplicationDocBuilder:
    """挑战杯课外学术科技作品竞赛申报书 docx 构建器"""

    def __init__(self):
        self.doc = Document()
        setup_page(self.doc)
        add_page_number(self.doc)
        style = self.doc.styles["Normal"]
        style.font.name = FONT_SONG
        style.font.size = SIZE_XIAO_SI
        rPr = style._element.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        rFonts.set(qn("w:eastAsia"), FONT_SONG)
        rFonts.set(qn("w:ascii"), FONT_SONG)
        rFonts.set(qn("w:hAnsi"), FONT_SONG)
        self.data: Dict[str, Any] = {}
        self.warnings: List[str] = []

    def _get(self, *keys, default=""):
        """安全取嵌套字段，缺字段返回默认值"""
        cur = self.data
        for k in keys:
            if not isinstance(cur, dict):
                return default
            cur = cur.get(k, default)
        return cur if cur is not None else default

    def _get_list(self, *keys, default=None):
        """安全取嵌套 list 字段"""
        if default is None:
            default = []
        v = self._get(*keys, default=default)
        if isinstance(v, str):
            return [v]
        if isinstance(v, list):
            return v
        return default

    def add_h1(self, text):
        return add_heading_level1(self.doc, text)

    def add_h2(self, text):
        return add_heading_level2(self.doc, text)

    def add_h3(self, text):
        return add_heading_level3(self.doc, text)

    def add_para(self, text, indent=True):
        return add_body_paragraph(self.doc, text, indent=indent)

    def add_table(self, headers, rows, col_widths=None, first_col_bold=False):
        return add_table_from_data(self.doc, headers, rows, col_widths, first_col_bold)

    def add_page_break(self):
        add_page_break(self.doc)

    # --------------------------------------------------------
    # 封面
    # --------------------------------------------------------

    def _add_cover(self):
        """封面：黑体二号标题 + 副标题 + 5 行下划线信息"""
        for _ in range(2):
            self.doc.add_paragraph()

        title = "第十八届" + "「挑战杯」大学生课外学术科技作品竞赛申报书"
        add_paragraph_with_format(
            self.doc, title,
            font_name=FONT_HEI, font_size=SIZE_ER, bold=True,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            first_line_indent=False,
            space_before=12, space_after=12,
        )

        subtitle = "（课外学术科技作品）"
        add_paragraph_with_format(
            self.doc, subtitle,
            font_name=FONT_HEI, font_size=SIZE_SAN, bold=False,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            first_line_indent=False,
            space_after=24,
        )

        for _ in range(2):
            self.doc.add_paragraph()

        category_label = CATEGORY_LABEL.get(self._get("category"), "科技发明制作")
        applicant_label = APPLICANT_LABEL.get(self._get("applicant_type"), "集体项目")

        info_items = [
            ("作品全称", self._get("work_full_name")),
            ("作品类别", category_label),
            ("申报者代表", self._get("leader_name")),
            ("所在学校", self._get("school")),
            ("申报日期", self._get("apply_date")),
        ]
        for label, value in info_items:
            p = self.doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = p.paragraph_format
            pf.line_spacing = 2.0
            pf.space_before = Pt(6)
            pf.space_after = Pt(6)
            run_label = p.add_run(f"{label}：")
            set_run_font(run_label, font_name=FONT_HEI,
                         font_size=SIZE_SI, bold=True)
            run_value = p.add_run(value if value else "                    ")
            set_run_font(run_value, font_name=FONT_SONG, font_size=SIZE_SI)
            run_value.underline = True

        self.add_page_break()

    # --------------------------------------------------------
    # 作品全称 + 作品类别
    # --------------------------------------------------------

    def _add_work_info(self):
        """一、作品全称 + 二、作品类别"""
        self.add_h1("一、作品全称")
        work_name = self._get("work_full_name")
        if work_name:
            add_paragraph_with_format(
                self.doc, work_name,
                font_name=FONT_HEI, font_size=SIZE_XIAO_ER, bold=True,
                alignment=WD_ALIGN_PARAGRAPH.CENTER,
                first_line_indent=False,
                space_before=6, space_after=6,
            )
            if len(work_name) > 30:
                self.warnings.append(f"作品全称 {len(work_name)} 字超过 30 字上限")
        else:
            self.add_para("（请填写作品全称，不超过 30 字，突出『做什么+为谁做』，禁三层『基于...的...』堆砌。）")

        self.add_h1("二、作品类别")
        category = self._get("category")
        category_options = [
            ("natural_science", "自然科学类学术论文"),
            ("social_science", "哲学社会科学类社会调查报告和学术论文"),
            ("tech_invention", "科技发明制作"),
        ]
        for key, label in category_options:
            mark = "☑" if category == key else "☐"
            add_paragraph_with_format(
                self.doc, f"{mark} {label}",
                font_name=FONT_SONG, font_size=SIZE_XIAO_SI,
                alignment=WD_ALIGN_PARAGRAPH.LEFT,
                first_line_indent=False,
                line_spacing=1.5,
            )

    # --------------------------------------------------------
    # 申报者情况
    # --------------------------------------------------------

    def _add_applicant_info(self):
        """三、申报者情况（个人项目/集体项目两栏）"""
        self.add_h1("三、申报者情况")
        applicant_type = self._get("applicant_type", default="collective")
        applicant_label = APPLICANT_LABEL.get(applicant_type, "集体项目")

        self.add_h2(f"（{applicant_label}）")

        leader_rows = [
            ["姓名", self._get("leader_name")],
            ["学号", self._get("leader_id")],
            ["性别", self._get("leader_gender")],
            ["年龄", self._get("leader_age")],
            ["专业", self._get("leader_major")],
            ["年级", self._get("leader_grade")],
            ["学院", self._get("college")],
            ["学校", self._get("school")],
            ["联系电话", self._get("leader_phone")],
        ]
        self.add_table(["字段", "内容"], leader_rows, col_widths=[4.0, 12.0],
                       first_col_bold=True)

        if applicant_type == "collective":
            self.add_h2("（集体项目团队成员）")
            members = self._get_list("team_members")
            if members:
                rows = []
                for i, m in enumerate(members, 1):
                    if isinstance(m, dict):
                        rows.append([
                            str(i),
                            m.get("name", ""),
                            m.get("id", ""),
                            m.get("major", ""),
                            m.get("grade", ""),
                            m.get("role", ""),
                        ])
                    else:
                        rows.append([str(i), str(m), "", "", "", ""])
                self.add_table(
                    ["序号", "姓名", "学号", "专业", "年级", "分工"],
                    rows,
                    col_widths=[1.5, 2.0, 2.5, 3.0, 2.0, 5.0],
                )
            else:
                self.add_para("（请填写集体项目团队成员 2~7 人，每人含姓名/学号/专业/年级/分工，禁只列姓名不写分工。）")

        self.add_h2("（指导教师信息）")
        advisor_rows = [
            ["姓名", self._get("advisor_name")],
            ["职称", self._get("advisor_title")],
            ["研究方向", self._get("advisor_research")],
            ["联系电话", self._get("advisor_phone")],
        ]
        self.add_table(["字段", "内容"], advisor_rows, col_widths=[4.0, 12.0],
                       first_col_bold=True)

    # --------------------------------------------------------
    # 作品简介
    # --------------------------------------------------------

    def _add_abstract(self):
        """四、作品简介（≤500 字）"""
        self.add_h1("四、作品简介")
        abstract = self._get("abstract")
        if abstract:
            self.add_para(abstract)
            char_count = count_chinese_chars(abstract)
            if char_count > 500:
                self.warnings.append(f"作品简介 {char_count} 字超过 500 字上限")
        else:
            self.add_para("（请填写作品简介，≤500 字，按 4 句结构：痛点+做什么 / 怎么做+量化目标 / 产出什么 / 研究现状。）")

    # --------------------------------------------------------
    # 选题背景与意义
    # --------------------------------------------------------

    def _add_background(self):
        """五、选题背景与意义（800~1200 字，4 子节）"""
        self.add_h1("五、选题背景与意义")
        background = self._get_list("background")
        if background:
            for para in background:
                self.add_para(para)
            total = count_chinese_chars(join_paragraphs(background))
            if total < 800 or total > 1200:
                self.warnings.append(f"选题背景 {total} 字不在 800~1200 区间")
        else:
            self.add_h2("（一）时代背景")
            self.add_para("（请填写时代背景，150~200 字，3 句话讲政策/行业/学术趋势，必须含权威数据来源。）")
            self.add_h2("（二）现实痛点")
            self.add_para("（请填写现实痛点，300~400 字，2~3 个真实场景，必须可量化。）")
            self.add_h2("（三）国内外研究现状")
            self.add_para("（请填写研究现状，200~300 字，评述已有方案不足，引出本项目差异。）")
            self.add_h2("（四）项目意义")
            self.add_para("（请填写项目意义，150~300 字，理论/实践/社会三角度，至少两个。）")

    # --------------------------------------------------------
    # 研究方法与过程【重点】
    # --------------------------------------------------------

    def _add_research_method(self):
        """六、研究方法与过程（1000~1500 字，4 子节）【重点】"""
        self.add_h1("六、研究方法与过程")
        methods = self._get_list("research_method")
        if methods:
            for para in methods:
                self.add_para(para)
            total = count_chinese_chars(join_paragraphs(methods))
            if total < 1000 or total > 1500:
                self.warnings.append(f"研究方法 {total} 字不在 1000~1500 区间")
        else:
            self.add_h2("（一）文献调研")
            self.add_para("（请填写文献调研，150~200 字，覆盖篇数+时间范围+核心结论，不止『查阅了大量文献』。）")
            self.add_h2("（二）数据/样本采集")
            self.add_para("（请填写数据采集，300~400 字，来源+数量+采集方式+伦理审查。社科类必须含伦理审查编号。）")
            self.add_h2("（三）实验/调研/仿真设计")
            self.add_para("（请填写实验设计，350~500 字，变量控制+对照组+重复次数+工具型号+统计方法。）")
            self.add_h2("（四）实地测试/验证")
            self.add_para("（请填写实地测试，200~400 字，地点+时长+评估指标+误差分析。）")

    # --------------------------------------------------------
    # 研究结果与讨论【重点】
    # --------------------------------------------------------

    def _add_results_discussion(self):
        """七、研究结果与讨论（1500~2000 字，4 子节）【重点】"""
        self.add_h1("七、研究结果与讨论")
        results = self._get_list("results_discussion")
        if results:
            for para in results:
                self.add_para(para)
            total = count_chinese_chars(join_paragraphs(results))
            if total < 1500 or total > 2000:
                self.warnings.append(f"研究结果 {total} 字不在 1500~2000 区间")
        else:
            self.add_h2("（一）主要发现")
            self.add_para("（请填写主要发现，500~700 字，按重要性排序，每发现含数据+图表编号引用。）")
            self.add_h2("（二）与已有研究对比")
            self.add_para("（请填写对比，300~500 字，量化差异，如准确率从 80% 提升至 92%，p<0.01。）")
            self.add_h2("（三）结果讨论")
            self.add_para("（请填写讨论，400~500 字，理论/实践含义，可推广边界。）")
            self.add_h2("（四）局限性分析")
            self.add_para("（请填写局限性，200~400 字，数据/方法/样本代表性 2~3 条，必须真实可改进。）")

    # --------------------------------------------------------
    # 结论
    # --------------------------------------------------------

    def _add_conclusion(self):
        """八、结论（300~500 字）"""
        self.add_h1("八、结论")
        conclusion = self._get("conclusion")
        if conclusion:
            self.add_para(conclusion)
            total = count_chinese_chars(conclusion)
            if total < 300 or total > 500:
                self.warnings.append(f"结论 {total} 字不在 300~500 区间")
        else:
            self.add_para("（请填写结论，300~500 字，3 句话结构：核心发现+学术贡献+下一步计划。）")

    # --------------------------------------------------------
    # 创新点
    # --------------------------------------------------------

    def _add_innovations(self):
        """九、创新点（3~5 个，每个 50~100 字，表格呈现）"""
        self.add_h1("九、创新点")
        innovations = self._get_list("innovations")
        if innovations:
            rows = []
            for i, inv in enumerate(innovations, 1):
                rows.append(format_innovation_row(i, inv))
            self.add_table(
                ["序号", "类型", "创新点描述", "对比优势"],
                rows,
                col_widths=[1.5, 2.5, 6.0, 6.0],
            )
            if len(innovations) < 3 or len(innovations) > 5:
                self.warnings.append(f"创新点 {len(innovations)} 个不在 3~5 区间")
        else:
            self.add_para("（请填写创新点 3~5 个，每个 50~100 字。表格 4 列：序号/类型/创新点描述/对比优势。"
                          "必须含对比+量化。禁用『首次』『先进』『实现』等无支撑词。）")

    # --------------------------------------------------------
    # 参考文献
    # --------------------------------------------------------

    def _add_references(self):
        """十、参考文献（GB/T 7714 格式，5~15 条）"""
        self.add_h1("十、参考文献")
        references = self._get_list("references")
        if references:
            for i, ref in enumerate(references, 1):
                ref_text = format_reference(i, ref)
                add_reference_paragraph(self.doc, ref_text)
            if len(references) < 5 or len(references) > 15:
                self.warnings.append(f"参考文献 {len(references)} 条不在 5~15 区间")
        else:
            add_reference_paragraph(self.doc, "[1] 张三, 李四. 示例文献题名[J]. 期刊名, 2024, 48(3): 45-52.")
            add_reference_paragraph(self.doc, "（请按 GB/T 7714 格式补 5~15 条参考文献，含[J]/[C]/[M]/[D]/[S]/[EB/OL] 类型。）")

    # --------------------------------------------------------
    # 附录
    # --------------------------------------------------------

    def _add_appendix(self):
        """十一、附录（图表/问卷/原始数据/代码/实物照片）"""
        self.add_h1("十一、附录")
        appendix = self._get_list("appendix")
        if appendix:
            for i, item in enumerate(appendix, 1):
                if isinstance(item, dict):
                    title = item.get("title", f"附录 {i}")
                    content = item.get("content", "")
                    self.add_h3(f"附录 {i}：{title}")
                    if content:
                        self.add_para(content)
                else:
                    self.add_h3(f"附录 {i}")
                    self.add_para(str(item))
        else:
            self.add_h3("附录 1：图表清单")
            self.add_para("（请补充图表清单，含图 X / 表 Y 编号对应正文引用。）")
            self.add_h3("附录 2：问卷（社科类必备）")
            self.add_para("（请补充调研问卷，含题项数+回收份数+有效率。）")
            self.add_h3("附录 3：原始数据")
            self.add_para("（请补充原始数据字段说明+样本量。）")
            self.add_h3("附录 4：代码/算法（科技发明类必备）")
            self.add_para("（请补充 CNN 模型核心代码，含数据加载+模型定义+训练循环。）")
            self.add_h3("附录 5：实物照片（科技发明类 A 类必备）")
            self.add_para("（请补充实物照片，含 5 视图+尺寸标注。）")

    # --------------------------------------------------------
    # 三级评审意见
    # --------------------------------------------------------

    def _add_review_sections(self):
        """十二/十三/十四：指导教师/学校/省级或国家级评审意见"""
        review_specs = [
            ("十二、指导教师推荐意见", "指导教师签字", "指导教师签字"),
            ("十三、学校评审意见", "学校盖章", "学校盖章"),
            ("十四、省级/国家级评审意见", "评审委员会盖章", "评审委员会盖章"),
        ]
        for title, sign_label, _ in review_specs:
            self.add_h1(title)
            for _ in range(6):
                self.doc.add_paragraph()
            self.add_para(
                f"{sign_label}：____________________    日期：______年____月____日",
                indent=False,
            )

    # --------------------------------------------------------
    # 主构建方法
    # --------------------------------------------------------

    def build(self, data: Dict[str, Any], output_path: str) -> str:
        """主构建方法：编排 14 栏目，生成 docx"""
        try:
            self.data = data if isinstance(data, dict) else {}
            self._validate_data()

            self._add_cover()
            self._add_work_info()
            self._add_applicant_info()
            self._add_abstract()
            self._add_background()
            self._add_research_method()
            self._add_results_discussion()
            self._add_conclusion()
            self._add_innovations()
            self._add_references()
            self._add_appendix()
            self._add_review_sections()

            if self.warnings:
                print("⚠️ 数据校验警告：", file=sys.stderr)
                for w in self.warnings:
                    print(f"  - {w}", file=sys.stderr)

            return self._save(output_path)
        except Exception as e:
            sys.stderr.write(f"❌ 生成失败：{e}\n")
            raise

    def _save(self, output_path: str) -> str:
        """保存文档，自动创建目录"""
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        self.doc.save(str(output_path))
        print(f"✅ 申报书已生成：{output_path}")
        return str(output_path)

    # --------------------------------------------------------
    # 数据校验
    # --------------------------------------------------------

    def _validate_data(self) -> List[str]:
        """校验数据完整性，返回缺失字段列表（不阻断生成）"""
        p0_fields = [
            ("work_full_name", "作品全称"),
            ("category", "作品类别"),
            ("applicant_type", "申报者类型"),
            ("leader_name", "申报者代表姓名"),
            ("leader_id", "学号"),
            ("college", "学院"),
            ("school", "学校"),
        ]
        for key, name in p0_fields:
            if not self._get(key):
                self.warnings.append(f"缺少 {name}（{key}）")

        if self._get("category") and self._get("category") not in CATEGORY_LABEL:
            self.warnings.append(f"作品类别 {self._get('category')} 不在三大类中")

        if self._get("applicant_type") and self._get("applicant_type") not in APPLICANT_LABEL:
            self.warnings.append(f"申报者类型 {self._get('applicant_type')} 不在 individual/collective 中")

        if not self._get("abstract"):
            self.warnings.append("缺少 作品简介（abstract），将使用占位文本")
        if not self._get("background"):
            self.warnings.append("缺少 选题背景（background），将使用占位文本")
        if not self._get("research_method"):
            self.warnings.append("缺少 研究方法（research_method），将使用占位文本")
        if not self._get("results_discussion"):
            self.warnings.append("缺少 研究结果（results_discussion），将使用占位文本")
        if not self._get("conclusion"):
            self.warnings.append("缺少 结论（conclusion），将使用占位文本")
        if not self._get("innovations"):
            self.warnings.append("缺少 创新点（innovations），将使用占位文本")
        if not self._get("references"):
            self.warnings.append("缺少 参考文献（references），将使用占位文本")

        if self._get("applicant_type") == "collective":
            members = self._get_list("team_members")
            if not members:
                self.warnings.append("集体申报缺少团队成员（team_members）")
            elif len(members) > 7:
                self.warnings.append(f"集体申报团队成员 {len(members)} 人超过 7 人上限")

        work_name = self._get("work_full_name")
        if work_name and len(work_name) > 30:
            self.warnings.append(f"作品全称 {len(work_name)} 字超过 30 字上限")

        return self.warnings


# ============================================================
# 默认示例数据
# ============================================================

DEFAULT_DATA = {
    "work_full_name": "基于深度学习的分布式光伏故障智能诊断系统",
    "category": "tech_invention",
    "applicant_type": "collective",
    "leader_name": "张三", "leader_id": "202212345", "leader_gender": "男",
    "leader_age": "21", "leader_major": "电气工程及其自动化",
    "leader_grade": "2022 级 大三", "leader_phone": "138XXXXXXXX",
    "college": "电气工程学院", "school": "XX 大学",
    "apply_date": "2025 年 3 月 15 日",
    "discipline": "0808 电气工程",
    "duration": "2024.09-2025.08（12 个月）",
    "advisor_name": "李教授", "advisor_title": "教授",
    "advisor_research": "新能源发电与并网技术", "advisor_phone": "139XXXXXXXX",
    "team_members": [
        {"name": "李四", "id": "202212346", "major": "电气工程及其自动化", "grade": "2022 级 大三", "role": "数据采集与标注"},
        {"name": "王五", "id": "202212347", "major": "自动化", "grade": "2022 级 大三", "role": "硬件搭建与现场测试"},
        {"name": "赵六", "id": "202212348", "major": "统计学", "grade": "2022 级 大三", "role": "数据分析与论文撰写"},
    ],
    "abstract": "本项目针对分布式光伏故障检测响应慢、误报率高、依赖人工巡检的痛点，开发基于卷积神经网络（CNN）的智能诊断系统。通过与 XX 新能源公司合作获取真实电站 5000+ 故障样本进行训练，覆盖热斑、隐裂、二极管故障等 8 类典型故障，预实验准确率已达 92.3%，单次检测响应时间小于 1 秒。预期产出 1 套原型系统、1 篇中文核心论文、1 项发明专利。项目已完成前期调研与 500 样本预训练。",
    "background": [
        "（一）时代背景：随着我国『双碳』战略推进，分布式光伏装机量 2024 年突破 6 亿千瓦（国家能源局数据），年均增速 35% 以上。但分布式光伏分布广、单点容量小，传统集中式故障检测方案难以适配，亟需智能化轻量化解决方案。",
        "（二）现实痛点：调研 XX 省 3 家光伏运维企业发现，分布式光伏故障检测主要依赖人工巡检，平均响应时间超过 48 小时，故障期间发电损失达 5%~8%。某 50 户村集体光伏项目 2024 年因热斑故障未及时处理，单户年损失超 800 元。若按全国 6 亿千瓦分布式光伏估算，年损失或超 200 亿元。现有商用方案准确率仅 75%~80%，且无法识别隐裂等微弱故障。",
        "（三）国内外研究现状：早期方法（Smith 2020, Wang 2021）主要基于 SVM、决策树等传统机器学习，依赖人工特征提取，准确率约 75%~80%，且对故障类型识别有限。近年来 Zhang (2022)、Li (2023) 尝试引入深度学习，但所用数据集多为实验室仿真，缺乏真实电站数据验证，覆盖故障类别仅 3~5 类。本项目关键差异在于：（1）使用真实电站 5000+ 故障样本；（2）覆盖 8 类典型故障（已有研究覆盖类别的 1.6~2.7 倍）；（3）融合可见光+红外+电气参数三模态数据。",
        "（四）项目意义：理论上探索深度学习在小样本、多类别故障诊断中的适用边界，验证端到端 CNN 的特征学习能力；实践上与 XX 公司合作开发可落地系统，预期将故障响应时间从 48 小时缩短至 1 小时，单户年增收约 600 元；社会上助力乡村振兴与『双碳』目标实现，推动分布式光伏智能化运维升级。",
    ],
    "research_method": [
        "（一）文献调研：系统梳理 2020~2024 年国内外光伏故障检测文献 50 篇（中文 28 篇、英文 22 篇），核心结论为传统 SVM/决策树方法准确率 75%~80%，深度学习方法近年兴起但数据集多为实验室仿真，真实电站验证不足。本项目以此为研究空白，聚焦真实电站数据。",
        "（二）数据采集：与 XX 新能源公司签署数据合作协议，获取 5 个电站 2023~2024 年运行数据，含故障样本 5000+。采用专家交叉标注（3 人独立标注+多数表决）保证标注质量，标注一致性 Kappa 系数 0.87。涉及电站运营数据已签署保密协议，未涉及个人信息，免伦理审查。样本覆盖热斑、隐裂、二极管故障等 8 类典型故障，每类含可见光/红外/电气参数三类数据。",
        "（三）实验设计：在 XX 大学新能源实验室搭建光伏组件故障模拟平台，控制环境温度 25±2℃、辐照度 1000 W/m²，对 8 类故障各采集 500+ 样本。以传统 SVM 为对照组，本组 CNN 模型随机初始化训练 3 次取平均，评估指标为准确率/召回率/F1。模型架构采用 ResNet-50 改进版，引入注意力机制，学习率 0.001，批量大小 32，训练 100 epoch。统计差异用配对 t 检验（p<0.05 视为显著）。",
        "（四）实地测试：将训练好的 CNN 模型部署在 XX 省 3 个分布式电站（合计装机 2.5 MW），自 2025 年 1 月起试运行 3 个月，累计处理实时数据 120 万条。评估指标为准确率/召回率/响应时间/误报率，对比基线为电站原有人工巡检记录。误差分析显示二极管故障因样本量小识别精度偏低，其余 7 类故障准确率均≥90%。",
    ],
    "results_discussion": [
        "（一）主要发现：本方法在测试集上准确率 92.3%，召回率 89.5%，F1 0.908（详见图 3、表 4）。其中热斑故障识别准确率 96.8%，隐裂故障 93.2%，二极管故障因样本仅 200 个准确率 84.2%。单次检测响应时间 0.78 秒，误报率 3.2%。多模态融合比单模态准确率提升 3~5 个百分点。在 3 个电站实地测试中累计识别故障 87 次，其中 81 次为真实故障，准确率 93.1%。",
        "（二）与已有研究对比：与 Smith (2020) 的 SVM 方法准确率 80.1% 相比，本方法准确率提升 12.2 个百分点（p<0.01，配对 t 检验）。与 Zhang (2022) 的 CNN 方法准确率 88.5% 相比，本方法准确率提升 3.8 个百分点，主要差异源于本方法使用真实电站数据（Zhang 用仿真数据）。本方法覆盖 8 类故障，已有研究多覆盖 3~5 类，覆盖度为 1.6~2.7 倍。",
        "（三）结果讨论：本方法验证了端到端 CNN 在小样本、多类别故障诊断中的有效性，证明真实电站数据训练的模型泛化能力显著优于仿真数据训练。多模态融合提升微小故障（如隐裂）识别精度，说明单一电气参数不足以刻画光伏故障全貌。本方法可推广至其他分布式能源（风电、储能）故障诊断领域，推广边界为单电站装机 ≤10 MW、故障类别 ≤15 类。",
        "（四）局限性分析：本研究存在 3 点局限：（1）样本仅来自华东地区电站，未涵盖西北高辐照、高粉尘环境，模型在西北地区电站泛化能力待验证；（2）二极管故障样本仅 200 个，小样本类别精度仍有提升空间，未来计划采集 1000+ 二极管故障样本；（3）模型在边缘设备部署推理延迟待优化，当前 RTX 4090 推理 0.78 秒，移植到 Jetson Nano 后延迟升至 2.1 秒，需进行模型轻量化。",
    ],
    "conclusion": "本研究开发了基于端到端 CNN 的分布式光伏故障智能诊断系统，在真实电站 5000+ 样本上准确率 92.3%，单次响应 0.78 秒，覆盖 8 类典型故障。学术贡献在于验证了端到端深度学习在小样本、多类别故障诊断中的适用边界，并提供真实电站数据基准。下一步计划扩展样本至 20000+，覆盖西北高辐照环境，并优化边缘设备部署推理延迟至 1 秒以内。",
    "innovations": [
        {"type": "方法创新", "description": "端到端 CNN 自动学习故障特征，引入注意力机制提升微小故障识别能力。", "advantage": "vs 传统 SVM 人工特征，准确率 80%→92.3%，提升 12.3 个百分点"},
        {"type": "数据创新", "description": "与 XX 公司合作获取真实电站 5000+ 故障样本，覆盖 8 类典型故障。", "advantage": "vs 实验室仿真数据集 3~5 类，覆盖度提升 1.6~2.7 倍"},
        {"type": "视角创新", "description": "融合可见光+红外+电气参数三模态数据，特征级与决策级双融合。", "advantage": "vs 单模态准确率提升 3~5 个百分点，微小故障识别显著增强"},
        {"type": "应用创新", "description": "原型系统部署 3 个电站试运行 3 个月，累计处理 120 万条实时数据。", "advantage": "vs 实验室原型无实地验证，故障响应时间 48 小时→0.78 秒"},
    ],
    "references": [
        {"authors": "张三, 李四", "title": "基于深度学习的光伏故障检测研究", "type": "J", "source": "电力系统自动化", "year": "2024", "volume": "48", "issue": "3", "pages": "45-52"},
        {"authors": "Smith J, Brown K", "title": "Deep learning for PV fault diagnosis", "type": "J", "source": "IEEE Transactions on Sustainable Energy", "year": "2023", "volume": "14", "issue": "2", "pages": "1234-1245"},
        {"authors": "Wang Y, Liu Z", "title": "A survey of PV fault detection methods", "type": "J", "source": "Renewable and Sustainable Energy Reviews", "year": "2021", "volume": "150", "issue": "", "pages": "111-125"},
        {"authors": "Zhang Q", "title": "CNN-based PV fault classification", "type": "C", "source": "Proc. of ICRA", "year": "2022", "volume": "", "issue": "", "pages": "234-240"},
        {"authors": "李华", "title": "分布式光伏运维技术", "type": "M", "source": "", "year": "2023", "publisher": "中国电力出版社", "city": "北京", "pages": "88-105"},
        {"authors": "赵六", "title": "风电场功率预测研究", "type": "D", "source": "清华大学", "year": "2022", "city": "北京", "publisher": ""},
        {"authors": "全国太阳能标准化技术委员会", "title": "分布式光伏发电系统技术规范", "type": "S", "source": "GB/T 19964-2024", "year": "2024", "city": "北京", "publisher": "中国标准出版社"},
        {"authors": "Li X, Wang M", "title": "Multi-modal fusion for fault diagnosis", "type": "J", "source": "Applied Energy", "year": "2023", "volume": "345", "issue": "", "pages": "121-135"},
    ],
    "appendix": [
        {"title": "8 类光伏故障样本示例图", "content": "图 1~图 8 展示热斑、隐裂、二极管故障等 8 类典型故障的可见光与红外双视图样本。每类含 5 个代表性样本，标注故障位置与特征。"},
        {"title": "原始数据字段说明", "content": "原始数据 5000+ 样本，每样本含 128 维特征（时域 64 维+频域 64 维），经 PCA 降维至 32 维。字段含 sample_id/fault_type/timestamp/irradiance/temperature/voltage/current 等。"},
        {"title": "CNN 模型核心代码", "content": "Python 3.10+PyTorch 2.1 实现，含数据加载（DataLoader）+模型定义（ResNet-50 改进版，引入 CBAM 注意力）+训练循环（100 epoch，学习率 0.001，批量 32）+评估脚本。"},
        {"title": "实物照片", "content": "原型系统实物照片含 5 视图（正/侧/背/内部/部署现场），尺寸 200mm×150mm×80mm，含 OLED 显示屏+按钮+以太网口+电源接口。"},
    ],
}


# ============================================================
# CLI 入口
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="挑战杯课外学术科技作品竞赛申报书 docx 生成器",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python build.py --data data.json --out output.docx\n"
            "  python build.py --demo --out demo.docx\n"
            "\n"
            "JSON 字段定义详见 SKILL.md 第十章。"
        ),
    )
    parser.add_argument("--data", type=str, default=None,
                        help="数据 JSON 文件路径（与 --demo 二选一）")
    parser.add_argument("--out", type=str, required=True,
                        help="输出 docx 文件路径")
    parser.add_argument("--demo", action="store_true",
                        help="使用内置示例数据生成演示文档")
    args = parser.parse_args()

    if args.demo:
        data = DEFAULT_DATA
        print("ℹ️ 使用内置示例数据生成演示文档")
    elif args.data:
        if not os.path.exists(args.data):
            sys.stderr.write(f"❌ 数据文件不存在：{args.data}\n")
            sys.exit(1)
        try:
            with open(args.data, "r", encoding="utf-8") as f:
                data = json.load(f)
        except json.JSONDecodeError as e:
            sys.stderr.write(f"❌ JSON 解析失败：{e}\n")
            sys.exit(1)
    else:
        sys.stderr.write("❌ 必须提供 --data 或 --demo 参数\n")
        parser.print_help()
        sys.exit(1)

    builder = ApplicationDocBuilder()
    try:
        builder.build(data, args.out)
    except Exception as e:
        sys.stderr.write(f"❌ 生成失败：{e}\n")
        sys.exit(1)


# 入口：python build.py --demo --out demo.docx
# 或：  python build.py --data data.json --out output.docx
if __name__ == "__main__":
    main()
