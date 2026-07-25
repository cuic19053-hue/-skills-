#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
大学生创新创业训练计划-创新训练项目申报书 docx 生成器

按统一格式标准生成 Word 文档：
- A4 纸张，页边距上下 2.54cm 左右 2.5cm
- 正文：宋体小四，1.5 倍行距，首行缩进 2 字符
- 一级标题：黑体三号，居中
- 二级标题：黑体小三，左对齐
- 三级标题：宋体四号加粗
- 表格：宋体五号，居中

使用方式：
    python build.py --data data.json --out output.docx

JSON 字段详见 SKILL.md 第十一章。
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


def add_table_from_data(doc, headers: List[str], rows: List[List[str]],
                        col_widths: Optional[List[float]] = None):
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
            set_cell_text(cells[j], val, font_name=FONT_SONG,
                          font_size=SIZE_WU, bold=False)
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


def merge_vertical_cells(table, col_idx: int, start_row: int, end_row: int):
    """纵向合并单元格（用于签字栏预留空白）"""
    cells = [table.rows[r].cells[col_idx] for r in range(start_row, end_row + 1)]
    merged = cells[0]
    for c in cells[1:]:
        merged = merged.merge(c)


# ============================================================
# ApplicationDocBuilder 主类
# ============================================================

class ApplicationDocBuilder:
    """大创-创新训练项目申报书 docx 构建器"""

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

    def _get(self, *keys, default=""):
        """安全取嵌套字段，缺字段返回默认值"""
        cur = self.data
        for k in keys:
            if not isinstance(cur, dict):
                return default
            cur = cur.get(k, default)
        return cur if cur is not None else default

    def add_h1(self, text):
        return add_heading_level1(self.doc, text)

    def add_h2(self, text):
        return add_heading_level2(self.doc, text)

    def add_h3(self, text):
        return add_heading_level3(self.doc, text)

    def add_para(self, text, indent=True):
        return add_body_paragraph(self.doc, text, indent=indent)

    def add_table(self, headers, rows, col_widths=None):
        return add_table_from_data(self.doc, headers, rows, col_widths)

    def add_page_break(self):
        add_page_break(self.doc)

    # --------------------------------------------------------
    # 封面
    # --------------------------------------------------------

    def _add_cover(self):
        """封面：黑体二号标题 + 副标题 + 5 行下划线信息"""
        for _ in range(2):
            self.doc.add_paragraph()

        title = "国家级大学生创新创业训练计划项目申报书"
        add_paragraph_with_format(
            self.doc, title,
            font_name=FONT_HEI, font_size=SIZE_ER, bold=True,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            first_line_indent=False,
            space_before=12, space_after=12,
        )

        subtitle = f"（{self._get('project_type', default='创新训练项目')}）"
        add_paragraph_with_format(
            self.doc, subtitle,
            font_name=FONT_HEI, font_size=SIZE_SAN, bold=False,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            first_line_indent=False,
            space_after=24,
        )

        for _ in range(2):
            self.doc.add_paragraph()

        info_items = [
            ("项目名称", self._get("project_name")),
            ("项目负责人", self._get("leader_name")),
            ("指导教师", self._get("advisor_name")),
            ("所在学院", self._get("college")),
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
    # 基本信息
    # --------------------------------------------------------

    def _add_basic_info_table(self):
        """一、基本信息表（9 行 2 列）"""
        self.add_h1("一、基本信息")
        basic = self._get("basic_info", default={})
        if not isinstance(basic, dict):
            basic = {}
        budget_str = str(basic.get("budget", self._get("budget_total", "")))
        if budget_str and not budget_str.endswith("元"):
            budget_str = f"{budget_str} 元"

        rows = [
            ["项目名称", basic.get("project_name", self._get("project_name"))],
            ["项目类型", basic.get("project_type",
                                    self._get("project_type", "创新训练项目"))],
            ["项目来源", basic.get("project_source", "A 学生自主选题")],
            ["所属学科", basic.get("discipline", "")],
            ["起止时间", basic.get("duration", "")],
            ["申请经费", budget_str],
            ["负责人", basic.get("leader_info",
                                  f"{self._get('leader_name')} / "
                                  f"{self._get('leader_id')} / "
                                  f"{self._get('leader_major')} / "
                                  f"{self._get('leader_grade')} / "
                                  f"{self._get('leader_phone')}")],
            ["团队成员", basic.get("team_members", "")],
            ["指导教师", basic.get("advisor_info",
                                    f"{self._get('advisor_name')} / "
                                    f"{self._get('advisor_title')} / "
                                    f"{self._get('advisor_research')}")],
        ]
        self.add_table(["项目", "内容"], rows, col_widths=[4.5, 11.5])

    # --------------------------------------------------------
    # 项目简介
    # --------------------------------------------------------

    def _add_abstract(self):
        """二、项目简介（300~500 字）"""
        self.add_h1("二、项目简介")
        abstract = self._get("abstract", default="")
        if abstract:
            self.add_para(abstract)
        else:
            self.add_para("（请填写项目简介，300~500 字，按 4 句结构撰写："
                          "痛点+做什么 / 怎么做+量化目标 / 产出什么 / 现状。）")

    # --------------------------------------------------------
    # 立项背景与意义
    # --------------------------------------------------------

    def _add_background(self):
        """三、立项背景与意义（800~1200 字，4 段）"""
        self.add_h1("三、立项背景与意义")
        background = self._get("background", default=[])
        if isinstance(background, str):
            background = [background]
        if not background:
            self.add_h2("（一）时代背景")
            self.add_para("（请填写时代背景，150~200 字，3 句话讲政策/行业/学术趋势，"
                          "必须有权威数据来源。）")
            self.add_h2("（二）现实痛点")
            self.add_para("（请填写现实痛点，300~400 字，2~3 个真实场景，必须可量化。）")
            self.add_h2("（三）国内外研究现状")
            self.add_para("（请填写研究现状，200~300 字，评述已有方案不足，引出本项目差异。）")
            self.add_h2("（四）项目意义")
            self.add_para("（请填写项目意义，150~300 字，理论/实践/社会三角度，至少两个。）")
        else:
            for para in background:
                self.add_para(para)

    # --------------------------------------------------------
    # 研究内容与目标
    # --------------------------------------------------------

    def _add_research_content(self):
        """四、项目研究内容与目标（800~1200 字，3 子节）"""
        self.add_h1("四、项目研究内容与目标")

        self.add_h2("（一）研究内容")
        contents = self._get("research_content", default=[])
        if isinstance(contents, str):
            contents = [contents]
        if contents:
            for i, c in enumerate(contents, 1):
                self.add_para(f"{i}. {c}")
        else:
            self.add_para("（请填写研究内容，3~5 个子任务，每个 100~200 字，"
                          "结构：任务名+做什么+方法+产出。）")

        self.add_h2("（二）研究目标")
        goal = self._get("research_goal", default="")
        if goal:
            self.add_para(goal)
        else:
            self.add_para("（请填写研究目标，1 个总目标 + 3~4 个阶段目标，全部可量化。）")

        self.add_h2("（三）拟解决的关键问题")
        problems = self._get("key_problems", default=[])
        if isinstance(problems, str):
            problems = [problems]
        if problems:
            for i, q in enumerate(problems, 1):
                self.add_para(f"{i}. {q}")
        else:
            self.add_para("（请填写关键问题，2~3 个，每个一句话讲清技术难点。）")

    # --------------------------------------------------------
    # 创新点
    # --------------------------------------------------------

    def _add_innovation(self):
        """五、项目创新点（400~600 字，至少 2 个）"""
        self.add_h1("五、项目创新点")
        innovations = self._get("innovations", default=[])
        if isinstance(innovations, str):
            innovations = [innovations]
        if innovations:
            for i, inv in enumerate(innovations, 1):
                self.add_para(f"创新点 {i}：{inv}")
        else:
            self.add_para("（请填写创新点，至少 2 个，每个 150~200 字。"
                          "结构：[类型]。传统方法[描述]，本项目[方法]，[量化优势]。"
                          "禁止使用『首次』『先进』『实现』等无支撑词。）")

    # --------------------------------------------------------
    # 技术路线与研究方法
    # --------------------------------------------------------

    def _add_technical_route(self):
        """六、技术路线与研究方法（500~800 字 + 流程图）"""
        self.add_h1("六、技术路线与研究方法")

        self.add_h2("（一）总体技术路线")
        route = self._get("tech_route", default="")
        if route:
            self.add_para(route)
        else:
            self.add_para("（请填写总体技术路线，1 段文字 + 1 张流程图。"
                          "横向 5 阶段，每阶段标注交付物。图下方加图注『图 1 项目技术路线图』。）")
        flowchart = self._get("tech_flowchart_image", default="")
        if flowchart and os.path.exists(flowchart):
            try:
                p = self.doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run()
                run.add_picture(flowchart, width=Cm(15))
                add_paragraph_with_format(
                    self.doc, "图 1 项目技术路线图",
                    font_name=FONT_HEI, font_size=SIZE_WU, bold=False,
                    alignment=WD_ALIGN_PARAGRAPH.CENTER,
                    first_line_indent=False,
                )
            except Exception as e:
                self.add_para(f"（流程图插入失败：{e}）")

        self.add_h2("（二）研究方法")
        methods = self._get("methods", default=[])
        if isinstance(methods, str):
            methods = [methods]
        if methods:
            for i, m in enumerate(methods, 1):
                self.add_para(f"{i}. {m}")
        else:
            self.add_para("（请填写研究方法，3~5 个，每个 50~100 字说明用途。）")

        self.add_h2("（三）数据来源与实验条件")
        data_src = self._get("data_source", default="")
        if data_src:
            self.add_para(data_src)
        else:
            self.add_para("（请填写数据来源、实验设备型号、软件工具及版本。）")

    # --------------------------------------------------------
    # 实施方案与进度
    # --------------------------------------------------------

    def _add_implementation_plan(self):
        """七、项目实施方案与进度安排（4 列表格）"""
        self.add_h1("七、项目实施方案与进度安排")
        schedule = self._get("schedule", default=[])
        if schedule:
            rows = []
            for s in schedule:
                if not isinstance(s, dict):
                    continue
                rows.append([
                    s.get("phase", ""),
                    s.get("time", ""),
                    s.get("work", ""),
                    s.get("output", ""),
                ])
            self.add_table(
                ["阶段", "时间", "主要工作", "阶段成果"],
                rows,
                col_widths=[2.5, 3.0, 6.5, 4.0],
            )
        else:
            self.add_para("（请填写进度安排，按月划分，每阶段标注交付物。"
                          "建议 5~6 行，留 1~2 月弹性时间。）")

    # --------------------------------------------------------
    # 预期成果
    # --------------------------------------------------------

    def _add_expected_results(self):
        """八、预期成果（必须可量化）"""
        self.add_h1("八、预期成果")
        outcomes = self._get("expected_outcomes", default=[])
        if isinstance(outcomes, str):
            outcomes = [outcomes]
        if outcomes:
            for o in outcomes:
                self.add_para(f"• {o}", indent=False)
        else:
            self.add_para("（请填写预期成果，每项含数量+级别+平台。"
                          "如：中文核心论文 1 篇（拟投《XX》）、发明专利 1 项、原型系统 1 套。）",
                          indent=False)

    # --------------------------------------------------------
    # 经费预算
    # --------------------------------------------------------

    def _add_budget(self):
        """九、经费预算（3 列表格：科目/金额/计算依据）"""
        self.add_h1("九、经费预算")
        items = self._get("budget_items", default=[])
        if items:
            rows = []
            total = 0
            for b in items:
                if not isinstance(b, dict):
                    continue
                amount_str = str(b.get("amount", "0"))
                try:
                    amount_num = int(amount_str)
                except ValueError:
                    amount_num = 0
                total += amount_num
                rows.append([
                    b.get("item", ""),
                    f"{amount_num} 元",
                    b.get("basis", ""),
                ])
            rows.append(["合计", f"{total} 元", ""])
            self.add_table(
                ["预算科目", "金额", "计算依据"],
                rows,
                col_widths=[3.5, 3.0, 9.5],
            )
        else:
            self.add_para("（请填写经费预算，5 类标准科目：资料费/调研差旅费/"
                          "实验材料费/会议费/印刷复印。每项金额非整数，附计算依据。）")

    # --------------------------------------------------------
    # 前期工作基础
    # --------------------------------------------------------

    def _add_preliminary_work(self):
        """十、前期工作基础（300~500 字，3 子节）"""
        self.add_h1("十、前期工作基础")

        self.add_h2("（一）团队基础")
        team = self._get("team_foundation", default="")
        self.add_para(team if team else
                      "（请填写团队基础：成员相关课程、项目经验、技能匹配度。）")

        self.add_h2("（二）指导教师基础")
        advisor = self._get("advisor_foundation", default="")
        self.add_para(advisor if advisor else
                      "（请填写指导教师基础：主持项目、发表论文、研究方向匹配度。）")

        self.add_h2("（三）实验条件")
        lab = self._get("lab_condition", default="")
        self.add_para(lab if lab else
                      "（请填写实验条件：实验室设备、软件平台、合作单位支持。）")

    # --------------------------------------------------------
    # 签字栏
    # --------------------------------------------------------

    def _add_signature_section(self):
        """十一/十二、指导教师意见、学院评审意见"""
        self.add_h1("十一、指导教师意见")
        for _ in range(6):
            self.doc.add_paragraph()
        self.add_para(
            "指导教师签字：____________________    "
            "日期：______年____月____日",
            indent=False,
        )

        self.add_h1("十二、学院评审意见")
        for _ in range(6):
            self.doc.add_paragraph()
        self.add_para(
            "学院盖章：____________________    "
            "日期：______年____月____日",
            indent=False,
        )

        if self._get("include_school_approval", default=False):
            self.add_h1("十三、学校审批意见")
            for _ in range(6):
                self.doc.add_paragraph()
            self.add_para(
                "学校盖章：____________________    "
                "日期：______年____月____日",
                indent=False,
            )

    # --------------------------------------------------------
    # 主构建方法
    # --------------------------------------------------------

    def build(self, data: Dict[str, Any], output_path: str) -> str:
        """主构建方法：编排 12 栏目，生成 docx

        Args:
            data: 申报书字段字典
            output_path: 输出 docx 路径

        Returns:
            实际保存路径
        """
        try:
            self.data = data if isinstance(data, dict) else {}
            self._validate_data()

            self._add_cover()
            self._add_basic_info_table()
            self._add_abstract()
            self._add_background()
            self._add_research_content()
            self._add_innovation()
            self._add_technical_route()
            self._add_implementation_plan()
            self._add_expected_results()
            self._add_budget()
            self._add_preliminary_work()
            self._add_signature_section()

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
        warnings = []
        p0_fields = [
            ("project_name", "项目名称"),
            ("leader_name", "负责人姓名"),
            ("advisor_name", "指导教师姓名"),
            ("college", "所在学院"),
        ]
        for key, name in p0_fields:
            if not self._get(key):
                warnings.append(f"缺少 {name}（{key}）")

        basic = self._get("basic_info", default={})
        if not isinstance(basic, dict):
            basic = {}
        if not basic.get("discipline") and not self._get("discipline"):
            warnings.append("缺少 所属学科（discipline）")
        if not basic.get("duration") and not self._get("duration"):
            warnings.append("缺少 起止时间（duration）")

        if not self._get("abstract"):
            warnings.append("缺少 项目简介（abstract），将使用占位文本")
        if not self._get("background"):
            warnings.append("缺少 立项背景（background），将使用占位文本")
        if not self._get("innovations"):
            warnings.append("缺少 创新点（innovations），将使用占位文本")

        items = self._get("budget_items", default=[])
        if items:
            total = 0
            for b in items:
                if isinstance(b, dict):
                    try:
                        total += int(b.get("amount", 0))
                    except (ValueError, TypeError):
                        pass
            budget_total_str = str(self._get("budget_total", default="")).strip()
            try:
                budget_total_num = int(budget_total_str)
            except ValueError:
                budget_total_num = -1
            if budget_total_num >= 0 and total != budget_total_num:
                warnings.append(
                    f"预算合计 {total} 元 与申请经费 {budget_total_num} 元不一致"
                )

        if warnings:
            print("⚠️ 数据校验警告：", file=sys.stderr)
            for w in warnings:
                print(f"  - {w}", file=sys.stderr)
        return warnings


# ============================================================
# 默认示例数据
# ============================================================

DEFAULT_DATA = {
    "project_name": "分布式光伏故障智能诊断系统",
    "project_level": "国家级",
    "project_type": "创新训练项目",
    "leader_name": "张三",
    "leader_id": "202212345",
    "leader_major": "电气工程及其自动化",
    "leader_grade": "2022 级 大三",
    "leader_phone": "138XXXXXXXX",
    "advisor_name": "李教授",
    "advisor_title": "教授",
    "advisor_research": "新能源发电与并网技术",
    "college": "电气工程学院",
    "apply_date": "2025 年 3 月 15 日",
    "basic_info": {
        "project_name": "分布式光伏故障智能诊断系统",
        "project_type": "创新训练项目",
        "project_source": "A 学生自主选题",
        "discipline": "0808 电气工程",
        "duration": "2025.04-2026.03（12 个月）",
        "budget": "9500",
        "leader_info": "张三 / 202212345 / 电气工程及其自动化 / 大三 / 138XXXXXXXX",
        "team_members": "李四（数据采集）、王五（硬件搭建）、赵六（数据分析）",
        "advisor_info": "李教授 / 教授 / 新能源发电与并网技术 / 139XXXXXXXX",
    },
    "abstract": "本项目针对分布式光伏故障检测响应慢、误报率高、依赖人工巡检的痛点，"
                "开发基于卷积神经网络的智能诊断系统。通过与 XX 新能源公司合作获取真实电站 "
                "5000+ 故障样本进行训练，覆盖热斑、隐裂、二极管故障等 8 类典型故障，"
                "预实验准确率已达 92%，单次检测响应时间小于 1 秒。预期产出：1 套原型系统、"
                "1 篇中文核心论文、1 项发明专利。项目已完成前期调研与 500 样本预训练。",
    "background": [
        "时代背景：随着我国『双碳』战略推进，分布式光伏装机量 2024 年突破 6 亿千瓦"
        "（国家能源局数据），年均增速 35% 以上。但分布式光伏分布广、单点容量小，"
        "传统集中式故障检测方案难以适配。",
        "现实痛点：调研 XX 省 3 家光伏运维企业发现，分布式光伏故障检测主要依赖人工巡检，"
        "平均响应时间超过 48 小时，故障期间发电损失达 5%~8%。某 50 户村集体光伏项目 "
        "2024 年因热斑故障未及时处理，单户年损失超 800 元。",
        "国内外研究现状：早期方法（Smith 2020, Wang 2021）主要基于 SVM、决策树等传统"
        "机器学习，依赖人工特征提取，准确率约 75%~80%。Zhang (2022)、Li (2023) 引入"
        "深度学习，但数据集多为实验室仿真。本项目关键差异：使用真实电站 5000+ 样本，"
        "覆盖 8 类故障。",
        "项目意义：理论上探索深度学习在小样本、多类别故障诊断中的适用边界；实践上与 XX "
        "公司合作开发可落地系统，预期将故障响应时间从 48 小时缩短至 1 小时，单户年增收 "
        "约 600 元；社会上助力乡村振兴与『双碳』目标实现。",
    ],
    "research_content": [
        "故障样本数据集构建：与 XX 新能源公司合作采集真实电站故障样本 5000+，"
        "覆盖热斑、隐裂、二极管故障等 8 类典型故障，每类含可见光/红外/电气参数三类数据。",
        "基于 CNN 的故障识别模型设计：设计适用于光伏故障识别的卷积神经网络架构，"
        "针对故障特征优化卷积核设计，引入注意力机制提升对微小故障的识别能力。",
        "多模态数据融合方法：研究可见光、红外、电气参数三类数据的融合方法，"
        "设计特征级/决策级融合策略，提升复杂故障的诊断准确率。",
        "原型系统开发与实地部署：开发包含数据采集、模型推理、告警推送、运维建议的"
        "完整系统，部署在合作电站试运行 3 个月。",
    ],
    "research_goal": "总目标：开发准确率 ≥95%、响应时间 <1 秒的分布式光伏故障智能诊断系统。"
                     "阶段目标 1：完成 5000+ 样本数据集构建（2025.06 前）；"
                     "阶段目标 2：完成 CNN 模型训练，准确率 ≥92%（2025.09 前）；"
                     "阶段目标 3：完成多模态融合，准确率 ≥95%（2025.11 前）；"
                     "阶段目标 4：原型系统部署试运行（2026.02 前）。",
    "key_problems": [
        "小样本故障类别（如二极管故障样本仅 200 个）的识别精度",
        "多模态数据的特征对齐与融合策略",
        "模型在实际电站部署后的泛化能力",
    ],
    "innovations": [
        "方法创新。传统光伏故障检测依赖人工特征提取（准确率约 80%），"
        "本项目采用端到端 CNN 自动学习故障特征，预实验准确率已达 92%。",
        "数据创新。公开数据集多为实验室仿真，本项目与 XX 公司合作获取真实电站 "
        "5000+ 故障样本，覆盖 8 类典型故障，是已有研究覆盖类别的 1.6~2.7 倍。",
        "视角创新。已有故障检测多从单一电气参数切入，本项目融合可见光、红外、"
        "电气参数三模态数据，预实验显示多模态融合比单模态准确率提升 3~5 个百分点。",
    ],
    "tech_route": "总体技术路线分 5 阶段（见图 1）：① 调研阶段（2025.03-04），"
                  "完成 50 篇文献调研与方案设计，产出调研报告；② 数据采集（2025.05-06），"
                  "获取真实电站故障样本 5000+，产出标注数据集；③ 预处理（2025.07），"
                  "完成数据清洗与增强，产出预处理脚本；④ 模型训练（2025.08-09），"
                  "完成 CNN 模型设计与调优，产出模型权重；⑤ 评估部署（2025.10-12），"
                  "实地测试与对比实验，产出评估报告与原型系统。",
    "methods": [
        "文献分析法：系统梳理国内外光伏故障检测研究，建立技术对比表，明确研究空白。",
        "数据采集与标注法：与合作企业共建故障样本库，采用专家交叉标注（3 人独立标注，"
        "多数表决）保证标注质量。",
        "实验法：设计 CNN 模型架构，在数据集上进行训练/验证/测试，与 SVM、决策树等"
        "基线模型对比。",
        "实地测试法：将训练好的模型部署在合作电站，收集 3 个月实际运行数据评估泛化能力。",
    ],
    "data_source": "数据来源：与 XX 新能源公司签署数据合作协议，获取 5 个电站 2023-2024 年"
                   "运行数据，含故障样本 5000+。实验设备：红外热成像仪 FLIR T630"
                   "（电气工程学院实验室）、光伏组件测试平台。软件工具：Python 3.10、"
                   "PyTorch 2.1、OpenCV 4.8。",
    "schedule": [
        {"phase": "准备", "time": "2025.03-04",
         "work": "文献调研 50 篇、方案设计", "output": "调研报告 1 份"},
        {"phase": "数据", "time": "2025.05-06",
         "work": "采集 5000+ 样本、标注", "output": "数据集 1 套"},
        {"phase": "模型", "time": "2025.07-09",
         "work": "CNN 模型训练、调优", "output": "模型代码 + 权重"},
        {"phase": "评估", "time": "2025.10-11",
         "work": "实地测试、对比实验", "output": "评估报告 1 份"},
        {"phase": "总结", "time": "2025.12",
         "work": "论文撰写、专利申请", "output": "论文 1 篇 + 专利 1 项"},
    ],
    "expected_outcomes": [
        "中文核心期刊论文 1 篇（拟投《电力系统自动化》）",
        "发明专利申请 1 项（与合作企业共同申请）",
        "原型系统 1 套（部署在合作电站试运行 3 个月）",
        "调研报告 1 份（约 2 万字）",
    ],
    "budget_items": [
        {"item": "资料费", "amount": "850",
         "basis": "图书 15 本 × 50 元 + 数据库订阅 100 元"},
        {"item": "调研差旅费", "amount": "2600",
         "basis": "2 次实地调研 × 1300 元（含交通、住宿）"},
        {"item": "实验材料费", "amount": "4320",
         "basis": "设备租赁 800 元/月 × 5 月 + 标注 320 元"},
        {"item": "会议费", "amount": "1200",
         "basis": "参加全国新能源会议 1 次"},
        {"item": "印刷复印", "amount": "530",
         "basis": "论文版面费 500 + 报告印刷 30"},
    ],
    "team_foundation": "团队 4 名成员已修读《机器学习》《数字图像处理》《自动控制原理》"
                       "等核心课程，3 人有校级大创参与经验。负责人张三掌握 Python/PyTorch，"
                       "曾参与校级光伏监测项目。",
    "advisor_foundation": "指导教师李教授主持国家自然科学基金 1 项（52377214），"
                          "近 3 年发表 SCI 论文 5 篇，研究方向为新能源发电与并网技术，"
                          "与本项目高度契合。",
    "lab_condition": "电气工程学院新能源实验室配备红外热成像仪 FLIR T630、"
                     "GPU 服务器（RTX 4090 ×4）、光伏组件测试平台，可满足本项目需求。",
    "budget_total": "9500",
    "include_school_approval": False,
}


# ============================================================
# CLI 入口
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="大创-创新训练项目申报书 docx 生成器",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python build.py --data data.json --out output.docx\n"
            "  python build.py --demo --out demo.docx\n"
            "\n"
            "JSON 字段定义详见 SKILL.md 第十一章。"
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

if __name__ == "__main__":
    main()
