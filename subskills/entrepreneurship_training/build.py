#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
大学生创新创业训练计划-创业训练项目申报书 docx 生成器

按统一格式标准生成 Word 文档：
- A4 纸张，页边距上下 2.54cm 左右 2.5cm
- 正文：宋体小四，1.5 倍行距，首行缩进 2 字符
- 一级标题：黑体三号，居中
- 二级标题：黑体小三，左对齐
- 三级标题：宋体四号加粗
- 表格：宋体五号，居中

13 栏目结构（创业训练专属）：
封面 / 基本信息 / 项目简介 / 行业背景与市场分析 / 产品或服务介绍 /
商业模式 / 运营方案 / 财务预测 / 团队介绍 / 风险分析 /
预期成果 / 经费预算 / 签字栏

使用方式：
    python build.py --data data.json --out output.docx
    python build.py --demo --out demo.docx

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
# 工具函数（中英文同步设置 eastAsia/ascii/hAnsi）
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
            align = WD_ALIGN_PARAGRAPH.LEFT if j == len(row) - 1 else WD_ALIGN_PARAGRAPH.CENTER
            set_cell_text(cells[j], val, font_name=FONT_SONG,
                          font_size=SIZE_WU, bold=False, alignment=align)
            cells[j].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(w)

    doc.add_paragraph()
    return table


def add_canvas_9grid(doc, canvas_items: List[Dict[str, str]]):
    """商业模式画布 9 宫格表格（3 行 × 3 列，每格含要素名+内容）"""
    table = doc.add_table(rows=3, cols=3)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for idx, item in enumerate(canvas_items[:9]):
        r, c = divmod(idx, 3)
        cell = table.rows[r].cells[c]
        cell.text = ""
        p1 = cell.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run1 = p1.add_run(item.get("element", ""))
        set_run_font(run1, font_name=FONT_HEI, font_size=SIZE_XIAO_WU, bold=True)
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p2.paragraph_format.line_spacing = 1.2
        run2 = p2.add_run(item.get("content", ""))
        set_run_font(run2, font_name=FONT_SONG, font_size=SIZE_XIAO_WU)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
    for row in table.rows:
        for cell in row.cells:
            cell.width = Cm(5.3)
    doc.add_paragraph()


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


# ============================================================
# ApplicationDocBuilder 主类
# ============================================================

class ApplicationDocBuilder:
    """大创-创业训练项目申报书 docx 构建器"""

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

        subtitle = f"（{self._get('project_type', default='创业训练项目')}）"
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
                                    self._get("project_type", "创业训练项目"))],
            ["项目来源", basic.get("project_source", "A 学生自主选题")],
            ["所属行业", basic.get("industry", self._get("industry", ""))],
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
                          "市场机会+做什么 / 怎么做+量化目标 / 商业模式+产出 / 现状。）")

    # --------------------------------------------------------
    # 行业背景与市场分析
    # --------------------------------------------------------

    def _add_market_analysis(self):
        """三、行业背景与市场分析（800~1200 字，4 段 + 2 表格）"""
        self.add_h1("三、行业背景与市场分析")
        market = self._get("market_analysis", default={})
        if not isinstance(market, dict):
            market = {}

        self.add_h2("（一）行业背景")
        bg = market.get("industry_background", "")
        self.add_para(bg if bg else "（请填写行业背景，150~200 字，3 句话讲行业规模/增速/政策，必须有权威数据来源：艾瑞/IDC/国家统计局/行业白皮书。）")

        self.add_h2("（二）目标市场（TAM/SAM/SOM）")
        tss = market.get("tam_sam_som", [])
        if tss:
            rows = [[item.get("level", ""), item.get("definition", ""), item.get("scale", "")] for item in tss if isinstance(item, dict)]
            self.add_table(["层级", "定义", "规模"], rows, col_widths=[2.5, 8.5, 5.0])
        else:
            self.add_para("（请填写 TAM/SAM/SOM 三级市场测算，必须含数据来源。TAM 总市场/SAM 可服务市场/SOM 可获取市场。）")

        self.add_h2("（三）竞品分析")
        comps = market.get("competitors", [])
        if comps:
            rows = [[c.get("name", ""), c.get("positioning", ""), c.get("users", ""), c.get("advantage", ""), c.get("disadvantage", "")] for c in comps if isinstance(c, dict)]
            self.add_table(["竞品", "定位", "用户规模", "优势", "劣势"], rows, col_widths=[2.2, 2.8, 2.5, 4.0, 4.5])
        else:
            self.add_para("（请填写竞品分析，3~5 个直接竞品 + 2~3 个间接竞品，5 列对比表：名称/定位/用户规模/优势/劣势。）")

        self.add_h2("（四）用户画像与需求验证")
        persona = market.get("user_persona", "")
        self.add_para(persona if persona else "（请填写用户画像，6 要素齐全：基本信息+行为+痛点+需求+付费意愿+决策因素。附调研样本量 N≥50。）")

    # --------------------------------------------------------
    # 产品或服务介绍
    # --------------------------------------------------------

    def _add_product_service(self):
        """四、产品或服务介绍（500~800 字，3 子节）"""
        self.add_h1("四、产品或服务介绍")
        ps = self._get("product_service", default={})
        if not isinstance(ps, dict):
            ps = {}

        self.add_h2("（一）产品形态")
        features = ps.get("features", [])
        if isinstance(features, str):
            features = [features]
        if features:
            for i, f in enumerate(features, 1):
                self.add_para(f"{i}. {f}")
        else:
            self.add_para("（请填写 MVP 功能清单，5~8 个核心功能，每个含名称+解决什么问题+关键指标。）")

        self.add_h2("（二）核心功能演示")
        demo = ps.get("demo", "")
        self.add_para(demo if demo else "（请填写 1~2 个核心功能详细描述，含用户操作流程。）")

        self.add_h2("（三）技术实现")
        tech = ps.get("tech_impl", "")
        self.add_para(tech if tech else "（请填写技术架构（前端/后端/数据库）、关键第三方服务、研发投入估算。）")

    # --------------------------------------------------------
    # 商业模式
    # --------------------------------------------------------

    def _add_business_model(self):
        """五、商业模式（500~800 字 + 9 宫格画布）"""
        self.add_h1("五、商业模式")
        bm = self._get("business_model", default={})
        if not isinstance(bm, dict):
            bm = {}

        self.add_h2("（一）商业模式画布（9 要素）")
        canvas = bm.get("canvas", [])
        if canvas and isinstance(canvas, list):
            add_canvas_9grid(self.doc, canvas)
        else:
            self.add_para("（请填写商业模式画布 9 要素：客户细分/价值主张/渠道通路/客户关系/收入来源/核心资源/关键业务/重要伙伴/成本结构。）")

        self.add_h2("（二）盈利模式")
        streams = bm.get("revenue_streams", [])
        if isinstance(streams, str):
            streams = [streams]
        if streams:
            for i, s in enumerate(streams, 1):
                self.add_para(f"{i}. {s}")
        else:
            self.add_para("（请填写盈利模式，1~3 种收入来源，每种含定价+计算依据+Year1 占营收比。）")

        self.add_h2("（三）定价策略")
        pricing = bm.get("pricing", "")
        self.add_para(pricing if pricing else "（请填写定价策略：成本加成/竞品对标/用户价值，含单品价格/月费/年费/竞品对照/用户付费意愿。）")

    # --------------------------------------------------------
    # 运营方案
    # --------------------------------------------------------

    def _add_operation_plan(self):
        """六、运营方案（500~800 字，3 子节 + 表格）"""
        self.add_h1("六、运营方案")
        op = self._get("operations_plan", default={})
        if not isinstance(op, dict):
            op = {}

        self.add_h2("（一）获客策略")
        acq = op.get("acquisition", "")
        self.add_para(acq if acq else "（请填写获客策略：渠道、CAC 估算、3 阶段获客计划。）")

        self.add_h2("（二）运营节奏")
        schedule = op.get("schedule", [])
        if schedule:
            rows = [[s.get("month", ""), s.get("action", ""), s.get("target", "")] for s in schedule if isinstance(s, dict)]
            self.add_table(["月份", "关键动作", "目标指标"], rows, col_widths=[2.5, 7.0, 6.5])
        else:
            self.add_para("（请填写按月运营节奏表：月份/关键动作/目标指标。）")

        self.add_h2("（三）关键指标")
        kpi = op.get("kpi", "")
        self.add_para(kpi if kpi else "（请填写北极星指标 + 次级指标 + 目标值。如 MATU、注册→认证转化率、月复购率等。）")

    # --------------------------------------------------------
    # 财务预测
    # --------------------------------------------------------

    def _add_financial_forecast(self):
        """七、财务预测（3 年表格）"""
        self.add_h1("七、财务预测")
        fin = self._get("financial_forecast", default={})
        if not isinstance(fin, dict):
            fin = {}

        rows_data = fin.get("rows", [])
        if rows_data:
            rows = [[r.get("item", ""), str(r.get("y1", "")), str(r.get("y2", "")), str(r.get("y3", ""))] for r in rows_data if isinstance(r, dict)]
            self.add_table(["财务项目", "Year1", "Year2", "Year3"], rows, col_widths=[5.0, 3.5, 3.5, 3.5])
        else:
            self.add_para("（请填写 3 年财务预测表，至少含：注册用户/月活用户/营业收入/营业成本/毛利率/净利润/净利率 7 行。）")

        self.add_h2("盈亏平衡点说明")
        be = fin.get("breakeven", "")
        self.add_para(be if be else "（请填写盈亏平衡点：YearX 第 Y 个月月营收达 Z 万元时毛利率覆盖固定成本。）")

    # --------------------------------------------------------
    # 团队介绍
    # --------------------------------------------------------

    def _add_team_intro(self):
        """八、团队介绍（300~500 字 + 表格）"""
        self.add_h1("八、团队介绍")
        team = self._get("team_intro", default={})
        if not isinstance(team, dict):
            team = {}

        self.add_h2("（一）团队构成")
        members = team.get("members", [])
        if members:
            rows = [[m.get("name", ""), m.get("id", ""), m.get("major", ""), m.get("role", ""), m.get("exp", "")] for m in members if isinstance(m, dict)]
            self.add_table(["姓名", "学号", "专业年级", "分工", "相关经历"], rows, col_widths=[1.8, 2.5, 3.0, 3.0, 5.7])
        else:
            self.add_para("（请填写团队成员表：姓名/学号/专业年级/分工/相关经历，4~5 人，商科+技术+设计复合。）")

        self.add_h2("（二）指导教师背景")
        adv = team.get("advisor_bg", "")
        self.add_para(adv if adv else "（请填写指导教师背景：职称、研究方向、主持项目、指导学生创业经历。）")

    # --------------------------------------------------------
    # 风险分析
    # --------------------------------------------------------

    def _add_risk_analysis(self):
        """九、风险分析与应对（400~600 字 + 表格）"""
        self.add_h1("九、风险分析与应对")
        risks = self._get("risk_analysis", default=[])
        if isinstance(risks, str):
            risks = [risks]
        if risks:
            rows = [[r.get("type", ""), r.get("risk", ""), r.get("prob", ""), r.get("impact", ""), r.get("measure", "")] for r in risks if isinstance(r, dict)]
            self.add_table(["风险类型", "具体风险", "概率", "影响", "应对措施"], rows, col_widths=[2.2, 3.8, 1.5, 1.5, 7.0])
        else:
            self.add_para("（请填写风险分析表，4 类风险齐全：市场/技术/运营/财务，每类含具体风险+概率+影响+应对措施。）")

    # --------------------------------------------------------
    # 预期成果
    # --------------------------------------------------------

    def _add_expected_results(self):
        """十、预期成果（必须可量化，创业训练型）"""
        self.add_h1("十、预期成果")
        outcomes = self._get("expected_outcomes", default=[])
        if isinstance(outcomes, str):
            outcomes = [outcomes]
        if outcomes:
            for o in outcomes:
                self.add_para(f"• {o}", indent=False)
        else:
            self.add_para("（请填写预期成果，以商业计划书+模拟运营报告为主，每项含数量+形态+验收标准。如：商业计划书 1 份（约 3 万字）、MVP 1 套、模拟运营报告 1 份、软著 1 项。）", indent=False)

    # --------------------------------------------------------
    # 经费预算
    # --------------------------------------------------------

    def _add_budget(self):
        """十一、经费预算（3 列表格：科目/金额/计算依据）"""
        self.add_h1("十一、经费预算")
        items = self._get("budget_items", default=[])
        if not items:
            self.add_para("（请填写经费预算，6 类标准科目：服务器与域名/用户调研/推广获客/物料与会议/印刷复印/软著申请。每项金额非整数，附计算依据。）")
            return
        rows = []
        total = 0
        for b in items:
            if not isinstance(b, dict):
                continue
            try:
                amount_num = int(str(b.get("amount", "0")))
            except ValueError:
                amount_num = 0
            total += amount_num
            rows.append([b.get("item", ""), f"{amount_num} 元", b.get("basis", "")])
        rows.append(["合计", f"{total} 元", ""])
        self.add_table(["预算科目", "金额", "计算依据"], rows, col_widths=[3.5, 3.0, 9.5])

    # --------------------------------------------------------
    # 签字栏
    # --------------------------------------------------------

    def _add_signature_section(self):
        """十二/十三、指导教师意见、学院评审意见"""
        sections = [("十二、指导教师意见", "指导教师签字", "指导教师签字：____________________    日期：______年____月____日")]
        sections.append(("十三、学院评审意见", "学院盖章", "学院盖章：____________________    日期：______年____月____日"))
        if self._get("include_school_approval", default=False):
            sections.append(("十四、学校审批意见", "学校盖章", "学校盖章：____________________    日期：______年____月____日"))
        for title, _, line in sections:
            self.add_h1(title)
            for _ in range(6):
                self.doc.add_paragraph()
            self.add_para(line, indent=False)

    # --------------------------------------------------------
    # 主构建方法
    # --------------------------------------------------------

    def build(self, data: Dict[str, Any], output_path: str) -> str:
        """主构建方法：编排 13 栏目，生成 docx。返回实际保存路径。"""
        try:
            self.data = data if isinstance(data, dict) else {}
            self._validate_data()
            self._add_cover()
            self._add_basic_info_table()
            self._add_abstract()
            self._add_market_analysis()
            self._add_product_service()
            self._add_business_model()
            self._add_operation_plan()
            self._add_financial_forecast()
            self._add_team_intro()
            self._add_risk_analysis()
            self._add_expected_results()
            self._add_budget()
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
        p0_fields = [("project_name", "项目名称"), ("leader_name", "负责人姓名"), ("advisor_name", "指导教师姓名"), ("college", "所在学院")]
        for key, name in p0_fields:
            if not self._get(key):
                warnings.append(f"缺少 {name}（{key}）")

        basic = self._get("basic_info", default={})
        if not isinstance(basic, dict):
            basic = {}
        if not basic.get("industry") and not self._get("industry"):
            warnings.append("缺少 所属行业（industry）")
        if not basic.get("duration") and not self._get("duration"):
            warnings.append("缺少 起止时间（duration）")

        for key, name in [("abstract", "项目简介"), ("market_analysis", "市场分析"), ("business_model", "商业模式"), ("financial_forecast", "财务预测")]:
            if not self._get(key):
                warnings.append(f"缺少 {name}（{key}），将使用占位文本")

        items = self._get("budget_items", default=[])
        if items:
            total = 0
            for b in items:
                if isinstance(b, dict):
                    try:
                        total += int(b.get("amount", 0))
                    except (ValueError, TypeError):
                        pass
            try:
                budget_total_num = int(str(self._get("budget_total", default="")).strip())
            except ValueError:
                budget_total_num = -1
            if budget_total_num >= 0 and total != budget_total_num:
                warnings.append(f"预算合计 {total} 元 与申请经费 {budget_total_num} 元不一致")

        if warnings:
            print("⚠️ 数据校验警告：", file=sys.stderr)
            for w in warnings:
                print(f"  - {w}", file=sys.stderr)
        return warnings


# ============================================================
# 默认示例数据
# ============================================================

DEFAULT_DATA = {
    "project_name": "校园闲置物品流转平台'易舍'",
    "project_level": "国家级",
    "project_type": "创业训练项目",
    "leader_name": "张三",
    "leader_id": "202212345",
    "leader_major": "工商管理",
    "leader_grade": "2022 级 大三",
    "leader_phone": "138XXXXXXXX",
    "advisor_name": "李教授",
    "advisor_title": "副教授",
    "advisor_research": "创业管理",
    "college": "工商管理学院",
    "apply_date": "2025 年 3 月 15 日",
    "basic_info": {
        "project_name": "校园闲置物品流转平台'易舍'",
        "project_type": "创业训练项目",
        "project_source": "A 学生自主选题",
        "industry": "互联网/二手交易",
        "duration": "2025.04-2026.03（12 个月）",
        "budget": "10000",
        "leader_info": "张三 / 202212345 / 工商管理 / 大三 / 138XXXXXXXX",
        "team_members": "李四（产品经理）、王五（前端开发）、赵六（市场运营）、孙七（UI 设计）",
        "advisor_info": "李教授 / 副教授 / 创业管理 / 139XXXXXXXX",
    },
    "abstract": "本项目针对高校校园闲置物品流转效率低、传统跳蚤市场信息不对称的痛点，开发基于微信小程序的校园闲置物品流转平台'易舍'。通过 LBS 同校面交 + 校园认证 + 信用积分机制，服务高校 18~24 岁在校生，预期上线 6 个月覆盖本市 5 所高校、月活 5000+、年交易额 50 万元。采用'交易抽佣 5% + 增值会员 9.9 元/月'盈利模式，预期产出商业计划书 1 份、模拟运营报告 1 份、MVP 1 套。项目已完成 50 份用户调研、MVP 原型上线、首批 200 名内测用户。",
    "market_analysis": {
        "industry_background": "2024 年中国二手交易市场规模达 5000 亿元（艾瑞咨询《2024 中国二手电商报告》），年均增速 20%+，其中校园场景占比约 5%（约 250 亿元）。国家发改委 2024 年《促进绿色消费实施方案》明确鼓励校园二手物品循环利用。",
        "tam_sam_som": [
            {"level": "TAM", "definition": "全国高校在校生 4700 万，人均年闲置消费 530 元",
             "scale": "250 亿元/年"},
            {"level": "SAM", "definition": "本省 50 所高校 120 万在校生",
             "scale": "6.4 亿元/年"},
            {"level": "SOM", "definition": "本市 5 所高校 12 万在校生，预期渗透率 20%",
             "scale": "1272 万元/年"},
        ],
        "competitors": [
            {"name": "闲鱼", "positioning": "全品类二手", "users": "5 亿注册",
             "advantage": "流量大、品牌强", "disadvantage": "校园场景弱、信任度低"},
            {"name": "转转", "positioning": "数码 3C 二手", "users": "2 亿注册",
             "advantage": "验机服务完善", "disadvantage": "非校园场景"},
            {"name": "校园跳蚤群", "positioning": "线下/微信群", "users": "不可统计",
             "advantage": "信任度高", "disadvantage": "信息散乱、效率低"},
        ],
        "user_persona": "典型用户'小张'，大三在校生，月生活费 2000 元，每年闲置物品价值约 500 元，曾尝试闲鱼但因跨城交易邮费高、面交难放弃。访谈 50 名目标用户，42 人表示愿意为'同校面交+信用背书'付费，付费意愿 9.9 元/月接受度 68%。",
    },
    "product_service": {
        "features": [
            "校园认证：学信网/校园邮箱认证，解决陌生人交易信任问题，预期认证率 ≥80%。",
            "LBS 同校匹配：基于地理位置匹配同校买卖双方，预期面交转化率 ≥60%。",
            "信用积分：交易评价+实名认证+校园认证累计积分，积分高者优先展示。",
            "一键发布：拍照识别+AI 描述生成，发布时间从 5 分钟降至 30 秒。",
            "担保交易：买家确认收货后打款，降低交易风险。",
        ],
        "demo": "核心功能'LBS 同校匹配'演示：用户发布闲置物品后，系统自动匹配同校潜在买家，买家可在小程序内查看物品位置（精确到教学楼）、发起面交邀约、双方约定时间地点，完成后互相评价。整个流程 5 步以内完成。",
        "tech_impl": "前端：微信小程序（Taro 3.x）；后端：Node.js + Express；数据库：MySQL 8.0 + Redis 7.0；第三方：腾讯云 LBS、阿里云 OSS、微信支付。研发投入：3 人 × 4 月 = 12 人月，约 6 万元（按校内外包价折算）。",
    },
    "business_model": {
        "canvas": [
            {"element": "客户细分", "content": "高校 18~24 岁在校生，月生活费 1500~3000 元"},
            {"element": "价值主张", "content": "同校面交、信用背书、一键发布"},
            {"element": "渠道通路", "content": "微信小程序、校园社团合作、社群裂变"},
            {"element": "客户关系", "content": "自助服务+社群运营+客服介入"},
            {"element": "收入来源", "content": "交易抽佣 5%+增值会员 9.9 元/月+广告位"},
            {"element": "核心资源", "content": "小程序代码、用户数据、校园合作关系"},
            {"element": "关键业务", "content": "产品开发、用户运营、商家拓展"},
            {"element": "重要伙伴", "content": "高校社团、腾讯云、顺丰快递"},
            {"element": "成本结构", "content": "服务器 1500/月、获客 5 元/人、人力 0"},
        ],
        "revenue_streams": [
            "交易抽佣：每笔成交收取 5% 服务费，参考闲鱼免费、转转 1%，本项目因校园场景溢价定 5%。预期 Year1 占营收 70%。",
            "增值会员：9.9 元/月，含优先展示、信用加权、免抽佣券。访谈数据显示 68% 用户接受。预期 Year1 占营收 20%。",
            "校园商家广告：本地商家投放，CPM 50 元。预期 Year1 占营收 10%。",
        ],
        "pricing": "对照竞品定价：闲鱼免费、转转 1% 抽佣。本项目定 5% 抽佣 + 9.9 元/月会员。基于 50 份用户调研，68% 用户接受 9.9 元/月，付费转化率预期 5%。CAC 估算 5 元/人，LTV 估算 48 元，LTV/CAC = 9.6（健康）。",
    },
    "operations_plan": {
        "acquisition": "获客策略分 3 阶段：① 种子期（M1-2）校园社团合作+线下地推，CAC 8 元；② 增长期（M3-6）社群裂变+口碑传播，CAC 5 元；③ 商业化期（M7-12）付费投放+异业合作，CAC 10 元。整体 CAC 加权 6 元，LTV 48 元，LTV/CAC = 8。",
        "schedule": [
            {"month": "M1", "action": "MVP 上线、首批 5 校推广", "target": "注册 1000、MAU 300"},
            {"month": "M2-3", "action": "社团合作、裂变活动", "target": "注册 5000、MAU 1500"},
            {"month": "M4-6", "action": "商家拓展、付费会员上线", "target": "注册 2 万、付费率 5%"},
            {"month": "M7-12", "action": "跨城复制、商业化优化", "target": "注册 5 万、月营收 5 万"},
        ],
        "kpi": "北极星指标：月活跃交易用户数（MATU）。次级指标：注册→认证转化率 ≥80%、认证→发布转化率 ≥40%、发布→成交转化率 ≥30%、月复购率 ≥25%、付费会员率 ≥5%。",
    },
    "financial_forecast": {
        "rows": [
            {"item": "注册用户（万）", "y1": "2", "y2": "8", "y3": "20"},
            {"item": "月活用户（万）", "y1": "0.5", "y2": "2", "y3": "6"},
            {"item": "营业收入（万元）", "y1": "12", "y2": "60", "y3": "200"},
            {"item": "营业成本（万元）", "y1": "8", "y2": "35", "y3": "100"},
            {"item": "毛利率", "y1": "33%", "y2": "42%", "y3": "50%"},
            {"item": "期间费用（万元）", "y1": "7", "y2": "17", "y3": "40"},
            {"item": "净利润（万元）", "y1": "-3", "y2": "8", "y3": "60"},
            {"item": "净利率", "y1": "-25%", "y2": "13%", "y3": "30%"},
        ],
        "breakeven": "盈亏平衡点：Year2 第 6 个月。届时月营收达 7.5 万元，毛利率 42% 覆盖月固定成本 3 万元（服务器 1500 + 客服 500 + 推广 2000 + 杂项 1000）。固定成本/毛利率 = 3/42% = 7.14 万元。",
    },
    "team_intro": {
        "members": [
            {"name": "张三", "id": "202212345", "major": "工商管理 大三",
             "role": "项目负责人/战略", "exp": "校 SRT 项目负责人、电商实习"},
            {"name": "李四", "id": "202212346", "major": "计算机 大三",
             "role": "产品经理", "exp": "字节产品实习、App 上线 1 款"},
            {"name": "王五", "id": "202212347", "major": "软件工程 大三",
             "role": "前端开发", "exp": "全栈开发 2 年、GitHub 1k star"},
            {"name": "赵六", "id": "202212348", "major": "市场营销 大二",
             "role": "市场运营", "exp": "校园社群运营 5000+ 用户"},
            {"name": "孙七", "id": "202212349", "major": "视觉传达 大三",
             "role": "UI 设计", "exp": "红点设计奖入围"},
        ],
        "advisor_bg": "李教授，副教授，工商管理学院，研究方向创业管理与商业模式创新，主持教育部人文社科项目 1 项，指导学生团队获'互联网+'省赛金奖 2 项。",
    },
    "risk_analysis": [
        {"type": "市场风险", "risk": "闲鱼/转转进入校园",
         "prob": "中", "impact": "高", "measure": "强化校园社团合作、深耕单校"},
        {"type": "技术风险", "risk": "高并发性能瓶颈",
         "prob": "低", "impact": "中", "measure": "腾讯云弹性扩容、Redis 缓存"},
        {"type": "运营风险", "risk": "用户增长不及预期",
         "prob": "中", "impact": "高", "measure": "多渠道获客、降低 CAC"},
        {"type": "财务风险", "risk": "现金流断裂",
         "prob": "低", "impact": "高", "measure": "控制 burn rate、申请延期经费"},
    ],
    "expected_outcomes": [
        "商业计划书 1 份（约 3 万字，含财务模型）",
        "模拟运营报告 1 份（含 12 个月运营数据）",
        "MVP 原型 1 套（已上线，注册用户 2 万+）",
        "用户调研报告 1 份（50 份深度访谈 + 500 份问卷）",
        "路演 PPT 1 套（20 页，用于校内/省级路演）",
        "软件著作权 1 项（《易舍校园闲置流转平台 V1.0》）",
    ],
    "budget_items": [
        {"item": "服务器与域名", "amount": "1800",
         "basis": "腾讯云轻量服务器 99 元/月 × 12 月 + 域名 180 元 + SSL 432 元"},
        {"item": "用户调研", "amount": "1200",
         "basis": "50 份深度访谈 × 20 元 + 500 份问卷红包 2 元/份"},
        {"item": "推广获客", "amount": "3500",
         "basis": "5 校 × 500 元线下活动 + 1000 元社群裂变红包"},
        {"item": "物料与会议", "amount": "1500",
         "basis": "海报印刷 800 + 路演物料 700"},
        {"item": "印刷复印", "amount": "800",
         "basis": "商业计划书印刷 30 份 × 25 元 + 报告 50 元"},
        {"item": "软著申请", "amount": "1200",
         "basis": "软件著作权申请费 1 项"},
    ],
    "team_foundation": "团队 5 名成员已修读《创业管理》《市场营销》《软件工程》《视觉设计》等核心课程，3 人有校级大创参与经验。",
    "budget_total": "10000",
    "include_school_approval": False,
}


# ============================================================
# CLI 入口
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="大创-创业训练项目申报书 docx 生成器",
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
