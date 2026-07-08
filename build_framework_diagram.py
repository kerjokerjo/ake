# -*- coding: utf-8 -*-
"""生成"图 2-1 本课题总体框架图"的原生 Word 图表（单独文件，可直接插入申报书
"图 2-1 总体框架图"占位处）。
结构：总体目标 → 三个现实问题（并列） → 三项研究内容（并列，与问题纵向对应，
研究内容二内嵌四阶段技术路线） → 预期成果 → 理论与技术支撑（基础支撑）。
三组问题/内容按列使用同一强调色的边框，形成纵向视觉关联。"""
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

CENTER = WD_ALIGN_PARAGRAPH.CENTER
CONTENT_W = 16.0  # 可用宽度（cm）

doc = Document()
normal = doc.styles['Normal']
normal.font.name = '仿宋_GB2312'
normal.font.size = Pt(12)
normal._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')

sec = doc.sections[0]
sec.top_margin = Cm(2.0)
sec.bottom_margin = Cm(2.0)
sec.left_margin = Cm(2.2)
sec.right_margin = Cm(2.2)
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)


def set_cn_font(run, cn='仿宋_GB2312', size=10.5, bold=False, en='Times New Roman', color=None):
    run.font.name = en
    run.font.size = Pt(size)
    run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)
    rpr = run._element.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    rf.set(qn('w:eastAsia'), cn)
    rf.set(qn('w:ascii'), en)
    rf.set(qn('w:hAnsi'), en)


def para(text='', size=12, bold=False, align=CENTER, cn='仿宋_GB2312', space_after=6, space_before=0):
    p = doc.add_paragraph()
    p.alignment = align
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    if text:
        r = p.add_run(text)
        set_cn_font(r, cn=cn, size=size, bold=bold)
    return p


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)


def cell_borders(cell, color='808080', sz='10'):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        e = OxmlElement('w:' + edge)
        e.set(qn('w:val'), 'single')
        e.set(qn('w:sz'), sz)
        e.set(qn('w:space'), '0')
        e.set(qn('w:color'), color)
        b.append(e)
    tcPr.append(b)


def set_cell_text(cell, lines, size=10, valign='center'):
    """lines: list of (text, bold, size_override or None, color or None)"""
    cell.text = ''
    tcPr = cell._tc.get_or_add_tcPr()
    va = OxmlElement('w:vAlign')
    va.set(qn('w:val'), valign)
    tcPr.append(va)
    p = cell.paragraphs[0]
    p.alignment = CENTER
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(size + 5)
    for i, (text, bold, sz, color) in enumerate(lines):
        if i > 0:
            p = cell.add_paragraph()
            p.alignment = CENTER
            p.paragraph_format.space_before = Pt(1)
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            p.paragraph_format.line_spacing = Pt((sz or size) + 4)
        r = p.add_run(text)
        set_cn_font(r, size=sz or size, bold=bold, color=color)


def full_box(lines, fill, border_color='808080', height_cm=1.4):
    tb = doc.add_table(rows=1, cols=1)
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = tb.rows[0].cells[0]
    c.width = Cm(CONTENT_W)
    shade(c, fill)
    cell_borders(c, color=border_color, sz='10')
    set_cell_text(c, lines, size=11)
    return tb


def arrow_line(symbol='▼', size=14, space_after=2, space_before=2, color='595959'):
    p = doc.add_paragraph()
    p.alignment = CENTER
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(space_before)
    r = p.add_run(symbol)
    set_cn_font(r, cn='宋体', size=size, bold=True, color=color)


def three_col_arrows(colors):
    tb = doc.add_table(rows=1, cols=3)
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    for j in range(3):
        c = tb.rows[0].cells[j]
        c.width = Cm(CONTENT_W / 3)
        c.text = ''
        p = c.paragraphs[0]
        p.alignment = CENTER
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.space_before = Pt(2)
        r = p.add_run('▼')
        set_cn_font(r, cn='宋体', size=14, bold=True, color=colors[j])
    return tb


def three_col_boxes(items, fill, colors, height_cm=2.6, size=9.5):
    """items: list of 3 -> list of (text, bold, sz, color) lines."""
    tb = doc.add_table(rows=1, cols=3)
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    tb.autofit = False
    for j in range(3):
        c = tb.rows[0].cells[j]
        c.width = Cm(CONTENT_W / 3 - 0.15)
        shade(c, fill)
        cell_borders(c, color=colors[j], sz='12')
        set_cell_text(c, items[j], size=size)
    return tb


COL_COLORS = ['4472C4', 'ED7D31', '548235']  # 蓝 / 橙 / 绿，纵向三列对应色

# ============================================================
para('图 2-1  本课题总体框架图', size=13, bold=True, cn='黑体', space_after=10, space_before=0)

# 总体目标
full_box(
    [('总体目标', True, 11.5, None),
     ('构建巴渝优秀传统文化资源知识图谱，揭示其价值内涵，提出传播创新路径', False, 10.5, None)],
    fill='FCE4D6', border_color='C55A11'
)
arrow_line()

# 三个现实问题（并列，色块与下方研究内容纵向对应）
problems = [
    [('问题一', True, 10.5, COL_COLORS[0]), ('资源分散、标准不一', False, 10, None)],
    [('问题二', True, 10.5, COL_COLORS[1]), ('关联薄弱、阐释不深', False, 10, None)],
    [('问题三', True, 10.5, COL_COLORS[2]), ('转化不足、场景不清', False, 10, None)],
]
three_col_boxes(problems, fill='FFF2CC', colors=COL_COLORS, size=10)
three_col_arrows(COL_COLORS)

# 三项研究内容（并列，内嵌四阶段技术路线，与问题纵向对应）
contents = [
    [
        ('研究内容一', True, 10.5, COL_COLORS[0]),
        ('资源数据化整理', True, 10, None),
        ('资源普查 → 分类体系 →', False, 9, None),
        ('数据规范 → 基础数据目录', False, 9, None),
    ],
    [
        ('研究内容二', True, 10.5, COL_COLORS[1]),
        ('知识图谱构建与价值挖掘', True, 10, None),
        ('本体建模→实体识别→关系抽取', False, 9, None),
        ('→知识融合→图谱原型', False, 9, None),
        ('→时空/主题/关系网络分析', False, 9, None),
        ('→价值内涵与精神标识', False, 9, None),
    ],
    [
        ('研究内容三', True, 10.5, COL_COLORS[2]),
        ('传播创新', True, 10, None),
        ('教育／文旅／公共文化服务／', False, 9, None),
        ('网络传播四类场景传播路径设计', False, 9, None),
    ],
]
three_col_boxes(contents, fill='DEEAF6', colors=COL_COLORS, size=9.5)

para('▼　　　　　　　　　　▼　　　　　　　　　　▼', size=13, bold=True, cn='宋体', space_after=0, space_before=2)
para('共同支撑', size=9.5, bold=True, cn='黑体', space_after=2, space_before=0)
arrow_line()

# 预期成果
full_box(
    [('预期成果', True, 11.5, None),
     ('学术论文　·　研究报告　·　知识图谱原型　·　可视化成果', False, 10.5, None)],
    fill='FCE4D6', border_color='C55A11'
)

para('▲', size=13, bold=True, cn='宋体', space_after=2, space_before=4)

# 理论与技术支撑（基础）
full_box(
    [('理论与技术支撑', True, 11, None),
     ('数字人文理论　·　信息资源管理　·　知识图谱技术　·　巴渝文化研究', False, 10, None)],
    fill='E2EFDA', border_color='548235'
)

para('', space_after=0)
para('说明：三个现实问题与三项研究内容按颜色纵向对应（蓝—问题一/内容一，橙—问题二/内容二，绿—问题三/内容三），'
     '研究内容二内嵌"资源数据化→数据知识化→知识价值化→价值传播化"技术路线的中段（数据知识化、知识价值化两个阶段）；'
     '三项内容共同支撑总体目标的实现，理论与技术支撑贯穿研究全过程。',
     size=9, align=CENTER, cn='楷体_GB2312', space_after=0)

OUT = '图2-1-本课题总体框架图.docx'
doc.save(OUT)
print('saved:', OUT)
