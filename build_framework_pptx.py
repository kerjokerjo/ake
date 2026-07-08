# -*- coding: utf-8 -*-
"""生成"图 2-1 本课题总体框架图"的可编辑 PPT 图表。
结构：总体目标 → 三个现实问题（并列，色块对应）→ 三项研究内容（并列，色块对应，
研究内容二内嵌"数据知识化→知识价值化"技术路线）→ 预期成果 ← 理论与技术支撑（基础支撑）。
全部使用 PPT 原生形状（矩形、箭头、连接线、文本框），可在 PowerPoint / WPS 中直接编辑
文字、颜色、位置与大小。"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

# 配色
COL_TITLE_FILL = RGBColor(0x2E, 0x4A, 0x62)
COL_GOAL_FILL = RGBColor(0xFC, 0xE4, 0xD6)
COL_GOAL_BORDER = RGBColor(0xC5, 0x5A, 0x11)
COL_OUTCOME_FILL = RGBColor(0xFC, 0xE4, 0xD6)
COL_OUTCOME_BORDER = RGBColor(0xC5, 0x5A, 0x11)
COL_BASE_FILL = RGBColor(0xE2, 0xEF, 0xDA)
COL_BASE_BORDER = RGBColor(0x54, 0x82, 0x35)
COL_PROBLEM_FILL = RGBColor(0xFF, 0xF2, 0xCC)
COL_CONTENT_FILL = RGBColor(0xDE, 0xEA, 0xF6)
COL_ARROW = RGBColor(0x59, 0x59, 0x59)
COL_TEXT_DARK = RGBColor(0x26, 0x26, 0x26)
COL_WHITE = RGBColor(0xFF, 0xFF, 0xFF)

COL_ACCENT = [RGBColor(0x44, 0x72, 0xC4),  # 蓝 - 问题一/内容一
              RGBColor(0xED, 0x7D, 0x31),  # 橙 - 问题二/内容二
              RGBColor(0x54, 0x82, 0x35)]  # 绿 - 问题三/内容三

FONT_CN = '微软雅黑'

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height

slide = prs.slides.add_slide(prs.slide_layouts[6])  # 空白版式


def set_run(run, text, size=12, bold=False, color=COL_TEXT_DARK, font=FONT_CN):
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    rPr = run._r.get_or_add_rPr()
    ea = rPr.makeelement(qn('a:ea'), {'typeface': font})
    rPr.append(ea)


def add_textbox(x, y, w, h, lines, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE):
    tb = slide.shapes.add_textbox(x, y, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Pt(2)
    tf.margin_right = Pt(2)
    tf.margin_top = Pt(1)
    tf.margin_bottom = Pt(1)
    for i, (text, size, bold, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = 1.05
        r = p.add_run()
        set_run(r, text, size=size, bold=bold, color=color)
    return tb


def add_box(shape_type, x, y, w, h, fill, border_color, border_w=1.75, lines=None,
            align=PP_ALIGN.CENTER, radius=None):
    sp = slide.shapes.add_shape(shape_type, x, y, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = fill
    sp.line.color.rgb = border_color
    sp.line.width = Pt(border_w)
    sp.shadow.inherit = False
    if lines:
        tf = sp.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = Pt(4)
        tf.margin_right = Pt(4)
        tf.margin_top = Pt(2)
        tf.margin_bottom = Pt(2)
        for i, (text, size, bold, color) in enumerate(lines):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.alignment = align
            p.line_spacing = 1.0
            r = p.add_run()
            set_run(r, text, size=size, bold=bold, color=color)
    return sp


def add_down_arrow(x, y, w, h, color=COL_ARROW):
    sp = slide.shapes.add_shape(MSO_SHAPE.DOWN_ARROW, x, y, w, h)
    sp.fill.solid()
    sp.fill.fore_color.rgb = color
    sp.line.fill.background()
    sp.shadow.inherit = False
    return sp


def add_line(x1, y1, x2, y2, color=COL_ARROW, width=1.5, dash=None):
    conn = slide.shapes.add_connector(1, x1, y1, x2, y2)  # 1 = straight
    conn.line.color.rgb = color
    conn.line.width = Pt(width)
    if dash:
        ln = conn.line._get_or_add_ln()
        d = ln.makeelement(qn('a:prstDash'), {'val': dash})
        ln.append(d)
    return conn


# ============================================================
# 标题
# ============================================================
add_textbox(Inches(0.4), Inches(0.20), Inches(12.5), Inches(0.5),
            [('图 2-1  本课题总体框架图', 22, True, COL_TITLE_FILL)])

# ============================================================
# 总体目标
# ============================================================
goal_w, goal_h = Inches(10.6), Inches(0.68)
goal_x = (SW - goal_w) // 2
goal_y = Inches(0.82)
add_box(MSO_SHAPE.ROUNDED_RECTANGLE, goal_x, goal_y, goal_w, goal_h,
        COL_GOAL_FILL, COL_GOAL_BORDER, border_w=2,
        lines=[('总体目标：构建巴渝优秀传统文化资源知识图谱，揭示价值内涵，提出传播创新路径',
                12.5, True, COL_GOAL_BORDER)])

# 箭头：总体目标 → 三问题
arrow_w, arrow_h = Inches(0.42), Inches(0.20)
add_down_arrow((SW - arrow_w) // 2, goal_y + goal_h + Inches(0.02), arrow_w, arrow_h)

# ============================================================
# 三个现实问题（并列）
# ============================================================
col_w = Inches(3.35)
gap = Inches(0.42)
total_w = col_w * 3 + gap * 2
start_x = (SW - total_w) // 2
prob_y = goal_y + goal_h + Inches(0.34)
prob_h = Inches(0.72)

problems = [
    ('问题一', '资源分散、标准不一'),
    ('问题二', '关联薄弱、阐释不深'),
    ('问题三', '转化不足、场景不清'),
]
col_x = []
for i, (tag, text) in enumerate(problems):
    x = start_x + i * (col_w + gap)
    col_x.append(x)
    add_box(MSO_SHAPE.ROUNDED_RECTANGLE, x, prob_y, col_w, prob_h,
            COL_PROBLEM_FILL, COL_ACCENT[i], border_w=2.25,
            lines=[(tag, 12, True, COL_ACCENT[i]),
                   (text, 11, False, COL_TEXT_DARK)])

# 三个下箭头
arrow2_y = prob_y + prob_h + Inches(0.02)
arrow2_h = Inches(0.18)
for i in range(3):
    ax = col_x[i] + col_w // 2 - Inches(0.20) // 2
    add_down_arrow(ax, arrow2_y, Inches(0.20), arrow2_h, color=COL_ACCENT[i])

# ============================================================
# 三项研究内容（并列，内嵌四阶段技术路线）
# ============================================================
content_y = arrow2_y + arrow2_h + Inches(0.12)
content_h = Inches(1.85)

contents = [
    [('研究内容一', 12, True, COL_ACCENT[0]),
     ('资源数据化整理', 11.5, True, COL_TEXT_DARK),
     ('资源普查→分类体系→数据规范→基础数据目录', 10, False, COL_TEXT_DARK)],
    [('研究内容二', 12, True, COL_ACCENT[1]),
     ('知识图谱构建与价值挖掘', 11.5, True, COL_TEXT_DARK),
     ('本体建模→实体识别→关系抽取→知识融合→图谱原型', 10, False, COL_TEXT_DARK),
     ('→时空/主题/关系网络分析→价值内涵与精神标识', 10, False, COL_TEXT_DARK)],
    [('研究内容三', 12, True, COL_ACCENT[2]),
     ('传播创新', 11.5, True, COL_TEXT_DARK),
     ('教育／文旅／公共文化服务／网络传播四类场景传播路径设计', 10, False, COL_TEXT_DARK)],
]
for i in range(3):
    add_box(MSO_SHAPE.ROUNDED_RECTANGLE, col_x[i], content_y, col_w, content_h,
            COL_CONTENT_FILL, COL_ACCENT[i], border_w=2.25, lines=contents[i])

# ============================================================
# 汇聚连接线 + "共同支撑" + 下箭头 → 预期成果
# ============================================================
merge_y = content_y + content_h + Inches(0.14)
mid_x = SW // 2
# 三条竖线（各列中心到汇聚横线）
for i in range(3):
    cx = col_x[i] + col_w // 2
    add_line(cx, content_y + content_h, cx, merge_y, color=COL_ACCENT[i], width=1.5)
# 汇聚横线
add_line(col_x[0] + col_w // 2, merge_y, col_x[2] + col_w // 2, merge_y, color=COL_ARROW, width=1.5)

add_textbox(mid_x - Inches(0.9), merge_y + Inches(0.01), Inches(1.8), Inches(0.24),
            [('共同支撑', 10.5, True, COL_ARROW)])

arrow3_y = merge_y + Inches(0.24)
add_down_arrow(mid_x - Inches(0.21), arrow3_y, Inches(0.42), Inches(0.20))

# ============================================================
# 预期成果
# ============================================================
outcome_y = arrow3_y + Inches(0.26)
outcome_h = Inches(0.56)
add_box(MSO_SHAPE.ROUNDED_RECTANGLE, goal_x, outcome_y, goal_w, outcome_h,
        COL_OUTCOME_FILL, COL_OUTCOME_BORDER, border_w=2,
        lines=[('预期成果：学术论文 · 研究报告 · 知识图谱原型 · 可视化成果', 11.5, True, COL_OUTCOME_BORDER)],
        align=PP_ALIGN.CENTER)

# ============================================================
# 理论与技术支撑（基础，位于底部，向上支撑预期成果）
# ============================================================
base_y = outcome_y + outcome_h + Inches(0.26)
base_h = Inches(0.52)
add_line(mid_x, outcome_y + outcome_h, mid_x, base_y, color=COL_BASE_BORDER, width=1.5, dash='dash')
# 向上箭头
add_box(MSO_SHAPE.UP_ARROW, mid_x - Inches(0.21), outcome_y + outcome_h + Inches(0.02), Inches(0.42), Inches(0.18),
        COL_BASE_BORDER, COL_BASE_BORDER, border_w=0)

add_box(MSO_SHAPE.ROUNDED_RECTANGLE, goal_x, base_y, goal_w, base_h,
        COL_BASE_FILL, COL_BASE_BORDER, border_w=2,
        lines=[('理论与技术支撑：数字人文理论 · 信息资源管理 · 知识图谱技术 · 巴渝文化研究', 11, True, COL_BASE_BORDER)])

# ============================================================
# 图注
# ============================================================
note_y = base_y + base_h + Inches(0.10)
add_textbox(Inches(0.5), note_y, Inches(12.3), Inches(0.5),
            [('说明：三个现实问题与三项研究内容按颜色纵向对应（蓝—问题一/内容一，橙—问题二/内容二，绿—问题三/内容三）；'
              '研究内容二内嵌"数据知识化→知识价值化"技术路线；三项内容共同支撑总体目标实现，理论与技术支撑贯穿全过程。',
              9.5, False, RGBColor(0x59, 0x59, 0x59))],
            align=PP_ALIGN.CENTER)

OUT = '图2-1-本课题总体框架图.pptx'
prs.save(OUT)
print('saved:', OUT)
