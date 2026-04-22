"""
原生可编辑 PPTX 版本 —— 简化版技术路线图
所有方框、箭头、文字均为 PowerPoint 原生形状
打开后可直接点击任一元素修改文字、颜色、位置
"""
from pptx import Presentation
from pptx.util import Cm, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ======== 页面尺寸（宽版 A4 横向） ========
SLIDE_W_CM = 33.87
SLIDE_H_CM = 19.05

prs = Presentation()
prs.slide_width = Cm(SLIDE_W_CM)
prs.slide_height = Cm(SLIDE_H_CM)

slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank

# ======== 配色（与 PNG 版本保持一致） ========
COLORS = {
    'bg':        RGBColor(0xFA, 0xFB, 0xFD),
    'prob':      RGBColor(0xFD, 0xEC, 0xEA),
    'prob_e':    RGBColor(0xC0, 0x39, 0x2B),
    'goal':      RGBColor(0xFF, 0xF4, 0xE0),
    'goal_e':    RGBColor(0xD6, 0x89, 0x10),
    'rq1':       RGBColor(0xE8, 0xF4, 0xFD),
    'rq1_e':     RGBColor(0x28, 0x74, 0xA6),
    'rq2':       RGBColor(0xE8, 0xF8, 0xEF),
    'rq2_e':     RGBColor(0x1E, 0x84, 0x49),
    'rq3':       RGBColor(0xF4, 0xEC, 0xF7),
    'rq3_e':     RGBColor(0x6C, 0x34, 0x83),
    'link':      RGBColor(0x5D, 0x6D, 0x7E),
    'loop':      RGBColor(0xD3, 0x54, 0x00),
    'out':       RGBColor(0xFE, 0xF9, 0xE7),
    'out_e':     RGBColor(0xB7, 0x95, 0x0B),
    'dark':      RGBColor(0x1B, 0x26, 0x31),
    'gray':      RGBColor(0x56, 0x65, 0x73),
    'white':     RGBColor(0xFF, 0xFF, 0xFF),
    'sub':       RGBColor(0x51, 0x5A, 0x5A),
}

# ======== 尺寸工具：把 0-100 百分比坐标 => EMU ========
def X(p): return Emu(int(prs.slide_width  * p / 100))
def Y(p): return Emu(int(prs.slide_height * p / 100))
def W(p): return Emu(int(prs.slide_width  * p / 100))
def H(p): return Emu(int(prs.slide_height * p / 100))

# ======== 辅助函数 ========
def add_rrect(x, y, w, h, fill, line, line_w=1.2, corner=0.08):
    shp = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, X(x), Y(y), W(w), H(h))
    shp.adjustments[0] = corner
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    shp.line.color.rgb = line
    shp.line.width = Pt(line_w)
    shp.shadow.inherit = False
    return shp

def add_rect(x, y, w, h, fill, line=None, line_w=0):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, X(x), Y(y), W(w), H(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is not None and line_w > 0:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp

def add_circle(cx, cy, r, fill, line=None, line_w=0):
    shp = slide.shapes.add_shape(MSO_SHAPE.OVAL, X(cx-r), Y(cy-r*SLIDE_W_CM/SLIDE_H_CM),
                                 W(r*2), H(r*2*SLIDE_W_CM/SLIDE_H_CM))
    shp.fill.solid()
    shp.fill.fore_color.rgb = fill
    if line is not None and line_w > 0:
        shp.line.color.rgb = line
        shp.line.width = Pt(line_w)
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp

def set_text(shape, text, size=11, bold=False, color=None, align='center', anchor='middle'):
    tf = shape.text_frame
    tf.margin_left = Emu(36000)
    tf.margin_right = Emu(36000)
    tf.margin_top = Emu(18000)
    tf.margin_bottom = Emu(18000)
    tf.word_wrap = True
    lines = text.split('\n')
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = {'left': PP_ALIGN.LEFT, 'center': PP_ALIGN.CENTER,
                       'right': PP_ALIGN.RIGHT}[align]
        run = p.add_run()
        run.text = line
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.name = '微软雅黑'
        if color is not None:
            run.font.color.rgb = color
    tf.vertical_anchor = {'top': MSO_ANCHOR.TOP, 'middle': MSO_ANCHOR.MIDDLE,
                          'bottom': MSO_ANCHOR.BOTTOM}[anchor]

def add_textbox(x, y, w, h, text, size=11, bold=False, color=None, align='center', anchor='middle'):
    tb = slide.shapes.add_textbox(X(x), Y(y), W(w), H(h))
    set_text(tb, text, size, bold, color, align, anchor)
    return tb

def add_arrow(x1, y1, x2, y2, color, width=1.6, dash=False):
    conn = slide.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, X(x1), Y(y1), X(x2), Y(y2))
    conn.line.color.rgb = color
    conn.line.width = Pt(width)
    # 添加箭头头部
    ln = conn.line._get_or_add_ln()
    tail_end = etree.SubElement(ln, qn('a:tailEnd'))
    tail_end.set('type', 'triangle')
    tail_end.set('w', 'med')
    tail_end.set('h', 'med')
    return conn

def add_curve_arrow(x1, y1, x2, y2, color, width=1.6):
    """ 通过两个 BENT 连接器模拟弧形反哺回流 """
    conn = slide.shapes.add_connector(MSO_CONNECTOR.CURVE, X(x1), Y(y1), X(x2), Y(y2))
    conn.line.color.rgb = color
    conn.line.width = Pt(width)
    ln = conn.line._get_or_add_ln()
    tail_end = etree.SubElement(ln, qn('a:tailEnd'))
    tail_end.set('type', 'triangle')
    tail_end.set('w', 'med')
    tail_end.set('h', 'med')
    return conn

# ======== 背景 ========
add_rect(0, 0, 100, 100, COLORS['bg'])

# ======== 标题 ========
add_textbox(0, 1.2, 100, 5,
            '图  "三图谱+双智能体" 赋能专业核心课程建设技术路线图（简版）',
            size=20, bold=True, color=COLORS['dark'])
add_textbox(0, 5.6, 100, 3.2,
            '问题 → 目标 → 内容 → 成果 · 四层递进 · 内容层闭环联动',
            size=12, color=COLORS['gray'])

# ======== 层标签 ========
def layer_label(y, text, color, h=9):
    add_rect(1.0, y - h/2, 4.6, h, color)
    add_textbox(1.0, y - h/2, 4.6, h, text,
                size=11, bold=True, color=COLORS['white'])

# =================================================================
# 第 1 层：问题层（y ≈ 11 - 23）
# =================================================================
layer_label(17, '问题层\n三重断裂', COLORS['prob_e'], h=12)

probs = [
    ("① \"图—学\" 断裂",
     "知识/能力图谱未连通\"问题侧\"数据\n学情诊断粗放 · 难以精准定位学习障碍"),
    ("② \"AI—用\" 断裂",
     "通用AI难以嵌入爬虫强实践/对抗/合规场景\n伴学缺位 · 合规悬空 · 重技术轻伦理"),
    ("③ \"教—评\" 断裂",
     "代码/调试/合规数据未被用于过程评价\n德技分离 · 改革经验缺乏可迁移载体"),
]
prob_y, prob_h = 11.5, 11.5
pw = 26
pgap = (100 - 8 - 3*pw) / 2
px0 = 8
prob_centers = []
for i, (title, sub) in enumerate(probs):
    x = px0 + i * (pw + pgap)
    shp = add_rrect(x, prob_y, pw, prob_h, COLORS['prob'], COLORS['prob_e'], line_w=1.8)
    shp.text_frame.text = ''
    add_textbox(x, prob_y + 0.8, pw, 3.5, title, size=14, bold=True,
                color=COLORS['prob_e'])
    add_textbox(x, prob_y + 4.5, pw, 6.5, sub, size=10.5, color=COLORS['sub'])
    prob_centers.append(x + pw/2)

# =================================================================
# 第 2 层：目标层（y ≈ 26 - 35）
# =================================================================
layer_label(30.5, '目标层\n总目标', COLORS['goal_e'], h=9)

goal_y, goal_h = 26.5, 9
shp = add_rrect(8, goal_y, 85, goal_h, COLORS['goal'], COLORS['goal_e'], line_w=1.8)
add_textbox(8, goal_y + 0.5, 85, 3.5,
            '建成以"三图谱 + 双智能体"为核心特征的专业核心课程样板',
            size=14, bold=True, color=COLORS['goal_e'])

tags = ['一标准', '三图谱', '两智能体', '一靶场', '一模式', '一评价']
tag_y = goal_y + 4.8
tag_w = 11.8
total_w = len(tags)*tag_w + (len(tags)-1)*1.5
tx0 = 50 - total_w/2
for i, t in enumerate(tags):
    xi = tx0 + i*(tag_w + 1.5)
    add_rrect(xi, tag_y, tag_w, 3.3, COLORS['white'], COLORS['goal_e'], line_w=1.0, corner=0.2)
    add_textbox(xi, tag_y, tag_w, 3.3, t, size=11, bold=True, color=COLORS['goal_e'])

# 问题 → 目标
for cx in prob_centers:
    add_arrow(cx, prob_y + prob_h, 50, goal_y, COLORS['prob_e'], width=1.0)

# =================================================================
# 第 3 层：内容层（y ≈ 38 - 76）
# =================================================================
layer_label(57, '内容层\n三条主线\n闭环联动', COLORS['rq1_e'], h=14)

rc_y, rc_h = 38.5, 37
rc_w = 24
rc_gap = (100 - 8 - 3*rc_w) / 2
rc_x0 = 8

rc_data = [
    {
        'bg': COLORS['rq1'], 'ec': COLORS['rq1_e'],
        'num': 'Ⅰ', 'tag': '认知底座',
        'title': '"三图谱" 构建与\n课程标准重构',
        'subtitle': '对应：图—学断裂 / KQ①',
        'items': [
            '知识图谱  12 模块 · 86 知识点',
            '能力图谱  127 能力点',
            '问题图谱  2300+ 提问 · 400+ 错误',
            '五级建模  项目→任务→流程→技能→知识',
            '课程标准  3 模块 · 5 项目 · 12 任务',
        ],
    },
    {
        'bg': COLORS['rq2'], 'ec': COLORS['rq2_e'],
        'num': 'Ⅱ', 'tag': '执行引擎',
        'title': '"双智能体" 研发\n与场景嵌入',
        'subtitle': '对应：AI—用断裂 / KQ②',
        'items': [
            '底层引擎  私有化 LLM + RAG + LoRA',
            '智学小助  代码伴学·路径推荐·归因',
            'AI 合规官  法规嵌入·预警·审查',
            '仿真靶场  4 类站点 · 10 级反爬',
            '红蓝对抗  实战演练模式',
        ],
    },
    {
        'bg': COLORS['rq3'], 'ec': COLORS['rq3_e'],
        'num': 'Ⅲ', 'tag': '评价+迁移',
        'title': '"双师+双工单"模式与\n五维评价·迁移推广',
        'subtitle': '对应：教—评断裂 / KQ③',
        'items': [
            '双师+双工单  教师+AI / 教学+企业',
            '五维评价  代码·调试·协作·合规·伦理',
            '伦理 20%   重大违规一票否决',
            '学习分析  精准补偿与过程性闭环',
            '迁移推广  跨课程指南+师资培训包',
        ],
    },
]

rc_boxes = []
for i, d in enumerate(rc_data):
    x = rc_x0 + i*(rc_w + rc_gap)
    add_rrect(x, rc_y, rc_w, rc_h, d['bg'], d['ec'], line_w=2.0)

    # 编号圆
    add_circle(x + 2.8, rc_y + 2.7, 1.7, d['ec'])
    add_textbox(x + 1.0, rc_y + 1.4, 3.6, 2.6, d['num'],
                size=14, bold=True, color=COLORS['white'])

    # 角色标签
    add_rrect(x + rc_w - 8.5, rc_y + 1.3, 7, 2.8, COLORS['white'], d['ec'],
              line_w=1.0, corner=0.2)
    add_textbox(x + rc_w - 8.5, rc_y + 1.3, 7, 2.8, d['tag'],
                size=10, bold=True, color=d['ec'])

    # 标题
    add_textbox(x + 0.5, rc_y + 4.6, rc_w - 1, 4.5, d['title'],
                size=12.5, bold=True, color=d['ec'])

    # 对应关系
    add_textbox(x + 0.5, rc_y + 9.3, rc_w - 1, 2.5, d['subtitle'],
                size=10, color=COLORS['gray'])

    # 分割线
    add_rect(x + 2, rc_y + 12.0, rc_w - 4, 0.12, d['ec'])

    # 子项
    for j, item in enumerate(d['items']):
        y_item = rc_y + 13.2 + j*4.2
        add_circle(x + 2.0, y_item + 1.1, 0.35, d['ec'])
        add_textbox(x + 3.0, y_item, rc_w - 3.5, 2.5, item,
                    size=10.5, color=COLORS['dark'], align='left')

    rc_boxes.append({
        'x_left': x, 'x_right': x + rc_w,
        'cx': x + rc_w/2,
        'top_y': rc_y, 'bot_y': rc_y + rc_h,
        'ec': d['ec'],
    })

# 问题 → 内容
for i in range(3):
    add_arrow(prob_centers[i], prob_y + prob_h + 0.3,
              rc_boxes[i]['cx'], rc_y,
              rc_boxes[i]['ec'], width=1.2)

# 内容层横向联动（顶部）
band_y = rc_y + 5
labels_top = [
    ('数据语义供给', 'RAG 检索源'),
    ('过程数据沉淀', '合规证据链'),
]
for i in range(2):
    xa = rc_boxes[i]['x_right'] + 0.3
    xb = rc_boxes[i+1]['x_left'] - 0.3
    add_arrow(xa, band_y, xb, band_y, COLORS['link'], width=2.2)
    mx = (xa + xb) / 2
    add_rrect(mx - 4.5, band_y - 3.5, 9, 3.2, COLORS['white'], COLORS['link'], line_w=1.0, corner=0.2)
    add_textbox(mx - 4.5, band_y - 3.5, 9, 1.8, labels_top[i][0],
                size=10, bold=True, color=COLORS['link'])
    add_textbox(mx - 4.5, band_y - 1.9, 9, 1.5, labels_top[i][1],
                size=9, color=COLORS['link'])

# 反哺闭环（底部弧形 Ⅲ → Ⅰ）
loop_y = rc_y + rc_h - 2.5
# 用三段直线拟合弧形：Ⅲ底部向下 → 水平 → Ⅰ底部向上
c3 = rc_boxes[2]['cx']
c1 = rc_boxes[0]['cx']
bot_line_y = rc_y + rc_h + 4.0
add_arrow(c3, rc_y + rc_h - 0.5, c3, bot_line_y, COLORS['loop'], width=2.0)
add_arrow(c3 + 0.3, bot_line_y, c1 - 0.3, bot_line_y, COLORS['loop'], width=2.0)
add_arrow(c1, bot_line_y, c1, rc_y + rc_h - 0.5, COLORS['loop'], width=2.0)

# 闭环标签
loop_label_x = 50 - 12
add_rrect(loop_label_x, bot_line_y - 1.5, 24, 3.2, COLORS['white'], COLORS['loop'],
          line_w=1.2, corner=0.2)
add_textbox(loop_label_x, bot_line_y - 1.5, 24, 1.8,
            '↻ 评价反哺 · 学习分析',
            size=11, bold=True, color=COLORS['loop'])
add_textbox(loop_label_x, bot_line_y + 0.0, 24, 1.4,
            '驱动问题图谱动态演化 · 形成"诊断—伴学—评价—反哺"闭环',
            size=8.5, color=COLORS['loop'])

# 底部 底座/引擎/闭环 徽章
pill_y = bot_line_y + 3.8
for i, (label, color) in enumerate([('底座', COLORS['rq1_e']),
                                      ('引擎', COLORS['rq2_e']),
                                      ('闭环', COLORS['rq3_e'])]):
    add_rrect(rc_boxes[i]['cx'] - 2.2, pill_y, 4.4, 2.8,
              COLORS['white'], color, line_w=1.2, corner=0.3)
    add_textbox(rc_boxes[i]['cx'] - 2.2, pill_y, 4.4, 2.8,
                label, size=10.5, bold=True, color=color)

# =================================================================
# 第 4 层：成果层（y ≈ 84 - 96）
# =================================================================
layer_label(89, '成果层\n八位一体', COLORS['out_e'], h=10)

out_y, out_h = 84, 11
add_rrect(8, out_y, 85, out_h, RGBColor(0xFF,0xFD,0xF4), COLORS['out_e'], line_w=1.8)

out_items = [
    ('研究报告', '≥ 5 千字'),
    ('课程标准', '2026 版'),
    ('三图谱', '知识+能力+问题'),
    ('智学小助', '私有化智能体'),
    ('AI 合规官', '私有化智能体'),
    ('仿真靶场', '含红蓝对抗'),
    ('教学数据集', '脱敏样本'),
    ('数字资源包', '≥ 300 件'),
    ('教研论文', '≥ 1 篇'),
    ('迁移指南', '+ 师资培训'),
]
on_w = 7.4
on_gap = (85 - len(out_items)*on_w) / (len(out_items)-1)
on_x0 = 8
for i, (t, s) in enumerate(out_items):
    x = on_x0 + i*(on_w + on_gap)
    add_rrect(x + 0.15, out_y + 0.8, on_w - 0.3, out_h - 1.6,
              COLORS['out'], COLORS['out_e'], line_w=1.0, corner=0.12)
    add_textbox(x, out_y + 1.5, on_w, 3.5, t,
                size=10, bold=True, color=RGBColor(0x7D, 0x66, 0x08))
    add_textbox(x, out_y + 5.5, on_w, 3.5, s,
                size=9.5, color=COLORS['out_e'])

# 内容 → 成果
for b in rc_boxes:
    add_arrow(b['cx'], pill_y + 3.0, b['cx'], out_y, RGBColor(0x7D, 0x66, 0x08), width=1.4)

# ======== 保存 ========
out_path = '/workspace/technical_roadmap_simple_editable.pptx'
prs.save(out_path)
print('Saved:', out_path)
