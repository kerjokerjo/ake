# -*- coding: utf-8 -*-
"""生成教育部人文社会科学研究项目《申请评审书》（优化版）。
本版在修订稿基础上进一步优化：
  1) 国内外研究现状为具体学者与成果补充年份；
  2) 前期研究基础改为"表格+文字"呈现；
  3) 新增原生 Word 版"图 2-1 总体框架图"（表格+箭头，清晰可编辑）；
  4) 统一字体、行距、对齐与标题层级等版式；并同步 v2 的成员与小节改动。
所有流动文本置于三重单引号块，避免中英文引号与字符串定界符冲突。"""
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

TITLE = '巴渝优秀传统文化资源的知识图谱构建与传播创新研究'
CENTER = WD_ALIGN_PARAGRAPH.CENTER
RIGHT = WD_ALIGN_PARAGRAPH.RIGHT
JUST = WD_ALIGN_PARAGRAPH.JUSTIFY
CONTENT_W = 15.2  # 正文可用宽度（cm）

doc = Document()
normal = doc.styles['Normal']
normal.font.name = '仿宋_GB2312'
normal.font.size = Pt(12)
normal._element.rPr.rFonts.set(qn('w:eastAsia'), '仿宋_GB2312')

sec = doc.sections[0]
sec.top_margin = Cm(2.5)
sec.bottom_margin = Cm(2.5)
sec.left_margin = Cm(2.9)
sec.right_margin = Cm(2.9)


def set_cn_font(run, cn='仿宋_GB2312', size=12, bold=False, en='Times New Roman'):
    run.font.name = en
    run.font.size = Pt(size)
    run.font.bold = bold
    rpr = run._element.get_or_add_rPr()
    rf = rpr.get_or_add_rFonts()
    rf.set(qn('w:eastAsia'), cn)
    rf.set(qn('w:ascii'), en)
    rf.set(qn('w:hAnsi'), en)


def para(text='', size=12, bold=False, align=None, cn='仿宋_GB2312',
         first_indent=True, space_after=4, line=22):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(0)
    if line:
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(line)
    if first_indent and align not in (CENTER, RIGHT):
        pf.first_line_indent = Pt(size * 2)
    if text:
        r = p.add_run(text)
        set_cn_font(r, cn=cn, size=size, bold=bold)
    return p


def heading(text, size=12, cn='黑体', space_before=8, space_after=4):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    pf.line_spacing = Pt(22)
    r = p.add_run(text)
    set_cn_font(r, cn=cn, size=size, bold=True)
    return p


def set_cell(cell, text, size=10.5, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT, cn='仿宋_GB2312'):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(size + 5)
    for i, line in enumerate(text.split('\n')):
        if i > 0:
            p = cell.add_paragraph()
            p.alignment = align
            p.paragraph_format.space_after = Pt(1)
        r = p.add_run(line)
        set_cn_font(r, cn=cn, size=size, bold=bold)


def shade(cell, fill='D9D9D9'):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)


def cell_borders(cell, val='single', sz='6', color='808080'):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        e = OxmlElement('w:' + edge)
        e.set(qn('w:val'), val)
        e.set(qn('w:sz'), sz)
        e.set(qn('w:space'), '0')
        e.set(qn('w:color'), color)
        b.append(e)
    tcPr.append(b)


def no_borders(cell):
    cell_borders(cell, val='nil', sz='0', color='auto')


def render(block):
    """#S 黑体主标题；#K 楷体小标题；- 缩进两端对齐正文；~ 不缩进小字段。"""
    for raw in block.strip('\n').split('\n'):
        line = raw.strip()
        if not line:
            continue
        if line.startswith('#S '):
            heading(line[3:], size=12, cn='黑体')
        elif line.startswith('#K '):
            heading(line[3:], size=12, cn='楷体_GB2312', space_before=6, space_after=3)
        elif line.startswith('- '):
            para(line[2:], align=JUST)
        elif line.startswith('~ '):
            para(line[2:], size=10.5, first_indent=False, space_after=2, line=18)
        else:
            para(line, align=JUST)


# ============================================================
# 封面 A 表
# ============================================================
para(space_after=6, first_indent=False)
para('教育部人文社会科学研究项目', size=22, bold=True, align=CENTER, cn='宋体', first_indent=False, space_after=6, line=0)
para('申 请 评 审 书', size=26, bold=True, align=CENTER, cn='宋体', first_indent=False, space_after=26, line=0)

cover = [
    ('项 目 类 别：', '青年基金项目'),
    ('受 理 范 围：', '信息资源管理'),
    ('课 题 名 称：', TITLE),
    ('项 目 负 责 人：', '舒劲秋'),
    ('所 在 学 校：', '重庆工程职业技术学院（盖章）'),
    ('学 校 代 码：', '12759'),
    ('申 请 日 期：', '2026-6'),
]
t = doc.add_table(rows=len(cover), cols=2)
t.alignment = WD_TABLE_ALIGNMENT.CENTER
t.columns[0].width = Cm(4.5)
t.columns[1].width = Cm(9.5)
for i, (k, v) in enumerate(cover):
    set_cell(t.rows[i].cells[0], k, size=15, bold=True, align=RIGHT)
    set_cell(t.rows[i].cells[1], v, size=15, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT)
    t.rows[i].cells[0].width = Cm(4.5)
    t.rows[i].cells[1].width = Cm(9.5)
for cell in t._cells:
    no_borders(cell)

para(space_after=26, first_indent=False, line=0)
para('教 育 部 社 会 科 学 司 制', size=15, bold=True, align=CENTER, first_indent=False, line=0)
doc.add_page_break()

# ============================================================
# 1 申请人信息
# ============================================================
heading('1  申请人信息', size=14, cn='黑体', space_before=0)
info = doc.add_table(rows=6, cols=6)
info.style = 'Table Grid'
info.alignment = WD_TABLE_ALIGNMENT.CENTER


def kv(cl, cv, label, value):
    set_cell(cl, label, size=10.5, bold=True, align=CENTER)
    shade(cl)
    set_cell(cv, value, size=10.5, align=CENTER)


row = info.rows[0].cells
kv(row[0], row[1], '姓  名', '舒劲秋'); kv(row[2], row[3], '性  别', '女'); kv(row[4], row[5], '出生年月', '1989-10-7')
row = info.rows[1].cells
kv(row[0], row[1], '职  称', '讲师'); kv(row[2], row[3], '所在部门', '大数据与物联网学院'); kv(row[4], row[5], '最后学位', '硕士')
row = info.rows[2].cells
kv(row[0], row[1], '职  务', '无'); kv(row[2], row[3], '最后学历', '硕士研究生'); kv(row[4], row[5], '外语语种', '英语')
row = info.rows[3].cells
set_cell(row[0], 'E-Mail', size=10.5, bold=True, align=CENTER); shade(row[0])
row[1].merge(row[2]); set_cell(row[1], 'sjq@cqvie.edu.cn', size=10.5, align=CENTER)
set_cell(row[3], '手  机', size=10.5, bold=True, align=CENTER); shade(row[3])
row[4].merge(row[5]); set_cell(row[4], '18323187354', size=10.5, align=CENTER)
row = info.rows[4].cells
set_cell(row[0], '通讯地址', size=10.5, bold=True, align=CENTER); shade(row[0])
row[1].merge(row[2]).merge(row[3]); set_cell(row[1], '重庆市江津区圣泉街道南北大道南段 1111 号', size=10.5)
set_cell(row[4], '邮  编', size=10.5, bold=True, align=CENTER); shade(row[4])
set_cell(row[5], '402260', size=10.5, align=CENTER)
row = info.rows[5].cells
set_cell(row[0], '固定电话', size=10.5, bold=True, align=CENTER); shade(row[0])
row[1].merge(row[2]).merge(row[3]).merge(row[4]).merge(row[5]); set_cell(row[1], '')

para(space_after=2, first_indent=False, line=0)
render(r'''
~ 申请者本人近三年来主要研究成果（注明刊物的年、期或出版社、出版日期，不得加页）：
~ 1. 2022 年 10 月，主持重庆市教委教育教学改革重点课题：GZ222029，已结题，排名第三；
~ 2. 2023 年 6 月，主持重庆市教育委员会科学技术研究项目：KJQN202303428，已结题，排名第二；
~ 3. 2023 年 9 月，主持重庆市教育科学规划课题重点课题：K22YC314056，已结题，排名第三；
~ 4. 2024 年 1 月，获得“教育部供需对接就业育人项目”，排名第一；
~ 5. 2024 年 9 月，主持“‘智能定制·逐级赋能·智评促优’的数字技术人才培养模式构建与实践”获全国煤炭行业教学成果一等奖，排名第三；
~ 6. 2025 年 12 月，“产业导向·实境驱动·数智适配：高职工程技术类专业群实践教学体系改革与实践”获重庆市教学成果奖特等奖，排名第四；
~ 7. 2025 年 9 月，主持重庆市教育委员会科学技术研究项目：KJQN202503408，排名第一；
~ 8. 2025 年 11 月，发表 EI 会议（CA 检索）论文 1 篇，排名第一；
~ 9. 2026 年 5 月，主持高等教育科学研究课题专项课题：cqgj25zx49C，排名第一；
~ 10. 2026 年 5 月，主持全国高等职业院校“‘人工智能+’职业教育教学改革创新研究课题”，排名第一。
''')
doc.add_page_break()

# ============================================================
# 2 课题组主要成员情况
# ============================================================
heading('2  课题组主要成员情况及签名', size=14, cn='黑体', space_before=0)
members = [
    ['刘宇', '教授', '1980-5-31', '软件技术', '重庆工程职业技术学院', '负责知识图谱构建与技术实现工作。', ''],
    ['廖清科', '中级', '1990-2-22', '软件技术', '重庆工程职业技术学院', '负责文化资源整理、应用场景设计与传播服务支持工作。', ''],
    ['刘胜久', '中级', '1988-6-9', '人工智能', '重庆工程职业技术学院', '负责模型算法编写、数据处理等工作。', ''],
]
mt = doc.add_table(rows=1 + len(members), cols=7)
mt.style = 'Table Grid'
mt.alignment = WD_TABLE_ALIGNMENT.CENTER
for j, h in enumerate(['姓 名', '职称/职务', '出生日期', '专 业', '工作单位', '分工情况', '签 名']):
    set_cell(mt.rows[0].cells[j], h, size=10.5, bold=True, align=CENTER)
    shade(mt.rows[0].cells[j])
for i, m in enumerate(members, start=1):
    for j, v in enumerate(m):
        set_cell(mt.rows[i].cells[j], v, size=10, align=CENTER if j != 5 else WD_ALIGN_PARAGRAPH.LEFT)

para(space_after=2, first_indent=False, line=0)
render(r'''
~ 以上成员近三年来与本课题有关的主要研究成果，注明刊物的年、期或出版社、出版日期（不得加页）：
~ 1. 2022 年 10 月，主持重庆市教委教育教学改革重点课题“人工智能技术赋能顶岗实习课程思政创新研究”：GZ222029，已结题，排名第二；
~ 2. 2023 年 6 月，主持重庆工程职业技术学院校级教改课题“高职院校书香文化建设探索与实践”：JG222006S，已结题，排名第一；
~ 3. 2023 年 9 月，主持重庆市教育科学规划课题重点课题“职业院校文化育人创新实践研究”：K22YC314056，已结题，排名第三；
~ 4. 2025 年 11 月，发表 EI 会议（CA 检索）论文 1 篇“基于人工智能的数字职业教育及其效果评估——以顶岗实习课程为例”，排名第一。
''')
doc.add_page_break()

# ============================================================
# 3 B 表
# ============================================================
heading('3  B 表', size=14, cn='黑体', space_before=0)
para('（自此往下不得出现申请人个人身份信息，否则申请书作废！）', size=10.5, bold=True, first_indent=False, space_after=4, line=0)

bt = doc.add_table(rows=6, cols=4)
bt.style = 'Table Grid'
bt.alignment = WD_TABLE_ALIGNMENT.CENTER


def brow(i, k1, v1, k2=None, v2=None):
    c = bt.rows[i].cells
    set_cell(c[0], k1, size=10.5, bold=True, align=CENTER); shade(c[0])
    if k2 is None:
        c[1].merge(c[2]).merge(c[3]); set_cell(c[1], v1, size=10.5)
    else:
        set_cell(c[1], v1, size=10.5)
        set_cell(c[2], k2, size=10.5, bold=True, align=CENTER); shade(c[2])
        set_cell(c[3], v2, size=10.5)


brow(0, '课题名称', TITLE)
brow(1, '重点研究方向', '中华优秀传统文化传承创新研究', '受理范围', '信息资源管理')
brow(2, '学科及代码', '8507030/文化遗产学、5202020/自然语言处理')
brow(3, '研究类别', '基础研究', '计划完成时间', '2029-6')
brow(4, '最终成果形式', '论文')
brow(5, '申请经费总额(万元)', '8.00')

# ---------- 一、（含年份的研究现状） ----------
render(r'''
#S 一、本课题研究的理论和实际应用价值，目前国内外研究的现状和趋势（限 2 页，不能加页）
#K （一）研究背景与问题提出
- 巴渝优秀传统文化是中华优秀传统文化在长江上游地区的重要区域形态，涵盖古巴文化、三峡文化、码头文化、移民文化、非物质文化遗产、传统技艺与地方文献等，凝结着开放包容、坚韧尚义、敢闯敢为的精神品格。随着国家文化数字化战略与中华优秀传统文化传承发展工程的深入推进，其数字化保护、知识化组织与创新性传播已成为亟待研究的课题。
- 然而，当前巴渝优秀传统文化资源的数字化仍存在三个突出问题：一是资源分散、标准不一，相关资料散见于非遗名录、地方志、馆藏资料、文旅平台与新媒体，缺乏统一目录与数据规范；二是关联薄弱、阐释不深，现有成果多停留于图文展示与数据库检索，对文化资源间的人物、时空、技艺、文献等关系揭示不足；三是转化不足、场景不清，图谱、可视化等技术尚未有效服务于课程建设、研学设计、数字展陈与分众传播等现实需求。
- 基于此，本课题以巴渝优秀传统文化资源为对象，以知识图谱为核心技术路径，按照“资源数据化→数据知识化→知识价值化→价值传播化”的逻辑，依次开展资源整理、图谱构建、价值挖掘与传播创新研究，以回应上述问题。
#K （二）理论和实际应用价值
- 1. 理论价值。第一，拓展区域优秀传统文化研究的数字人文路径，推动其由经验描述向数据支撑与结构化解释转变；第二，构建面向区域传统文化资源的“资源类型—时空分布—人物主体—文献依据—文化价值—传播场景”知识组织框架，丰富文化资源知识组织与信息资源管理研究；第三，坚持以文化问题牵引技术应用，推动知识图谱技术与文化价值阐释深度融合；第四，以“资源数据化→数据知识化→知识价值化→价值传播化”为主线，构建传统文化活化传播的逻辑模型，为中华优秀传统文化创造性转化、创新性发展提供方法借鉴。
- 2. 实际应用价值。第一，服务巴渝优秀传统文化资源的系统整理，形成可持续更新的分类体系、数据目录与图谱模型；第二，服务中华优秀传统文化教育，形成主题学习路径、文化知识关联图与研学地图；第三，服务文旅融合与城市文化品牌建设，支持文化资源地图、非遗分布图、传承谱系与文化线路开发；第四，服务公共文化机构数字化转型与网络传播创新，为关联展示、智能检索、数字展陈、融媒体传播与文化 IP 开发提供知识支撑。
#K （三）国内外研究现状
- 1. 国外研究：起步较早，语义化、关联化基础成熟。在数字人文方面，2004 年 Schreibman、Siemens 与 Unsworth 主编的《A Companion to Digital Humanities》系统确立了数字技术在人文研究中的应用范式；2000 年 Moretti 提出“远读”（distant reading），主张以大规模文本分析发现文化的宏观结构；2007 年前后 Manovich 提出“文化分析学”（cultural analytics），关注大规模文化数据的计算分析。在文化遗产知识组织方面，1993 年 Gruber 关于本体（ontology）的研究奠定了概念建模基础，2001 年 Berners-Lee 等提出语义网思想；国际博物馆协会国际文献委员会（CIDOC）研制的 CIDOC CRM 于 2006 年成为国际标准（ISO 21127），成为文化遗产语义表达与互操作的核心参考模型；2005 年以来 Hyvönen 团队相继构建 MuseumFinland（2005）、CultureSampo（2009）等关联数据平台。近年来，知识图谱进一步应用于博物馆藏品管理、历史人物关系挖掘与智能导览等领域。
- 2. 国内研究：发展迅速，成果集中于三条主线。其一，数字人文与文化遗产数字化：2009 年前后武汉大学王晓光团队较早引介数字人文理论，并于 2019 年前后围绕敦煌壁画等建设文化遗产知识库与知识组织；2015 年前后中国人民大学冯惠玲团队提出“数字记忆”理论并建设“北京记忆”等平台；2016 年前后上海图书馆刘炜、夏翠娟等开展关联数据与历史人文大数据研究，建成家谱、古籍等关联数据知识库。其二，古籍与专门领域知识图谱构建：2016 年前后北京大学欧阳剑开展古籍数字化与文本挖掘，2019 年前后南京农业大学包平团队建设方志物产知识库；2021 年前后刘石、李飞跃等推进古籍智能整理与知识关联；2020 年前后吉林大学邓君团队围绕口述历史、红色文化与非遗开展数字记忆与知识图谱研究。其三，巴渝文化研究：周勇主编《重庆通史》（2002）、蓝勇《西南历史文化地理》（2003）、管维良《巴文化》（2004）等在地方通史、区域历史地理与巴文化源流方面积累丰厚，为理解巴渝文化的历史脉络与精神内涵提供了坚实基础。
#K （四）研究述评：现有研究空白与本课题切入点
- 综观国内外研究，既有成果为本课题提供了理论视角、语义建模方法与文化史料基础，但针对巴渝优秀传统文化资源的知识图谱构建与传播创新，仍存在四个明显空白点，构成本课题的逻辑出发点：
- 一是资源整合的空白：重单类、轻综合。国外研究对象多集中于西方文化遗产与馆藏数据，国内成果亦多聚焦古籍、非遗、红色文化等单一门类，缺少面向巴渝多类型、跨载体文化资源的系统化整理与统一知识组织。本课题拟建立“资源类型—时空信息—文化主题—数据来源—传播载体”分类体系，填补综合性区域文化资源整合的空白。
- 二是知识建模的空白：重技术、轻文化。现有知识图谱研究多偏重实体识别、关系抽取等技术流程，对文化资源背后的历史脉络与价值谱系揭示不足，且巴渝文化尚缺专门的本体模型。本课题拟以文化问题牵引技术，构建适配巴渝文化语义特征的本体与图谱原型，填补区域文化专门知识建模的空白。
- 三是价值阐释的空白：重展示、轻解释。已有巴渝文化研究多为专题、个案的定性描述，缺少数据化、结构化、可计算的价值挖掘。本课题拟依托图谱从时空、主题与关系网络维度提炼巴渝文化的价值内涵与精神标识，填补由知识组织走向价值解释的空白。
- 四是传播转化的空白：重构建、轻应用。现有研究普遍止步于图谱构建，缺少“构建—阐释—传播”贯通的机制设计。本课题拟提出知识图谱驱动的分场景传播创新路径，填补研究成果向现实应用转化的空白。
- 综上，本课题以“资源数据化→数据知识化→知识价值化→价值传播化”为主线，构建资源分类体系、本体模型与知识图谱原型，揭示其结构特征与价值内涵，并提出面向现实场景的传播创新路径，形成完整而有逻辑起点的研究论证。
''')

# ---------- 二、 ----------
render(r'''
#S 二、本课题的研究目标、研究内容、拟突破的重点和难点（限 2 页，不能加页）
#K （一）研究目标
- 总体目标：以数字人文为视角，以巴渝优秀传统文化资源为研究对象，围绕资源分散、关联薄弱、价值挖掘不足和传播转化不足等问题，构建其知识组织框架与知识图谱原型，揭示内在结构、历史脉络、文化价值与传播规律，提出知识图谱驱动的传播创新路径。具体目标按“资源数据化→数据知识化→知识价值化→价值传播化”递进：一是构建资源分类体系与数据规范，形成基础数据目录；二是设计实体、属性与关系类型，形成知识图谱原型；三是从时空、主题与关系网络维度分析，提炼价值谱系与精神特质；四是面向教育、文旅、公共文化服务与网络传播场景提出可操作的传播策略与机制建议。
#K （二）研究内容
- 1. 巴渝优秀传统文化资源数据化整理研究。系统梳理历史文化、非遗文化、民俗文化、地方文献、传统技艺、三峡文化、码头文化、移民文化、饮食文化等资源，明确对象边界与类型，分析其数据来源、载体形态、空间分布与保存现状，建立“资源类型—时空信息—文化主题—数据来源—传播载体”分类体系。
- 2. 巴渝优秀传统文化资源知识图谱构建研究。设置“文化资源、人物、地点、时间、事件、文献、技艺、机构、符号、传播载体”等核心实体，“位于、起源于、属于、记载于、传承于、发生于、影响、融合、传播于、关联于”等关系，以及名称、类别、时间、地点、来源、代表人物、传承方式、文化价值、传播渠道等属性；按“数据采集→数据清洗→实体识别→关系抽取→知识融合→人工校核→图数据库存储→可视化呈现”流程构建图谱原型。
- 3. 基于知识图谱的巴渝文化价值挖掘研究。空间维度识别区域分布与集聚特征；时间维度梳理古巴文化、三峡文化、移民文化、码头文化与当代城市文化的演进脉络；主题维度提炼山水、江河、移民、码头、工匠、民俗、开放、坚韧等核心主题；关系维度分析人物、技艺、文献、事件、机构与传播平台的网络结构，进而提炼价值内涵与精神标识。
- 4. 知识图谱驱动的巴渝文化传播创新研究。教育场景设计课程资源图谱、主题学习路径、研学地图与知识问答；文旅场景设计文化资源地图、非遗主题线路、古镇文化导览与三峡文化路线；公共文化服务场景探索博物馆、文化馆、非遗馆的关联展示与数字展陈；网络传播场景提出基于图谱的内容选题、故事化表达、分众传播与文化 IP 开发策略。
#K （三）拟突破的重点
- 第一，突破资源分散化问题，形成系统化数据整理框架；第二，突破文化资源弱关联问题，形成区域传统文化知识图谱模型；第三，突破文化阐释浅表化问题，形成基于图谱的价值挖掘方法；第四，突破研究成果应用不足问题，形成场景化传播机制。
#K （四）拟解决的难点
- 第一，多源异构数据整合难，拟通过统一字段设计、数据清洗规范、来源标注与人工复核提高数据一致性；第二，文化语义抽取准确难，拟采用“规则词表+自然语言处理+人工校核”方式提高实体识别与关系抽取准确率；第三，技术建模与文化阐释结合难，坚持以文化问题牵引技术应用；第四，成果应用转化难，从教育、文旅、公共文化服务与网络传播四类场景切入，设计可操作方案与示范案例。
''')

# ---------- 图 2-1 总体框架图（原生表格+箭头） ----------
def frame_box(label, text, fill='DEEAF6', label_font='黑体', text_font='仿宋_GB2312', size=10.5):
    tb = doc.add_table(rows=1, cols=1)
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    c = tb.rows[0].cells[0]
    c.width = Cm(CONTENT_W)
    cell_borders(c, color='4472C4', sz='8')
    shade(c, fill)
    c.text = ''
    p = c.paragraphs[0]
    p.alignment = CENTER
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(size + 4)
    if label:
        r = p.add_run(label)
        set_cn_font(r, cn=label_font, size=size, bold=True)
    r = p.add_run(text)
    set_cn_font(r, cn=text_font, size=size, bold=False)
    return tb


def frame_arrow():
    p = doc.add_paragraph()
    p.alignment = CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(12)
    r = p.add_run('▼')
    set_cn_font(r, cn='宋体', size=9, bold=True)


para(space_after=2, first_indent=False, line=0)
frame_box('研究目标：', '构建巴渝优秀传统文化资源知识图谱，揭示价值内涵，提出传播创新路径', fill='FCE4D6')
frame_arrow()
frame_box('现实问题：', '资源分散　·　关联薄弱　·　阐释不深　·　转化不足', fill='FFF2CC')
frame_arrow()
frame_box('① 资源数据化　', '资源普查 → 分类体系 → 数据规范 → 基础数据目录')
frame_arrow()
frame_box('② 数据知识化　', '本体建模 → 实体识别 → 关系抽取 → 知识融合 → 图谱原型')
frame_arrow()
frame_box('③ 知识价值化　', '时空分析 → 主题分析 → 关系网络分析 → 价值内涵与精神标识')
frame_arrow()
frame_box('④ 价值传播化　', '文化教育　·　文旅融合　·　公共文化服务　·　网络传播')
frame_arrow()
frame_box('预期成果：', '学术论文　·　研究报告　·　知识图谱原型　·　可视化成果', fill='FCE4D6')
frame_box('理论与技术支撑：', '数字人文理论　·　信息资源管理　·　知识图谱技术　·　巴渝文化研究', fill='E2EFDA')
para('图 2-1  本课题总体框架图', size=10.5, bold=True, align=CENTER, first_indent=False, cn='楷体_GB2312', space_after=4, line=16)

# ---------- 三、 ----------
render(r'''
#S 三、本课题的研究思路和研究方法、计划进度、前期研究基础及资料准备情况（限 2 页，不能加页）
#K （一）研究思路
- 本课题遵循“资源数据化→数据知识化→知识价值化→价值传播化”的总体思路：首先，通过文献梳理、数据采集与内容分析明确资源范围、类型、来源与特征，形成基础数据目录与分类体系；其次，构建本体模型并通过实体识别、关系抽取、知识融合与图数据库存储形成图谱原型；再次，从时间、空间、主题与关系维度揭示结构特征、演变脉络与价值内涵，提炼精神标识；最后，将图谱分析结果转化为教育、文旅、公共文化服务与网络传播中的应用方案。
#K （二）研究方法
- 1. 文献研究与内容分析法。系统梳理数字人文、知识图谱、文化遗产数字化、信息资源管理与巴渝文化研究成果，对政策文件、非遗名录、地方志、文旅文本、博物馆资料与网络文章进行内容编码。
- 2. 知识图谱构建方法。采用本体建模、实体识别、关系抽取、知识融合、图数据库存储与可视化表达等方法，技术流程涵盖数据采集、清洗、标注、抽取、融合、校验、存储与展示。
- 3. 图谱分析与可视化方法。运用关系网络分析、时空可视化、主题聚类与传承谱系分析等方法，揭示空间分布、时间演化、主题结构与关系网络。
- 4. 案例研究与调研访谈法。选取川江号子、铜梁龙舞、梁平木版年画、荣昌夏布、秀山花灯、大足石刻相关文化、钓鱼城历史文化、磁器口古镇等代表性案例进行验证，并调研相关主体的数字传播需求。
#K （三）计划进度
''')

prog = doc.add_table(rows=5, cols=5)
prog.style = 'Table Grid'
prog.alignment = WD_TABLE_ALIGNMENT.CENTER
for j, h in enumerate(['阶段', '时间安排', '阶段主题', '主要研究任务', '阶段成果']):
    set_cell(prog.rows[0].cells[j], h, size=9.5, bold=True, align=CENTER); shade(prog.rows[0].cells[j])
prog_rows = [
    ['第一阶段', '2026.7—2026.12', '资源数据化', '完成国内外文献综述，明确研究对象边界，制定数据采集规范；收集非遗名录、地方志、文旅资料、学术文献与网络公开资料，形成资源初步目录。', '文献综述；资源采集方案；资源初步目录。'],
    ['第二阶段', '2027.1—2027.12', '数据知识化', '完成资源分类体系设计，构建知识图谱本体模型，明确核心实体、属性与关系类型；开展数据清洗、实体识别与关系抽取，形成结构化数据集。', '资源分类体系；本体模型；结构化数据集；阶段论文。'],
    ['第三阶段', '2028.1—2028.12', '知识价值化', '完成知识融合、人工校核、图数据库存储与图谱可视化，形成知识图谱原型；开展空间、时间、主题与关系网络分析，提炼文化价值内涵。', '知识图谱原型；可视化图表；阶段研究报告；阶段论文。'],
    ['第四阶段', '2029.1—2029.6', '价值传播化', '提出面向教育、文旅、公共文化服务与网络传播的传播创新路径，完成最终研究报告、学术论文与结项材料。', '最终研究报告；咨询报告；论文；结项材料。'],
]
for i, rd in enumerate(prog_rows, start=1):
    for j, v in enumerate(rd):
        set_cell(prog.rows[i].cells[j], v, size=9, align=CENTER if j in (0, 1, 2) else WD_ALIGN_PARAGRAPH.LEFT)

# 前期研究基础：文字导语 + 表格 + 衔接
render(r'''
#K （四）前期研究基础及资料准备情况
- 1. 前期研究基础。课题负责人及团队围绕数字技术应用、人工智能、大数据处理与文化育人等领域已开展系统研究，主持和参与省部级及以上科研与教改课题 7 项，形成扎实积累。核心前期成果及其对本课题的支撑作用如下表所示。
''')

pre = doc.add_table(rows=8, cols=4)
pre.style = 'Table Grid'
pre.alignment = WD_TABLE_ALIGNMENT.CENTER
for j, h in enumerate(['前期项目 / 成果（编号）', '类别 / 时间', '本人角色', '对本课题的支撑作用']):
    set_cell(pre.rows[0].cells[j], h, size=9.5, bold=True, align=CENTER); shade(pre.rows[0].cells[j])
pre_rows = [
    ['高职院校书香文化建设探索与实践（JG222006S）', '校级教改·已结题', '主持', '积累传统文化校园传播的受众反馈与实践经验，支撑传播路径设计'],
    ['职业院校文化育人创新实践研究（K22YC314056）', '市教科规划重点·已结题', '参与', '梳理巴渝区域文化资源的教育价值，支撑资源范围界定与类型划分'],
    ['重庆市教委科学技术研究项目（KJQN202303428、KJQN202503408）', '市级科研', '主持', '积累人工智能与大数据技术，支撑图谱构建与智能分析'],
    ['高等教育科学研究专项重点课题（cqgj25zx49C）', '省级专项重点', '主持', 'AI 赋能教育改革，支撑教育应用与传播场景设计'],
    ['全国高职“人工智能+职业教育教学改革”创新研究课题', '国家级行指委课题', '主持', '探索智能技术驱动文化教育的实施路径'],
    ['基于大数据技术的高速公路营运数据分析与应用等校级科研', '校级科研·已结题', '主持', '积累数据采集清洗、建模分析与可视化能力'],
    ['课程思政创新研究（GZ222029）及 EI 会议论文（AI+职业教育）', '市级重点 / 学术论文', '参与 / 主持', '形成课程思政与文化育人的跨学科融合视角'],
]
for i, rd in enumerate(pre_rows, start=1):
    for j, v in enumerate(rd):
        set_cell(pre.rows[i].cells[j], v, size=9, align=CENTER if j in (1, 2) else WD_ALIGN_PARAGRAPH.LEFT)

render(r'''
- 前期研究与本课题在研究对象、技术路线与应用场景三个维度上层层递进、相互支撑：文化育人类课题奠定资源与场景基础，人工智能与大数据类课题提供本体建模、实体识别、关系抽取与可视化的方法与工具储备，共同构成体系化的研究布局。
- 2. 资料准备情况。课题组已在文献、数据与技术工具三方面完成前期储备，具体如下表所示。
''')

mat = doc.add_table(rows=4, cols=2)
mat.style = 'Table Grid'
mat.alignment = WD_TABLE_ALIGNMENT.CENTER
for j, h in enumerate(['类别', '内容与规模']):
    set_cell(mat.rows[0].cells[j], h, size=9.5, bold=True, align=CENTER); shade(mat.rows[0].cells[j])
mat_rows = [
    ['文献资料', '数字人文与文化遗产数字化中外文献 200 余篇（含敦煌学、非遗保护等应用案例）；知识图谱构建技术文献及开源工具文档；《巴渝文化概论》《重庆通史》等专著及专题论文 50 余篇。'],
    ['数据资源', '重庆中国三峡博物馆、巴渝民俗博物馆等线上展览与藏品目录；知网、万方等库中巴渝文化研究论文元数据；百度百科、微信公众号、抖音等新媒体传播内容；已初步设计采集规范与元数据方案。'],
    ['技术工具与环境', '大数据实训中心与人工智能创新实验室，配备高性能服务器与 GPU 工作站；已部署 Scrapy+Pandas 采集、Neo4j+Protégé+DeepKE 图谱构建、HanLP+BERT 自然语言处理、ECharts+Gephi 可视化工具链，均经前期项目验证。'],
]
for i, rd in enumerate(mat_rows, start=1):
    set_cell(mat.rows[i].cells[0], rd[0], size=9.5, bold=True, align=CENTER)
    set_cell(mat.rows[i].cells[1], rd[1], size=9)

# ---------- 四、 ----------
render(r'''
#S 四、本课题研究的中期成果、最终成果，研究成果的预计去向（不得加页）
#K （一）中期成果
- 1. 巴渝优秀传统文化资源分类体系与数据目录；2. 资源知识图谱本体模型；3. 结构化数据集与知识图谱初步原型；4. 阶段性论文、研究报告。
#K （二）最终成果
- 1. 学术论文：拟发表学术论文 1 篇，主题涵盖资源知识图谱构建、基于知识图谱的文化价值挖掘、知识图谱驱动的传统文化传播创新等；2. 研究报告：形成最终研究报告 1 份，系统呈现资源分类体系、图谱构建过程、价值挖掘结果与传播创新路径；3. 知识图谱原型及可视化成果：形成资源知识图谱原型，以及文化资源地图、非遗分布图、传承谱系图、主题关系网络图等可视化成果。
#K （三）成果预计去向
- 1. 服务学术研究，为数字人文、信息资源管理、文化遗产数字化、知识组织与区域文化研究提供案例、数据与方法参考；2. 服务文化教育，转化为地方传统文化课程资源、研学资源、主题学习路径与文化知识图谱；3. 服务文旅融合，为文化线路设计、非遗体验、城市文化品牌建设、古镇导览与三峡文化传播提供支撑；4. 服务公共文化和网络传播，为博物馆、文化馆、非遗保护机构开展数字展陈、知识检索、关联展示与融媒体传播提供参考。
#S 五、经费预算（单位：万元）
''')

bud = doc.add_table(rows=6, cols=3)
bud.style = 'Table Grid'
bud.alignment = WD_TABLE_ALIGNMENT.CENTER
for j, h in enumerate(['类别', '金额（万元）', '说明']):
    set_cell(bud.rows[0].cells[j], h, size=10.5, bold=True, align=CENTER); shade(bud.rows[0].cells[j])
bud_rows = [
    ['直接费用合计', '7.00', ''],
    ['　业务费', '3.00', '资料费、调研差旅、数据采集与会议交流等'],
    ['　劳务费', '2.00', '数据标注、图谱校核与咨询劳务'],
    ['　设备费', '2.00', '计算与存储资源、软件工具等'],
    ['间接经费', '1.00', '其中外拨经费 0.00'],
]
for i, rd in enumerate(bud_rows, start=1):
    for j, v in enumerate(rd):
        set_cell(bud.rows[i].cells[j], v, size=10, align=WD_ALIGN_PARAGRAPH.LEFT if j == 2 else CENTER)
para(space_after=2, first_indent=False, line=0)
para('申请经费年度预算（不含其他来源经费）：2026 年 3.00 万元；2027 年 3.00 万元；2028 年 2.00 万元。', size=10.5, first_indent=False, line=18)
doc.add_page_break()

# ============================================================
# 承诺书
# ============================================================
heading('项目承诺书', size=16, cn='宋体', space_before=0, space_after=6)
para('本人保证项目申请书填报内容真实，不存在任何知识产权问题。若获准立项，本人将严格按照本表填写内容，按时完成研究计划，按要求及时报送中检、终结等相关材料。遵守教育部关于科研项目管理的各项规定，如有违反，本人将承担相关责任。', align=JUST)
para(space_after=18, first_indent=False, line=0)
para('申请者（签章）：', first_indent=False, align=RIGHT, line=0)
para('年    月    日', first_indent=False, align=RIGHT, space_after=16, line=0)
heading('学校科研管理部门意见', size=12, cn='黑体')
para('经审核，申请人符合资格，申请书内容属实，同意上报。若批准立项，学校保证上交与电子版内容一致、签章齐全的纸质申请书以供存档，保证为本项目的科学研究提供必要条件，并严格按照教育部关于科研项目管理的各项规定对项目实施进行管理。', align=JUST)
para('科研管理部门公章', first_indent=False, align=RIGHT, line=0)
para('年    月    日', first_indent=False, align=RIGHT, space_after=12, line=0)
heading('省市自治区教育部门、其他部委教育司社科研究管理部门意见', size=12, cn='黑体')
para('同意申报。', align=JUST)
para('单 位 公 章', first_indent=False, align=RIGHT, line=0)
para('年    月    日', first_indent=False, align=RIGHT, line=0)

OUT = '巴渝优秀传统文化资源知识图谱构建与传播创新研究-申请评审书（优化版）.docx'
doc.save(OUT)
print('saved:', OUT)
