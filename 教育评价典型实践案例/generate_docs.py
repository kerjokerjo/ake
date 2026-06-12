# -*- coding: utf-8 -*-
"""生成两份提交材料（“三智能体”AI赋能版）：
1. 师范生教育教学能力认定抽检评价改革实践-案例文稿.docx
2. 附件1 教育评价领域典型实践案例征集表.docx
图1~图4 以占位行标出，由用户自行插入图片。
"""
import os
import re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = os.path.dirname(os.path.abspath(__file__))

CASE_TITLE_MAIN = "三智能体协同　智赋评价"
CASE_TITLE_SUB = "——师范生教育教学能力认定抽检的AI赋能全流程质量监测实践"
CASE_NAME_ONELINE = "三智能体协同 智赋评价——师范生教育教学能力认定抽检的AI赋能全流程质量监测实践"


def set_font(run, cn_font="仿宋", en_font="Times New Roman", size=14, bold=False, color=None):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = en_font
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:eastAsia'), cn_font)
    rfonts.set(qn('w:ascii'), en_font)
    rfonts.set(qn('w:hAnsi'), en_font)
    if color is not None:
        run.font.color.rgb = color


# ============================================================
# 文档一：案例文稿
# ============================================================

def build_case_doc():
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(2.5)
    sec.left_margin = Cm(3.0)
    sec.right_margin = Cm(2.6)

    def title():
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(CASE_TITLE_MAIN)
        set_font(r, cn_font="方正小标宋简体", size=20, bold=True)
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after = Pt(16)
        r2 = p2.add_run(CASE_TITLE_SUB)
        set_font(r2, cn_font="黑体", size=15, bold=False)

    def h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.5
        r = p.add_run(text)
        set_font(r, cn_font="黑体", size=15, bold=True)

    def h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.5
        r = p.add_run(text)
        set_font(r, cn_font="黑体", size=14, bold=True)

    def h3(text):
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Pt(28)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.line_spacing = 1.5
        r = p.add_run(text)
        set_font(r, cn_font="仿宋", size=14, bold=True)

    def body(text):
        # 使用 ** 包裹的片段加粗
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Pt(28)
        p.paragraph_format.line_spacing = 1.5
        for i, seg in enumerate(text.split("**")):
            if not seg:
                continue
            r = p.add_run(seg)
            set_font(r, cn_font="仿宋", size=14, bold=(i % 2 == 1))

    def figure(caption):
        ph = doc.add_paragraph()
        ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ph.paragraph_format.space_before = Pt(6)
        ph.paragraph_format.space_after = Pt(2)
        rr = ph.add_run("［请在此处插入图片］")
        set_font(rr, cn_font="仿宋", size=12, color=RGBColor(0x99, 0x99, 0x99))
        pc = doc.add_paragraph()
        pc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pc.paragraph_format.space_after = Pt(8)
        rc = pc.add_run(caption)
        set_font(rc, cn_font="楷体", size=12, bold=True)

    title()

    # ---------- 一、案例背景 ----------
    h1("一、案例背景")
    body("师范生教育教学能力认定是教师教育质量保障的重要环节，直接关系师范生培养质量监测的科学性、公平性和公信力。"
         "本系统始终立足“学生成长”，旨在借助智能化技术手段，**对师范生的教学设计、课堂实施、育人能力等核心素养进行全过程动态监测**，"
         "**实现从实习实践到毕业考核的全流程数据采集与智能分析**，把师范生能力成长贯穿评价始终。"
         "按照《重庆市2025年师范生教育教学能力认定抽检实施方案（试行）》要求，抽检需覆盖样本抽取、材料上传、校级审核、"
         "市教育评估院审核、专家评分、结果发布等环节，涉及评估院、高校、学生、专家等多类主体，**业务链条长、协同难度大、规范要求高**。")
    body("以往工作主要依赖线下通知、人工汇总和经验判断，存在三方面突出问题：一是**标准执行不统一**，指标体系、评分规则、举证要求复杂，"
         "缺乏数字化承载，各校理解执行易产生偏差；二是**审核评审效率低**，传统系统仅校验材料数量、格式、大小，内容是否合规仍靠人工逐份查看，"
         "专家需阅读大量PDF和视频并撰写意见，单生评阅耗时长，且易意见趋同化、模板化；三是**结果运用支撑弱**，评审结果多以分数明细呈现，"
         "缺乏结构化诊断，学校难获改进信息，评价的诊断与反哺功能未能释放。基于此，**本单位联合重庆市教育评估院**，对接重庆市师范生教育教学能力认定抽检实施要求，"
         "提出**“把政策要求转化为平台规则、把人工核查升级为智能预检、把专家评审升级为人机协同、把结果汇总升级为智能诊断”**的思路，"
         "联合研发师范生教育教学能力认定抽检系统并**创新嵌入三大AI智能体**，以**数字化与智能化双轮驱动**，"
         "推动抽检从经验驱动走向规则驱动、从人工核查走向人机协同、从结果管理走向智能诊断。目前系统已**支撑全市10余所师范类高校、30余个师范类专业、500余名抽检师范生完成评审任务**，"
         "为教育评价数字化与智能化转型提供了**可复制、可推广的重庆实践样本**。")
    figure("图1　系统总体架构与三智能体业务流程图")

    # ---------- 二、具体举措及实施成效 ----------
    h1("二、具体举措及实施成效")
    h2("（一）构建“三智能体”AI赋能体系，破解评价全链条共性难题")
    body("本案例最具突破性的创新，在于围绕抽检业务全链条**精准嵌入三个差异化AI智能体**，分别赋能学生与学校审核、专家评审、"
         "结果反馈三大关键环节，形成**“贯通上下游、人机各司其职”**的智能化质量监测体系。")

    h3("1.佐证材料“智能预检”，赋能学生上传与学校审核")
    body("在学生上传与校级审核环节嵌入**智能预检智能体**，构建**“双场景预检”机制**：学生提交前可一键启动预检，"
         "智能体基于该指标的评价标准与举证要求，自动检测内容相关性、要素完整性、表述规范性、内容深度及多文档一致性等维度，"
         "生成可视化报告供学生迭代优化后再提交；校管审核时系统自动呈现预检结果，重点问题以红、黄、绿三色标识并定位到具体页面段落，"
         "校管在AI辅助下高效审阅，**“通过、驳回”决定权仍完全保留在校管手中，AI只作“审阅助手”，不替代专业判断**。")
    figure("图2　佐证材料“智能预检”反馈界面截图")

    h3("2.专家评审“AI辅助”，赋能专家评阅与意见撰写")
    body("在专家评审环节嵌入**评审辅助智能体**，提供**“评审参考意见智能生成”服务**：专家打开评审页面时，"
         "智能体基于学生完整举证材料、认定标准与评分规则、各二级指标考察要点，自动生成针对每条二级指标的结构化初步意见，"
         "含达成度初判、关键证据定位、优势点、不足点与建议参考。AI意见明确标注为“参考建议”，专家可采纳、修改或自主撰写，"
         "**评分与最终意见均由专家独立确认，AI不参与最终评分**，系统留存AI与专家内容差异，确保评审可追溯。")
    figure("图3　专家评审界面（含AI辅助意见生成）")

    h3("3.评审结果“智能诊断报告”，赋能学校反馈与管理决策")
    body("构建**认定结果分析智能体**，在工具库评审基本完成后自动激活，基于全量数据生成涵盖总体概览、能力维度画像、指标达成分析、"
         "典型问题归纳、横向对比、改进建议等板块的**“诊断书+处方”式报告**。学校**收到的不再是单一分数**，"
         "而是可直接对接培养方案优化、实践教学改进、薄弱指标强化的诊断建议；评估院则可汇总各校共性发现，"
         "为全市师范生培养质量年度分析提供数据支撑。")
    figure("图4　评审结果学校维度、管理端“智能诊断报告”展示界面截图")

    h2("（二）推动标准数字化，夯实智能运行底座")
    body("一是将抽检标准转化为**可配置指标表**，支持分层配置指标与评分规则，把文本标准变为可执行规则，为三个智能体提供核心规则依据；"
         "二是将评分办法转化为**可计算规则**，支持等级制与分数制，等级可设对应分值，兼顾专业判断与定量分析；"
         "三是将举证要求转化为**材料清单**，建立举证材料库（支持PDF、MP4，可设数量、大小、上传主体、排序），"
         "通过指标与材料多对多关联，明确“用什么材料证明什么能力”，为智能预检提供检测靶向。")

    h2("（三）推动流程闭环化，破解多方协同难题")
    body("一是构建**多角色分权协同机制**，设评估院管理员、学校管理员、学生、专家等角色，各司其职、边界清晰，AI在各环节差异化辅助；"
         "二是建立**“学生上传—校级审核—评估院复核—专家评分”分级审核链条**，“谁提交、谁审核、谁把关”清晰嵌入流程；"
         "三是实现**在线查证与全过程留痕**，专家在线查看关联佐证材料并评分，系统记录登录、操作、AI使用及流程节点日志，"
         "确保可查询、可追溯、可复盘。")

    h2("（四）推动规则公平化，统一抽样与评审")
    body("一是创新样本抽取规则，按专业人数动态设抽取比例，按同校同专业**“按序等距抽取”**，兼顾随机性、代表性与公平，"
         "抽取后可手动微调并确认名单；二是创新专家配置机制，支持**一生多专家**、同一专家按学校或指标范围承担任务，"
         "避免任务失衡、提升组织灵活性；三是创新结果查看与反馈，实时查看评审进度、导出得分明细，"
         "并基于智能诊断报告向学校反馈，使结果由结论层面升级为改进依据。")

    h2("（五）推动成果可复制，放大改革示范效应")
    body("平台采用**“指标表+工具库+角色权限+流程配置+智能体”的模块化设计**，三个智能体可基于不同项目的指标体系与材料类型快速适配，"
         "不依赖单一学校、专业或年度任务；改革重构了标准承载、流程组织、人机协同与结果应用方式，形成**“数字化+智能化”双轮驱动的完整路径**；"
         "其**“预检前置、辅助评审、智能诊断”理念**还可迁移至专业认证、教育实习评价、课程达成度监测、学生发展评价等场景，具有较强推广价值。")

    h2("（六）实施成效")
    body("系统上线以来在效率、质量、覆盖、应用等维度成效显著。一是智能预检大幅提升审核效率与材料质量：**单份材料审核时长由约20分钟降至约10分钟"
         "（降约50%）**，全市**10余所高校累计完成500余份材料审核**，500余名抽检师范生中**超90%材料一次性通过校审、一次性通过率较改革前提升约30个百分点**，"
         "校审环节即识别约100余处多文档数据不一致。二是AI评审辅助显著减负提质：**单生平均评审时长由约30分钟降至约17分钟（降约43%）**，"
         "评审意见针对性、专业性明显增强。三是智能诊断推动结果运用由“分数通报”转向“质量反哺”：**单份学校诊断报告生成由约1天降至约90分钟**，"
         "参检学校普遍反映便于培养方案优化与薄弱指标强化，汇总形成的全市培养质量分析报告为教育主管部门决策提供有力支撑。")

    # ---------- 三、经验总结 ----------
    h1("三、经验总结")
    body("一是坚持**AI辅助、人主决策**。教育评价的专业性、敏感性、责任性决定AI不能替代人的判断，三个智能体始终坚守"
         "**“AI识别问题、提供参考、辅助决策，人作最终判断、承担责任”**的边界，既释放效率价值又守住专业底线。")
    body("二是坚持**精准嵌入、靶向赋能**。AI赋能不“撒胡椒面”，三个智能体分别针对审核效率、评审质量、结果运用三个环节的核心矛盾精准切入，"
         "避免“为AI而AI”的泛化应用。")
    body("三是坚持**标准先行、数据贯通**。智能体能力建立在结构化指标、规范化数据与闭环化流程之上，"
         "必须先把指标配置、举证管理、数据沉淀等基础打牢，智能化才能真正落地。")
    body("四是坚持**闭环设计、持续改进**。通过智能诊断把评价结果转化为改进建议，推动学校**“评价—诊断—改进—再评价”良性循环**，"
         "让教育评价真正发挥质量提升功能。")
    body("总体来看，本系统以**“三智能体”AI赋能体系**为核心创新，破解了审核效率低、评审负担重、结果运用浅等共性问题，"
         "构建**“标准入库→AI预检→规范审核→智能辅助→自动诊断→数据反馈”全流程闭环质量监测模式**，"
         "形成可复制、可推广、可持续迭代的**“AI+教育评价”实践范式**，对推进教育评价现代化、人工智能赋能教师教育、"
         "教育数字化转型走深走实具有较强参考与示范意义。")

    out = os.path.join(BASE, "师范生教育教学能力认定抽检评价改革实践-案例文稿.docx")
    doc.save(out)

    # 统计正文中文字数（不含图题占位）
    cn = 0
    for p in doc.paragraphs:
        if "请在此处插入图片" in p.text:
            continue
        cn += len(re.findall(r'[\u4e00-\u9fff]', p.text))
    return out, cn


# ============================================================
# 文档二：附件1 征集表
# ============================================================

def set_cell(cell, text, bold=False, size=12, align=WD_ALIGN_PARAGRAPH.LEFT,
             cn_font="仿宋", valign=WD_ALIGN_VERTICAL.CENTER):
    cell.vertical_alignment = valign
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.line_spacing = 1.3
    lines = text.split("\n")
    for i, line in enumerate(lines):
        if i > 0:
            p = cell.add_paragraph()
            p.alignment = align
            p.paragraph_format.line_spacing = 1.3
        run = p.add_run(line)
        set_font(run, cn_font=cn_font, size=size, bold=bold)


def shade_cell(cell, color="EAEFF7"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)


def build_form_doc():
    doc = Document()
    sec = doc.sections[0]
    sec.top_margin = Cm(2.2)
    sec.bottom_margin = Cm(2.2)
    sec.left_margin = Cm(2.5)
    sec.right_margin = Cm(2.5)

    pa = doc.add_paragraph()
    pa.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    ra = pa.add_run("附件1")
    set_font(ra, cn_font="黑体", size=14, bold=True)

    pt = doc.add_paragraph()
    pt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pt.paragraph_format.space_after = Pt(10)
    rt = pt.add_run("教育评价领域典型实践案例征集表")
    set_font(rt, cn_font="方正小标宋简体", size=18, bold=True)

    table = doc.add_table(rows=0, cols=6)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    def add_row():
        return table.add_row().cells

    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "案例类别", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    set_cell(c[2], "□ 学校发展", align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell(c[3], "□ 教师发展", align=WD_ALIGN_PARAGRAPH.CENTER)
    c[4].merge(c[5])
    set_cell(c[4], "☑ 学生成长", align=WD_ALIGN_PARAGRAPH.CENTER)

    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "案例名称", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    merged = c[2].merge(c[3]).merge(c[4]).merge(c[5])
    set_cell(merged, CASE_NAME_ONELINE, align=WD_ALIGN_PARAGRAPH.LEFT)

    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "案例单位", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    merged = c[2].merge(c[3]).merge(c[4]).merge(c[5])
    set_cell(merged, "【请填写申报单位全称】", align=WD_ALIGN_PARAGRAPH.LEFT)

    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "类别", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    merged = c[2].merge(c[3]).merge(c[4]).merge(c[5])
    set_cell(merged, "□ 区（县）教科研院所　　□ 评估监测机构　　□ 高校　　□ 中小学校（含中职学校）　　□ 幼儿园\n（请按本单位性质勾选）",
             align=WD_ALIGN_PARAGRAPH.LEFT, size=11)

    c = add_row()
    set_cell(c[0], "作者姓名", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    set_cell(c[1], "【请填写】", align=WD_ALIGN_PARAGRAPH.CENTER)
    c[2].merge(c[3])
    set_cell(c[2], "作者单位", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[2])
    c[4].merge(c[5])
    set_cell(c[4], "【请填写】", align=WD_ALIGN_PARAGRAPH.CENTER)

    c = add_row()
    set_cell(c[0], "联系电话\n（手机）", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    set_cell(c[1], "【请填写】", align=WD_ALIGN_PARAGRAPH.CENTER)
    c[2].merge(c[3])
    set_cell(c[2], "电子邮箱", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[2])
    c[4].merge(c[5])
    set_cell(c[4], "【请填写】", align=WD_ALIGN_PARAGRAPH.CENTER)

    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "案例所获\n区级及以上\n奖励情况", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    merged = c[2].merge(c[3]).merge(c[4]).merge(c[5])
    set_cell(merged, "无\n（如未获奖，填写“无”；如填写获奖情况，须提供相应佐证材料）",
             align=WD_ALIGN_PARAGRAPH.LEFT, size=11)

    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "解决的\n主要问题", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    merged = c[2].merge(c[3]).merge(c[4]).merge(c[5])
    set_cell(merged,
             "（分条罗列关键点，150字以内）\n"
             "1.标准执行不统一，指标、评分、举证要求复杂且缺乏数字化承载，各校理解执行易偏差；\n"
             "2.审核评审效率低，材料校验仅停留在格式、数量层面，内容判断与专家评阅依赖人工，耗时长、意见易模板化；\n"
             "3.结果运用支撑弱，结果多为分数明细，缺乏结构化诊断，难以反哺培养改进。",
             align=WD_ALIGN_PARAGRAPH.LEFT, size=12)

    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "采取的\n主要举措", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    merged = c[2].merge(c[3]).merge(c[4]).merge(c[5])
    set_cell(merged,
             "（分条罗列关键点，200字以内）\n"
             "研发师范生教育教学能力认定抽检系统，创新嵌入“三智能体”AI赋能体系：\n"
             "1.佐证材料智能预检，赋能学生自检与校级审核；\n"
             "2.专家评审AI辅助，自动生成结构化参考意见；\n"
             "3.认定结果智能诊断，自动生成学校诊断报告。\n"
             "同步推动标准数字化、流程闭环化、规则公平化、成果可复制，以“数字化+智能化”双轮驱动抽检质量监测。",
             align=WD_ALIGN_PARAGRAPH.LEFT, size=12)

    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "取得的\n具体成效", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    merged = c[2].merge(c[3]).merge(c[4]).merge(c[5])
    set_cell(merged,
             "（分条罗列关键点，200字以内）\n"
             "1.智能预检使单份材料审核时长由约20分钟降至约10分钟（约50%），一次性通过率提升约30个百分点；\n"
             "2.AI辅助使单生评审时长由约30分钟降至约17分钟（约43%），评审意见质量提升；\n"
             "3.智能诊断报告生成由约1天降至约90分钟，结果从“分数通报”转向“质量反哺”。\n"
             "系统已支撑全市10余所高校、30余个专业、500余名师范生。",
             align=WD_ALIGN_PARAGRAPH.LEFT, size=12)

    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "案例作者\n所属单位\n推荐意见", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    merged = c[2].merge(c[3]).merge(c[4]).merge(c[5])
    set_cell(merged, "\n\n\n\n单位（盖章）\n\n　　　　　　　　　　　　　　　　　年　　月　　日",
             align=WD_ALIGN_PARAGRAPH.LEFT, size=12)

    widths = [Cm(2.0), Cm(2.0), Cm(3.0), Cm(3.0), Cm(3.0), Cm(3.0)]
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            if idx < len(widths):
                cell.width = widths[idx]

    out = os.path.join(BASE, "附件1 教育评价领域典型实践案例征集表.docx")
    doc.save(out)
    return out


if __name__ == "__main__":
    f1, cn = build_case_doc()
    f2 = build_form_doc()
    print("已生成：")
    print(" -", f1, "（正文中文字数约 %d 字）" % cn)
    print(" -", f2)
