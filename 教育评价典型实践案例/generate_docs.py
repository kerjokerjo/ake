# -*- coding: utf-8 -*-
"""生成两份提交材料：
1. 师范生教育教学能力认定抽检评价改革实践——案例文稿.docx
2. 附件1 教育评价领域典型实践案例征集表.docx
"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = os.path.dirname(os.path.abspath(__file__))

CASE_TITLE = "数字赋能 标准统一\n——师范生教育教学能力认定抽检全流程数字化评价改革实践"
CASE_TITLE_ONELINE = "数字赋能 标准统一——师范生教育教学能力认定抽检全流程数字化评价改革实践"


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


def add_para(doc, text, cn_font="仿宋", size=14, bold=False, align=None,
             first_indent=True, line_spacing=1.5, space_after=0):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    pf = p.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(0)
    if first_indent:
        pf.first_line_indent = Pt(size * 2)
    run = p.add_run(text)
    set_font(run, cn_font=cn_font, size=size, bold=bold)
    return p


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

    # 标题
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run("数字赋能　标准统一")
    set_font(run, cn_font="方正小标宋简体", size=20, bold=True)
    p2 = doc.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after = Pt(16)
    run2 = p2.add_run("——师范生教育教学能力认定抽检全流程数字化评价改革实践")
    set_font(run2, cn_font="黑体", size=15, bold=False)

    def heading(text):
        ph = doc.add_paragraph()
        ph.paragraph_format.space_before = Pt(10)
        ph.paragraph_format.space_after = Pt(4)
        ph.paragraph_format.line_spacing = 1.5
        r = ph.add_run(text)
        set_font(r, cn_font="黑体", size=15, bold=True)

    def body(text):
        add_para(doc, text, cn_font="仿宋", size=14, line_spacing=1.5)

    # 一、案例背景
    heading("一、案例背景")
    body("2021年教育部印发《关于推进师范生免试认定中小学教师资格改革的通知》，"
         "参加改革的师范生不再参加国家中小学教师资格考试，改由所在高校自主组织教育教学能力考核认定；"
         "省级教育行政部门及专业评价机构须对高校认定工作开展抽检，以保障“放权不降标准、自主不失质量”。"
         "重庆市师范生培养规模大、院校多、学科专业分布广，传统抽检主要依赖线下集中、纸质举证、人工抽签和专家现场评审，"
         "存在四方面突出矛盾：一是抽样不够规范，人工抽取易受主观影响，代表性与公平性难以保证；"
         "二是标准难统一，各校指标口径、举证要求不一致，跨校结果不可比；"
         "三是过程不透明、难追溯，纸质材料流转环节多、留痕弱；"
         "四是数据汇聚难，专家评审分散、进度不可视、结果统计耗时长。"
         "在新时代教育评价改革“破五唯”、重过程、强证据、用数据的导向下，亟需以数字化手段重构师范生能力认定抽检的评价流程。")

    # 二、具体举措及实施成效
    heading("二、具体举措及实施成效")
    body("围绕“科学抽样、标准统一、全程留痕、数据驱动”的思路，自主设计开发了“师范生教育教学能力认定抽检平台”，"
         "并在重庆市相关高校连续运行【X】年，累计覆盖【XX】所院校、【XX】个师范专业、【XX】名师范生。主要举措与成效如下：")

    body("（一）构建可配置、可复用的标准化指标工具库。系统以“指标表—二级指标—三级指标—评分项”分层建模，"
         "专家依据三级指标打分，支持等级制与分数制并存，并可为各等级配置对应分值、实现“等级自动折算分数”；"
         "举证材料与三级指标建立多对多映射，同一材料可被多条指标引用，并可设置最小/最大上传数量、文件大小、是否仅校级管理员上传等约束。"
         "工具库支持一键复制，便于跨批次、跨学科快速复用。"
         "成效：评价标准由“一校一策”转为“全市一把尺”，指标口径与举证要求统一，跨校结果具备可比性，新建一套评价工具的配置时间显著缩短。")

    body("（二）实施科学规范的等间隔随机抽样。按专业人数分档动态设置抽取比例，对同校同专业学生按序等距抽取，"
         "兼顾抽样的随机性与代表性；抽取结果支持评估院管理员手动微调并确认上传名单，学校可一键导出被抽中学生名单及登录信息以组织上传。"
         "成效：抽样过程由“人工抽签”升级为“算法等距抽取＋可追溯确认”，压缩了人为干预空间，抽样的公平性与公信力明显提升。")

    body("（三）打通“上传—审核—评审—反馈”全流程线上闭环。被抽中学生在线上传PDF/MP4举证材料，"
         "经校级管理员审核（并可代传指定为校级上传的举证材料）、评估院管理员复核后进入专家评审；"
         "专家基于指标在线查看佐证材料、打分（或评等级）并撰写评审意见；评估院最终以PDF形式向各校反馈整体抽检结果。"
         "成效：举证与评审全程电子化、全程留痕，告别纸质材料邮寄与现场集中，评审周期大幅压缩，过程可回溯、可审计。")

    body("（四）实施灵活的多专家交叉评审与分级权限管理。支持一名学生关联多位专家交叉评审，"
         "支持“一名专家评多校全部指标”或“一名专家评指定指标、覆盖多校全部待评学生”等多种配置；"
         "系统按系统管理员、评估院管理员、校级管理员、专家、学生分级授权，并留存登录与操作日志。"
         "成效：交叉评审降低了单一专家偏差，评分客观性增强；权限与日志机制保障了数据安全与责任可追溯。")

    body("（五）实现进度可视与数据驱动的结果运用。评估院管理员可实时查看各专家评审进度、调度评审工作，"
         "并导出学生总成绩及各指标得分明细。"
         "成效：管理者由“事后统计”转向“过程调度”，结果数据可直接用于诊断各校师范生能力短板，反哺人才培养。")

    # 三、经验总结
    heading("三、经验总结")
    body("一是坚持问题导向与标准先行。先统一指标与举证标准、再用技术固化标准，是实现跨校可比、公平公正的前提。")
    body("二是以“过程留痕＋数据说话”落实评价改革精神。全流程电子化使评价从“看材料”转向“重证据、可追溯”，"
         "契合“破五唯”与过程性、增值性评价导向。")
    body("三是注重机制的可配置与可复制。工具库分层建模与一键复制，使平台能以较低成本适配不同学科、不同批次，"
         "具备向师范生评价之外同类能力认定抽检场景推广的潜力。")
    body("四是抽检的根本目的是以评促建。应强化结果反馈与闭环运用，让抽检数据真正服务于师范人才培养质量的持续提升。"
         "本案例可为各级评价机构开展大规模、跨单位、标准化的能力认定抽检提供可复制、能推广的实践范式。")

    out = os.path.join(BASE, "师范生教育教学能力认定抽检评价改革实践-案例文稿.docx")
    doc.save(out)
    return out


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

    # 附件标识
    pa = doc.add_paragraph()
    pa.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    ra = pa.add_run("附件1")
    set_font(ra, cn_font="黑体", size=14, bold=True)

    # 标题
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

    # 案例类别
    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "案例类别", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    set_cell(c[2], "□ 学校发展", align=WD_ALIGN_PARAGRAPH.CENTER)
    set_cell(c[3], "□ 教师发展", align=WD_ALIGN_PARAGRAPH.CENTER)
    c[4].merge(c[5])
    set_cell(c[4], "☑ 学生成长", align=WD_ALIGN_PARAGRAPH.CENTER)

    # 案例名称
    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "案例名称", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    merged = c[2].merge(c[3]).merge(c[4]).merge(c[5])
    set_cell(merged, CASE_TITLE_ONELINE, align=WD_ALIGN_PARAGRAPH.LEFT)

    # 案例单位 + 类别
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

    # 作者姓名 / 作者单位
    c = add_row()
    set_cell(c[0], "作者姓名", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    set_cell(c[1], "【请填写】", align=WD_ALIGN_PARAGRAPH.CENTER)
    c[2].merge(c[3])
    set_cell(c[2], "作者单位", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[2])
    c[4].merge(c[5])
    set_cell(c[4], "【请填写】", align=WD_ALIGN_PARAGRAPH.CENTER)

    # 联系电话 / 电子邮箱
    c = add_row()
    set_cell(c[0], "联系电话\n（手机）", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    set_cell(c[1], "【请填写】", align=WD_ALIGN_PARAGRAPH.CENTER)
    c[2].merge(c[3])
    set_cell(c[2], "电子邮箱", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[2])
    c[4].merge(c[5])
    set_cell(c[4], "【请填写】", align=WD_ALIGN_PARAGRAPH.CENTER)

    # 获奖情况
    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "案例所获\n区级及以上\n奖励情况", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    merged = c[2].merge(c[3]).merge(c[4]).merge(c[5])
    set_cell(merged, "无\n（如未获奖，填写“无”；如填写获奖情况，须提供相应佐证材料）",
             align=WD_ALIGN_PARAGRAPH.LEFT, size=11)

    # 解决的主要问题
    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "解决的\n主要问题", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    merged = c[2].merge(c[3]).merge(c[4]).merge(c[5])
    set_cell(merged,
             "（分条罗列关键点，150字以内）\n"
             "1.传统人工抽签代表性、公平性不足；\n"
             "2.各校指标口径、举证要求不统一，跨校结果不可比；\n"
             "3.纸质举证流转环节多、过程不透明、难留痕追溯；\n"
             "4.专家评审分散，进度不可视、结果统计汇聚效率低。",
             align=WD_ALIGN_PARAGRAPH.LEFT, size=12)

    # 采取的主要举措
    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "采取的\n主要举措", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    merged = c[2].merge(c[3]).merge(c[4]).merge(c[5])
    set_cell(merged,
             "（分条罗列关键点，200字以内）\n"
             "1.搭建“指标表—二三级指标—评分项”可配置、可复制的标准化工具库，等级制/分数制并存并自动折算；\n"
             "2.按专业人数分档、等间隔等距随机抽取并可追溯确认名单；\n"
             "3.打通“学生上传—校级审核—评估院复核—专家评审—结果反馈”全流程线上闭环；\n"
             "4.支持多专家交叉评审、分级权限与操作留痕；\n"
             "5.评审进度可视、得分明细可导出。",
             align=WD_ALIGN_PARAGRAPH.LEFT, size=12)

    # 取得的具体成效
    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "取得的\n具体成效", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    merged = c[2].merge(c[3]).merge(c[4]).merge(c[5])
    set_cell(merged,
             "（分条罗列关键点，200字以内）\n"
             "1.评价标准全市统一、跨校结果可比，工具配置效率显著提升；\n"
             "2.抽样由人工抽签升级为算法等距抽取，公平性与公信力增强；\n"
             "3.举证与评审全程电子化、全程留痕，评审周期大幅压缩、过程可审计；\n"
             "4.交叉评审与权限日志机制提升评分客观性与数据安全；\n"
             "5.进度可视、数据可导出，结果用于诊断改进、以评促建。平台已连续运行【X】年，覆盖【XX】所院校。",
             align=WD_ALIGN_PARAGRAPH.LEFT, size=12)

    # 推荐意见
    c = add_row()
    c[0].merge(c[1])
    set_cell(c[0], "案例作者\n所属单位\n推荐意见", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    shade_cell(c[0])
    merged = c[2].merge(c[3]).merge(c[4]).merge(c[5])
    set_cell(merged, "\n\n\n\n单位（盖章）\n\n　　　　　　　　　　　　　　　　　年　　月　　日",
             align=WD_ALIGN_PARAGRAPH.LEFT, size=12)

    # 列宽设置
    widths = [Cm(2.0), Cm(2.0), Cm(3.0), Cm(3.0), Cm(3.0), Cm(3.0)]
    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            if idx < len(widths):
                cell.width = widths[idx]

    out = os.path.join(BASE, "附件1 教育评价领域典型实践案例征集表.docx")
    doc.save(out)
    return out


if __name__ == "__main__":
    f1 = build_case_doc()
    f2 = build_form_doc()
    print("已生成：")
    print(" -", f1)
    print(" -", f2)
