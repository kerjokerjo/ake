"""
技术路线图（活页简化版 · V2）
   四层结构：问题 — 目标 — 内容 — 成果
   内容层三块之间建立横向联动 + 弧形反哺闭环
"""
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import matplotlib as mpl
import numpy as np

mpl.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'Arial Unicode MS']
mpl.rcParams['axes.unicode_minus'] = False

fig, ax = plt.subplots(figsize=(17, 11.5), dpi=150)
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

# ============== 配色 ==============
C_BG        = '#FAFBFD'
C_PROBLEM   = '#FDECEA'; C_PROBLEM_E = '#C0392B'
C_GOAL      = '#FFF4E0'; C_GOAL_E    = '#D68910'
C_RQ1       = '#E8F4FD'; C_RQ1_E     = '#2874A6'
C_RQ2       = '#E8F8EF'; C_RQ2_E     = '#1E8449'
C_RQ3       = '#F4ECF7'; C_RQ3_E     = '#6C3483'
C_LINK      = '#5D6D7E'
C_LOOP      = '#D35400'
C_OUT       = '#FEF9E7'; C_OUT_E     = '#B7950B'

ax.add_patch(plt.Rectangle((0, 0), 100, 100, facecolor=C_BG, edgecolor='none', zorder=0))

# ============== 标题 ==============
ax.text(50, 97.0,
        '图  "三图谱+双智能体" 赋能专业核心课程建设技术路线图（简版）',
        ha='center', va='center', fontsize=16.5, fontweight='bold', color='#1B2631')
ax.text(50, 94.2,
        '问题 → 目标 → 内容 → 成果 · 四层递进 · 内容层闭环联动',
        ha='center', va='center', fontsize=11.2, color='#566573', style='italic')

def arrow(x1, y1, x2, y2, color='#566573', style='-|>', lw=1.8, mutation=20, zorder=2,
          connectionstyle="arc3,rad=0"):
    ar = FancyArrowPatch((x1, y1), (x2, y2),
                         arrowstyle=style, mutation_scale=mutation,
                         color=color, linewidth=lw, zorder=zorder,
                         connectionstyle=connectionstyle)
    ax.add_patch(ar)

def layer_label(y, text, color, h=4.5):
    ax.add_patch(plt.Rectangle((1.0, y-h/2), 4.6, h,
                               facecolor=color, edgecolor='none', alpha=0.9, zorder=2))
    ax.text(3.3, y, text, ha='center', va='center',
            fontsize=11.2, fontweight='bold', color='white', zorder=3)

# =================================================================
# 第 1 层：问题层                                      y ≈ 81 — 92
# =================================================================
layer_label(86.5, '问题层\n三重断裂', C_PROBLEM_E, h=9)

prob_y, prob_h = 81.5, 10.0
pw = 25
pgap = (100 - 8 - 3*pw) / 2
px0 = 8
probs = [
    ("① \"图—学\" 断裂",
     "知识/能力图谱未连通\"问题侧\"数据\n学情诊断粗放 · 难以精准定位学习障碍"),
    ("② \"AI—用\" 断裂",
     "通用 AI 难以嵌入爬虫强实践/对抗/合规场景\n伴学缺位 · 合规悬空 · 重技术轻伦理"),
    ("③ \"教—评\" 断裂",
     "代码/调试/合规数据未被用于过程评价\n德技分离 · 改革经验缺乏可迁移载体"),
]
prob_centers = []
for i, (title, sub) in enumerate(probs):
    x = px0 + i * (pw + pgap)
    ax.add_patch(FancyBboxPatch((x, prob_y), pw, prob_h,
                                boxstyle="round,pad=0.02,rounding_size=0.8",
                                facecolor=C_PROBLEM, edgecolor=C_PROBLEM_E,
                                linewidth=2.0, zorder=3))
    ax.text(x + pw/2, prob_y + prob_h - 1.8, title,
            ha='center', va='center', fontsize=11.8, fontweight='bold',
            color=C_PROBLEM_E, zorder=4)
    ax.add_patch(plt.Rectangle((x+1.5, prob_y + prob_h - 3.2), pw-3, 0.08,
                               facecolor=C_PROBLEM_E, alpha=0.4, zorder=4))
    ax.text(x + pw/2, prob_y + prob_h/2 - 1.2, sub,
            ha='center', va='center', fontsize=9.8, color='#515A5A', zorder=4)
    prob_centers.append((x + pw/2, prob_y, x + pw/2, prob_y + prob_h))

# =================================================================
# 第 2 层：目标层                                      y ≈ 70 — 79
# =================================================================
layer_label(74.5, '目标层\n总目标', C_GOAL_E, h=7.5)

goal_y, goal_h = 70.5, 8.0
ax.add_patch(FancyBboxPatch((8, goal_y), 85, goal_h,
                            boxstyle="round,pad=0.02,rounding_size=0.7",
                            facecolor=C_GOAL, edgecolor=C_GOAL_E,
                            linewidth=2.0, zorder=3))
ax.text(50, goal_y + goal_h - 1.8,
        '建成以"三图谱 + 双智能体"为核心特征的专业核心课程样板',
        ha='center', va='center', fontsize=13, fontweight='bold',
        color=C_GOAL_E, zorder=4)

# 八位一体标签
tags = ['一标准', '三图谱', '两智能体', '一靶场', '一模式', '一评价']
tag_y = goal_y + 1.6
tag_w = 11.5
total_w = len(tags) * tag_w + (len(tags)-1)*1.5
tx0 = 50 - total_w/2
for i, t in enumerate(tags):
    xi = tx0 + i*(tag_w + 1.5)
    ax.add_patch(FancyBboxPatch((xi, tag_y), tag_w, 2.4,
                                boxstyle="round,pad=0.02,rounding_size=0.3",
                                facecolor='white', edgecolor=C_GOAL_E, linewidth=1.2, zorder=4))
    ax.text(xi + tag_w/2, tag_y + 1.2, t, ha='center', va='center',
            fontsize=10.2, fontweight='bold', color=C_GOAL_E, zorder=5)

# 问题 → 目标（汇聚）
for cx, py, _, _ in prob_centers:
    arrow(cx, py, 50, goal_y + goal_h, color=C_PROBLEM_E, lw=1.1, mutation=12)

# =================================================================
# 第 3 层：内容层 —— 三条主线 + 联动闭环                y ≈ 26 — 65
# =================================================================
layer_label(47, '内容层\n三条主线\n闭环联动', '#2874A6', h=12)

rc_y, rc_h = 28, 36
rc_w = 23
rc_gap = (100 - 8 - 3*rc_w) / 2
rc_x0 = 8

rc_data = [
    {
        'color_bg': C_RQ1, 'color_ec': C_RQ1_E,
        'num': 'Ⅰ', 'layer_tag': '认知底座',
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
        'color_bg': C_RQ2, 'color_ec': C_RQ2_E,
        'num': 'Ⅱ', 'layer_tag': '执行引擎',
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
        'color_bg': C_RQ3, 'color_ec': C_RQ3_E,
        'num': 'Ⅲ', 'layer_tag': '评价+迁移',
        'title': '"双师+双工单" 模式与\n五维评价·迁移推广',
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

rc_boxes = []  # 记录 (x, rc_y, rc_w, rc_h, center_x, top_y, bot_y, ec, mid_y)
for i, d in enumerate(rc_data):
    x = rc_x0 + i*(rc_w + rc_gap)
    ax.add_patch(FancyBboxPatch((x, rc_y), rc_w, rc_h,
                                boxstyle="round,pad=0.02,rounding_size=0.9",
                                facecolor=d['color_bg'], edgecolor=d['color_ec'],
                                linewidth=2.2, zorder=3))
    # 编号徽章
    ax.add_patch(plt.Circle((x+3.5, rc_y + rc_h - 3.2), 1.9,
                            facecolor=d['color_ec'], edgecolor='none', zorder=4))
    ax.text(x+3.5, rc_y + rc_h - 3.2, d['num'],
            ha='center', va='center', fontsize=12.5, fontweight='bold',
            color='white', zorder=5)
    # 分层标签（认知底座 / 执行引擎 / 评价+迁移）
    ax.add_patch(FancyBboxPatch((x + rc_w - 8.5, rc_y + rc_h - 4.4), 7, 2.5,
                                boxstyle="round,pad=0.02,rounding_size=0.25",
                                facecolor='white', edgecolor=d['color_ec'],
                                linewidth=1.2, zorder=4))
    ax.text(x + rc_w - 5.0, rc_y + rc_h - 3.2, d['layer_tag'],
            ha='center', va='center', fontsize=9.3, fontweight='bold',
            color=d['color_ec'], zorder=5)
    # 标题
    ax.text(x + rc_w/2, rc_y + rc_h - 7.2, d['title'],
            ha='center', va='center', fontsize=11.5, fontweight='bold',
            color=d['color_ec'], zorder=4)
    # 对应关系
    ax.text(x + rc_w/2, rc_y + rc_h - 10.2, d['subtitle'],
            ha='center', va='center', fontsize=9.4, color='#566573',
            style='italic', zorder=4)
    # 分割线
    ax.add_patch(plt.Rectangle((x+2, rc_y + rc_h - 11.8), rc_w-4, 0.08,
                               facecolor=d['color_ec'], alpha=0.45, zorder=4))
    # 子项
    for j, item in enumerate(d['items']):
        ax.add_patch(plt.Circle((x + 2.3, rc_y + rc_h - 14.0 - j*3.8), 0.35,
                                facecolor=d['color_ec'], edgecolor='none', zorder=4))
        ax.text(x + 3.5, rc_y + rc_h - 14.0 - j*3.8, item,
                ha='left', va='center', fontsize=9.5, color='#1B2631', zorder=4)
    rc_boxes.append({
        'x_left': x, 'x_right': x + rc_w,
        'cx': x + rc_w/2, 'top_y': rc_y + rc_h, 'bot_y': rc_y,
        'mid_y': rc_y + rc_h - 9.5,   # 箭头挂接点
        'ec': d['color_ec'],
    })

# ----- 问题层 → 内容层：三条"破解"箭头 -----
for i in range(3):
    arrow(prob_centers[i][0], prob_centers[i][1], rc_boxes[i]['cx'], rc_boxes[i]['top_y'],
          color=rc_boxes[i]['ec'], lw=1.4, mutation=14, style='-|>',
          connectionstyle="arc3,rad=0.0")

# ----- 内容层之间 · 横向联动（顶部）-----
# Ⅰ → Ⅱ  &  Ⅱ → Ⅲ
band_y_top = rc_boxes[0]['top_y'] - 4.5     # 横向箭头所在高度
labels_top = [
    ('数据语义供给', 'RAG 检索源', C_LINK),
    ('过程数据沉淀', '合规证据链', C_LINK),
]
for i in range(2):
    xa = rc_boxes[i]['x_right'] + 0.3
    xb = rc_boxes[i+1]['x_left'] - 0.3
    arrow(xa, band_y_top, xb, band_y_top,
          color=C_LINK, lw=2.3, mutation=22, style='-|>')
    # 顶部标签
    mx = (xa + xb) / 2
    ax.add_patch(FancyBboxPatch((mx-4.5, band_y_top+1.0), 9, 4,
                                boxstyle="round,pad=0.02,rounding_size=0.35",
                                facecolor='white', edgecolor=C_LINK,
                                linewidth=1.2, zorder=5))
    ax.text(mx, band_y_top+3.3, labels_top[i][0], ha='center', va='center',
            fontsize=9.8, fontweight='bold', color=C_LINK, zorder=6)
    ax.text(mx, band_y_top+1.8, labels_top[i][1], ha='center', va='center',
            fontsize=8.8, color=C_LINK, style='italic', zorder=6)

# ----- 内容层之间 · 弧形反哺闭环（底部回流 Ⅲ → Ⅰ）-----
loop_y = rc_boxes[0]['bot_y'] - 2.5
# 从 Ⅲ 底部 → 弧线回流 → Ⅰ 底部
arrow(rc_boxes[2]['cx'], rc_boxes[2]['bot_y'] - 0.0,
      rc_boxes[0]['cx'], rc_boxes[0]['bot_y'] - 0.0,
      color=C_LOOP, lw=2.3, mutation=22, style='-|>',
      connectionstyle="arc3,rad=-0.18")

# 闭环标签
ax.add_patch(FancyBboxPatch((38, loop_y-2.3), 24, 4.0,
                            boxstyle="round,pad=0.02,rounding_size=0.35",
                            facecolor='white', edgecolor=C_LOOP,
                            linewidth=1.4, zorder=5))
ax.text(50, loop_y-0.2, '↻ 评价反哺 · 学习分析',
        ha='center', va='center', fontsize=10.5, fontweight='bold',
        color=C_LOOP, zorder=6)
ax.text(50, loop_y-1.7, '驱动问题图谱动态演化，形成"诊断—伴学—评价—反哺"闭环',
        ha='center', va='center', fontsize=8.8, color=C_LOOP,
        style='italic', zorder=6)

# ----- 递进小标语 -----
ax.text(rc_boxes[0]['cx'], rc_y - 5.2,
        '底座', ha='center', va='center',
        fontsize=10.5, fontweight='bold', color=C_RQ1_E, zorder=6,
        bbox=dict(boxstyle='round,pad=0.35', facecolor='white',
                  edgecolor=C_RQ1_E, linewidth=1.2))
ax.text(rc_boxes[1]['cx'], rc_y - 5.2,
        '引擎', ha='center', va='center',
        fontsize=10.5, fontweight='bold', color=C_RQ2_E, zorder=6,
        bbox=dict(boxstyle='round,pad=0.35', facecolor='white',
                  edgecolor=C_RQ2_E, linewidth=1.2))
ax.text(rc_boxes[2]['cx'], rc_y - 5.2,
        '闭环', ha='center', va='center',
        fontsize=10.5, fontweight='bold', color=C_RQ3_E, zorder=6,
        bbox=dict(boxstyle='round,pad=0.35', facecolor='white',
                  edgecolor=C_RQ3_E, linewidth=1.2))

# =================================================================
# 第 4 层：成果层                                      y ≈ 5 — 18
# =================================================================
layer_label(11.5, '成果层\n八位一体', C_OUT_E, h=8)

out_y, out_h = 7.5, 11
ax.add_patch(FancyBboxPatch((8, out_y), 85, out_h,
                            boxstyle="round,pad=0.02,rounding_size=0.6",
                            facecolor='#FFFDF4', edgecolor=C_OUT_E,
                            linewidth=1.8, zorder=2))
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
    ax.add_patch(FancyBboxPatch((x+0.12, out_y+0.8), on_w-0.24, out_h-1.6,
                                boxstyle="round,pad=0.02,rounding_size=0.35",
                                facecolor=C_OUT, edgecolor=C_OUT_E,
                                linewidth=1.2, zorder=3))
    ax.text(x + on_w/2, out_y + out_h - 2.6, t,
            ha='center', va='center', fontsize=10, fontweight='bold',
            color='#7D6608', zorder=4)
    ax.text(x + on_w/2, out_y + 2.6, s,
            ha='center', va='center', fontsize=9, color=C_OUT_E,
            style='italic', zorder=4)

# 内容 → 成果 向下汇聚箭头
for b in rc_boxes:
    arrow(b['cx'], b['bot_y'] - 0.1, b['cx'], out_y + out_h,
          color='#7D6608', lw=1.4, mutation=16,
          connectionstyle="arc3,rad=0.0")

# =================================================================
# 底部图例
# =================================================================
legend_elems = [
    mpatches.Patch(facecolor=C_PROBLEM, edgecolor=C_PROBLEM_E, label='问题层 · 三重断裂'),
    mpatches.Patch(facecolor=C_GOAL,    edgecolor=C_GOAL_E,    label='目标层 · 八位一体'),
    mpatches.Patch(facecolor=C_RQ1,     edgecolor=C_RQ1_E,     label='Ⅰ 三图谱 · 认知底座'),
    mpatches.Patch(facecolor=C_RQ2,     edgecolor=C_RQ2_E,     label='Ⅱ 双智能体 · 执行引擎'),
    mpatches.Patch(facecolor=C_RQ3,     edgecolor=C_RQ3_E,     label='Ⅲ 模式评价 · 闭环迁移'),
    mpatches.Patch(facecolor=C_OUT,     edgecolor=C_OUT_E,     label='成果层 · 10 项交付'),
]
ax.legend(handles=legend_elems, loc='lower center', bbox_to_anchor=(0.5, -0.015),
          ncol=6, frameon=False, fontsize=9.4, handlelength=1.4, columnspacing=2.0)

plt.subplots_adjust(left=0.01, right=0.99, top=0.98, bottom=0.045)

out_png = '/workspace/technical_roadmap_simple.png'
out_pdf = '/workspace/technical_roadmap_simple.pdf'
out_svg = '/workspace/technical_roadmap_simple.svg'
out_emf = '/workspace/technical_roadmap_simple.emf'
plt.savefig(out_png, dpi=220, bbox_inches='tight', facecolor='white')
plt.savefig(out_pdf, bbox_inches='tight', facecolor='white')
plt.savefig(out_svg, bbox_inches='tight', facecolor='white')
try:
    plt.savefig(out_emf, bbox_inches='tight', facecolor='white')
    print('Saved:', out_emf)
except Exception as e:
    print('EMF skipped:', e)
print('Saved:', out_png)
print('Saved:', out_pdf)
print('Saved:', out_svg)
