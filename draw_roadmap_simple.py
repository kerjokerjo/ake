"""
技术路线图（活页简化版）
   —— 仅保留 "问题—内容—成果" 三层，不含时间轴，横版 A4
"""
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import matplotlib as mpl

mpl.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'Arial Unicode MS']
mpl.rcParams['axes.unicode_minus'] = False

fig, ax = plt.subplots(figsize=(16, 10), dpi=150)  # 横版 A4 比例
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

# ============== 配色 ==============
C_BG      = '#FAFBFD'
C_PROBLEM = '#FDECEA'
C_PROBLEM_E = '#C0392B'
C_RQ1     = '#E8F4FD'
C_RQ1_E   = '#2874A6'
C_RQ2     = '#E8F8EF'
C_RQ2_E   = '#1E8449'
C_RQ3     = '#F4ECF7'
C_RQ3_E   = '#6C3483'
C_OUT     = '#FEF9E7'
C_OUT_E   = '#B7950B'

ax.add_patch(plt.Rectangle((0, 0), 100, 100, facecolor=C_BG, edgecolor='none', zorder=0))

# ============== 标题 ==============
ax.text(50, 96.5,
        '图  "三图谱+双智能体" 赋能专业核心课程建设技术路线图（简版）',
        ha='center', va='center', fontsize=17, fontweight='bold', color='#1B2631')
ax.text(50, 93.3,
        '问题层 → 内容层 → 成果层  三位一体对应',
        ha='center', va='center', fontsize=11.5, color='#566573', style='italic')

def arrow(x1, y1, x2, y2, color='#566573', style='-|>', lw=1.8, mutation=20, zorder=2,
          connectionstyle="arc3,rad=0"):
    ar = FancyArrowPatch((x1, y1), (x2, y2),
                         arrowstyle=style, mutation_scale=mutation,
                         color=color, linewidth=lw, zorder=zorder,
                         connectionstyle=connectionstyle)
    ax.add_patch(ar)

def layer_label(y, text, color):
    ax.add_patch(plt.Rectangle((1.0, y-2.2), 4.8, 4.4,
                               facecolor=color, edgecolor='none', alpha=0.88, zorder=2))
    ax.text(3.4, y, text, ha='center', va='center',
            fontsize=11.5, fontweight='bold', color='white', zorder=3)

# =================================================================
# 第 1 层：问题层 —— 三重断裂                           y ≈ 76—88
# =================================================================
layer_label(82, '问题层\n\n三重\n断裂', C_PROBLEM_E)

prob_y = 76.5
prob_h = 10.5
pw = 27
pgap = (100 - 8 - 3*pw) / 2
px0 = 8
probs = [
    ("① \"图—学\" 断裂",
     "知识/能力图谱未连通\n\"问题侧\" 数据；\n学情诊断粗放，\n难以精准定位学习障碍"),
    ("② \"AI—用\" 断裂",
     "通用 AI 工具难以嵌入爬虫\n强实践 / 强对抗 / 强合规场景；\n伴学缺位、合规悬空；\n重技术轻伦理问题突出"),
    ("③ \"教—评\" 断裂",
     "过程评价数据化不足，\n代码/调试/合规行为数据未被利用；\n德技分离，\n改革经验缺乏可迁移载体"),
]
prob_centers = []
for i, (title, sub) in enumerate(probs):
    x = px0 + i * (pw + pgap)
    ax.add_patch(FancyBboxPatch((x, prob_y), pw, prob_h,
                                boxstyle="round,pad=0.02,rounding_size=0.8",
                                facecolor=C_PROBLEM, edgecolor=C_PROBLEM_E,
                                linewidth=2.2, zorder=3))
    ax.text(x + pw/2, prob_y + prob_h - 1.8, title,
            ha='center', va='center', fontsize=12.5, fontweight='bold',
            color=C_PROBLEM_E, zorder=4)
    ax.add_patch(plt.Rectangle((x+2, prob_y + prob_h - 3.2), pw-4, 0.08,
                               facecolor=C_PROBLEM_E, alpha=0.4, zorder=4))
    ax.text(x + pw/2, prob_y + prob_h/2 - 1.8, sub,
            ha='center', va='center', fontsize=10, color='#515A5A', zorder=4)
    prob_centers.append((x + pw/2, prob_y))

# =================================================================
# 第 2 层：内容层 —— 三大研究内容                        y ≈ 30—68
# =================================================================
layer_label(48, '内容层\n\n三条\n主线', '#2874A6')

rc_y, rc_h = 30, 38
rc_w = 27
rc_gap = (100 - 8 - 3*rc_w) / 2
rc_x0 = 8

rc_data = [
    {
        'color_bg': C_RQ1, 'color_ec': C_RQ1_E,
        'num': 'Ⅰ',
        'title': '"三图谱" 构建与\n课程标准重构',
        'subtitle': '对应：图—学断裂 / KQ①',
        'items': [
            '知识图谱  12 模块 · 86 知识点',
            '能力图谱  127 能力点',
            '问题图谱  2300+ 提问 · 400+ 错误',
            '三图谱融合  + Neo4j 建模',
            '课程标准  3 模块 · 5 项目 · 12 任务',
            '五级建模  项目→任务→流程→技能→知识',
        ],
    },
    {
        'color_bg': C_RQ2, 'color_ec': C_RQ2_E,
        'num': 'Ⅱ',
        'title': '"双智能体" 研发\n与场景嵌入',
        'subtitle': '对应：AI—用断裂 / KQ②',
        'items': [
            '底层引擎  私有化 LLM + RAG + LoRA',
            '智学小助  代码伴学·路径推荐·归因',
            'AI 合规官  法规嵌入·预警·审查',
            '仿真靶场  电商/招聘/政务/工业 4 类',
            '反爬等级  10 级对抗场景',
            '红蓝对抗  实战演练模式',
        ],
    },
    {
        'color_bg': C_RQ3, 'color_ec': C_RQ3_E,
        'num': 'Ⅲ',
        'title': '"双师 + 双工单" 模式与\n五维评价·迁移推广',
        'subtitle': '对应：教—评断裂 / KQ③',
        'items': [
            '双师  教师 + AI',
            '双工单  教学工单 + 企业工单',
            '五维评价  代码·调试·协作·合规·伦理',
            '伦理占比 20%  重大违规一票否决',
            '学习分析  精准补偿与过程性闭环',
            '迁移推广  跨课程指南 + 师资培训包',
        ],
    },
]

rc_centers_top = []
rc_centers_bot = []
for i, d in enumerate(rc_data):
    x = rc_x0 + i*(rc_w + rc_gap)
    ax.add_patch(FancyBboxPatch((x, rc_y), rc_w, rc_h,
                                boxstyle="round,pad=0.02,rounding_size=0.9",
                                facecolor=d['color_bg'], edgecolor=d['color_ec'],
                                linewidth=2.2, zorder=3))
    # 罗马编号徽章
    ax.add_patch(plt.Circle((x+3.5, rc_y + rc_h - 3.2), 1.8,
                            facecolor=d['color_ec'], edgecolor='none', zorder=4))
    ax.text(x+3.5, rc_y + rc_h - 3.2, d['num'],
            ha='center', va='center', fontsize=12, fontweight='bold',
            color='white', zorder=5)
    # 标题
    ax.text(x + rc_w/2 + 1.5, rc_y + rc_h - 3.2, d['title'],
            ha='center', va='center', fontsize=12, fontweight='bold',
            color=d['color_ec'], zorder=4)
    # 对应关系
    ax.text(x + rc_w/2, rc_y + rc_h - 6.8, d['subtitle'],
            ha='center', va='center', fontsize=10, color='#566573',
            style='italic', zorder=4)
    # 分割线
    ax.add_patch(plt.Rectangle((x+2, rc_y + rc_h - 8.5), rc_w-4, 0.08,
                               facecolor=d['color_ec'], alpha=0.45, zorder=4))
    # 子项
    for j, item in enumerate(d['items']):
        ax.add_patch(plt.Circle((x + 2.3, rc_y + rc_h - 10.5 - j*3.8), 0.35,
                                facecolor=d['color_ec'], edgecolor='none', zorder=4))
        ax.text(x + 3.5, rc_y + rc_h - 10.5 - j*3.8, item,
                ha='left', va='center', fontsize=10, color='#1B2631', zorder=4)
    rc_centers_top.append((x + rc_w/2, rc_y + rc_h))
    rc_centers_bot.append((x + rc_w/2, rc_y))

# 问题 -> 内容 一一对应（彩色粗箭头）
arrow_colors = [C_RQ1_E, C_RQ2_E, C_RQ3_E]
for i in range(3):
    cx_top, cy_top = prob_centers[i]
    cx_bot, cy_bot = rc_centers_top[i]
    arrow(cx_top, cy_top, cx_bot, cy_bot + 0.2,
          color=arrow_colors[i], lw=2.2, mutation=22,
          style='-|>', connectionstyle="arc3,rad=0.0")
    # 箭头旁文字
    ax.text((cx_top+cx_bot)/2 + 0.8, (cy_top+cy_bot)/2,
            '破解', ha='left', va='center', fontsize=9,
            color=arrow_colors[i], fontweight='bold', style='italic')

# =================================================================
# 第 3 层：成果层 —— 八位一体交付物                     y ≈ 8—22
# =================================================================
layer_label(15, '成果层\n\n八位\n一体', C_OUT_E)

out_y, out_h = 8, 14

# 核心大框
ax.add_patch(FancyBboxPatch((8, out_y), 87, out_h,
                            boxstyle="round,pad=0.02,rounding_size=0.7",
                            facecolor='#FFFDF4', edgecolor=C_OUT_E,
                            linewidth=2, zorder=2))

# 子项：10 个交付物
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
    ('迁移指南', '+ 师资培训包'),
]
on_w = 7.6
on_gap = (87 - len(out_items)*on_w) / (len(out_items)-1)
on_x0 = 8
for i, (t, s) in enumerate(out_items):
    x = on_x0 + i*(on_w + on_gap)
    ax.add_patch(FancyBboxPatch((x+0.15, out_y+1), on_w-0.3, out_h-2,
                                boxstyle="round,pad=0.02,rounding_size=0.35",
                                facecolor=C_OUT, edgecolor=C_OUT_E,
                                linewidth=1.3, zorder=3))
    ax.text(x + on_w/2, out_y + out_h - 3.2, t,
            ha='center', va='center', fontsize=10.2, fontweight='bold',
            color='#7D6608', zorder=4)
    ax.text(x + on_w/2, out_y + 3.5, s,
            ha='center', va='center', fontsize=9.2, color=C_OUT_E,
            style='italic', zorder=4)

# 内容 -> 成果 向下汇聚箭头
for cx, cy in rc_centers_bot:
    arrow(cx, cy, cx, out_y + out_h, color='#7D6608', lw=1.8, mutation=20,
          connectionstyle="arc3,rad=0.0")

# 八位一体标签（成果中部）
ax.text(50, out_y - 1.4,
        '一标准 · 三图谱 · 两智能体 · 一靶场 · 一模式 · 一评价',
        ha='center', va='center', fontsize=11.2, fontweight='bold',
        color=C_OUT_E)

# =================================================================
# 底部图例（精简）
# =================================================================
legend_elems = [
    mpatches.Patch(facecolor=C_PROBLEM, edgecolor=C_PROBLEM_E, label='问题层 · 三重断裂'),
    mpatches.Patch(facecolor=C_RQ1,     edgecolor=C_RQ1_E,    label='Ⅰ 三图谱'),
    mpatches.Patch(facecolor=C_RQ2,     edgecolor=C_RQ2_E,    label='Ⅱ 双智能体'),
    mpatches.Patch(facecolor=C_RQ3,     edgecolor=C_RQ3_E,    label='Ⅲ 模式与评价'),
    mpatches.Patch(facecolor=C_OUT,     edgecolor=C_OUT_E,    label='成果层 · 八位一体'),
]
ax.legend(handles=legend_elems, loc='lower center', bbox_to_anchor=(0.5, -0.02),
          ncol=5, frameon=False, fontsize=10, handlelength=1.4, columnspacing=2.5)

plt.subplots_adjust(left=0.01, right=0.99, top=0.98, bottom=0.04)

out_png = '/workspace/technical_roadmap_simple.png'
out_pdf = '/workspace/technical_roadmap_simple.pdf'
plt.savefig(out_png, dpi=220, bbox_inches='tight', facecolor='white')
plt.savefig(out_pdf, bbox_inches='tight', facecolor='white')
print('Saved:', out_png)
print('Saved:', out_pdf)
