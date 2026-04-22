"""
技术路线图 - "三图谱+双智能体"赋能高职专业核心课程建设的实践研究
               ——以《数据采集技术》为例
"""
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
from matplotlib.lines import Line2D
import matplotlib as mpl

mpl.rcParams['font.sans-serif'] = ['WenQuanYi Micro Hei', 'SimHei', 'Arial Unicode MS']
mpl.rcParams['axes.unicode_minus'] = False

fig, ax = plt.subplots(figsize=(18, 12.5), dpi=150)
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

# ============== 配色 ==============
C_BG      = '#FAFBFD'
C_PROBLEM = '#FDECEA'   # 红
C_PROBLEM_E = '#C0392B'
C_GOAL    = '#FFF4E0'   # 橙
C_GOAL_E  = '#D68910'
C_RQ1     = '#E8F4FD'   # 蓝
C_RQ1_E   = '#2874A6'
C_RQ2     = '#E8F8EF'   # 绿
C_RQ2_E   = '#1E8449'
C_RQ3     = '#F4ECF7'   # 紫
C_RQ3_E   = '#6C3483'
C_PATH    = '#EAF2F8'
C_PATH_E  = '#1A5276'
C_OUT     = '#FEF9E7'
C_OUT_E   = '#B7950B'
C_TIME    = '#D5DBDB'

# 背景
ax.add_patch(plt.Rectangle((0, 0), 100, 100, facecolor=C_BG, edgecolor='none', zorder=0))

# ============== 标题 ==============
ax.text(50, 97.5,
        '图  "三图谱+双智能体"赋能高职专业核心课程建设技术路线图',
        ha='center', va='center', fontsize=17, fontweight='bold', color='#1B2631')
ax.text(50, 94.6,
        '—— 以《数据采集技术》课程为例',
        ha='center', va='center', fontsize=12.5, color='#566573', style='italic')

# ============== 通用绘图函数 ==============
def rbox(x, y, w, h, text, fc, ec, fs=10.2, fw='normal', color='#1B2631',
         lw=1.4, rad=0.02, ha='center', va='center', zorder=3):
    box = FancyBboxPatch((x, y), w, h,
                         boxstyle=f"round,pad=0.02,rounding_size={rad*25}",
                         linewidth=lw, edgecolor=ec, facecolor=fc, zorder=zorder)
    ax.add_patch(box)
    ax.text(x + w/2, y + h/2, text, ha=ha, va=va, fontsize=fs,
            fontweight=fw, color=color, zorder=zorder+1)

def arrow(x1, y1, x2, y2, color='#566573', style='-|>', lw=1.6, mutation=18, zorder=2,
          connectionstyle="arc3,rad=0"):
    ar = FancyArrowPatch((x1, y1), (x2, y2),
                         arrowstyle=style, mutation_scale=mutation,
                         color=color, linewidth=lw, zorder=zorder,
                         connectionstyle=connectionstyle)
    ax.add_patch(ar)

def layer_label(y, text, color):
    ax.add_patch(plt.Rectangle((1.2, y-1.1), 4.2, 2.2,
                               facecolor=color, edgecolor='none', alpha=0.85, zorder=2))
    ax.text(3.3, y, text, ha='center', va='center',
            fontsize=11, fontweight='bold', color='white', zorder=3, rotation=0)

# =================================================================
# 第 1 层：问题层 —— 三重断裂                           y ≈ 85—91
# =================================================================
layer_label(88, '问题层\n三重断裂', C_PROBLEM_E)

prob_y = 85.5
prob_h = 6.0
pw = 27
pgap = (100 - 7 - 3*pw) / 2  # 从 x=7 起
px0 = 7
probs = [
    ("① \"图—学\" 断裂",
     "知识/能力图谱未连通\"问题侧\"数据\n学情诊断粗放，难以精准定位学习障碍"),
    ("② \"AI—用\" 断裂",
     "通用 AI 工具难嵌入爬虫强实践/强对抗/强合规场景\n伴学缺位、合规悬空"),
    ("③ \"教—评\" 断裂",
     "过程评价数据化不足，德技分离\n改革经验缺乏可迁移载体"),
]
prob_centers = []
for i, (title, sub) in enumerate(probs):
    x = px0 + i * (pw + pgap)
    ax.add_patch(FancyBboxPatch((x, prob_y), pw, prob_h,
                                boxstyle="round,pad=0.02,rounding_size=0.6",
                                facecolor=C_PROBLEM, edgecolor=C_PROBLEM_E,
                                linewidth=1.8, zorder=3))
    ax.text(x + pw/2, prob_y + prob_h - 1.3, title,
            ha='center', va='center', fontsize=11.5, fontweight='bold',
            color=C_PROBLEM_E, zorder=4)
    ax.text(x + pw/2, prob_y + prob_h/2 - 0.8, sub,
            ha='center', va='center', fontsize=9.5, color='#515A5A', zorder=4)
    prob_centers.append((x + pw/2, prob_y))

# =================================================================
# 第 2 层：目标层 —— 八位一体成果总目标                 y ≈ 76—82
# =================================================================
layer_label(79, '目标层\n总目标', C_GOAL_E)

goal_y, goal_h = 76, 6.0
ax.add_patch(FancyBboxPatch((7, goal_y), 86, goal_h,
                            boxstyle="round,pad=0.02,rounding_size=0.6",
                            facecolor=C_GOAL, edgecolor=C_GOAL_E, linewidth=1.8, zorder=3))
ax.text(50, goal_y + goal_h - 1.3,
        '建成一门以 "三图谱 + 双智能体" 为核心特征的专业核心课程样板',
        ha='center', va='center', fontsize=12.5, fontweight='bold',
        color=C_GOAL_E, zorder=4)

tags = ['一标准', '三图谱', '两智能体', '一靶场', '一模式', '一评价']
tag_y = goal_y + 1.9
tag_w = 11.5
total_w = len(tags) * tag_w + (len(tags)-1)*1.2
tx0 = 50 - total_w/2
for i, t in enumerate(tags):
    xi = tx0 + i*(tag_w + 1.2)
    ax.add_patch(FancyBboxPatch((xi, tag_y), tag_w, 2.2,
                                boxstyle="round,pad=0.02,rounding_size=0.3",
                                facecolor='white', edgecolor=C_GOAL_E, linewidth=1.2, zorder=4))
    ax.text(xi + tag_w/2, tag_y + 1.1, t, ha='center', va='center',
            fontsize=10, fontweight='bold', color=C_GOAL_E, zorder=5)

# 三重断裂 -> 总目标 箭头（汇聚）
for cx, cy in prob_centers:
    arrow(cx, cy, 50, goal_y + goal_h, color=C_PROBLEM_E, lw=1.2, mutation=13,
          connectionstyle="arc3,rad=0.0")

# =================================================================
# 第 3 层：内容层 —— 三大研究内容(与关键问题一一对应)     y ≈ 53—73
# =================================================================
layer_label(63, '研究内容\n三条主线', '#2874A6')

rc_y, rc_h = 53.5, 19.5
rc_w = 27
rc_gap = (100 - 7 - 3*rc_w) / 2
rc_x0 = 7

rc_data = [
    {
        'color_bg': C_RQ1, 'color_ec': C_RQ1_E,
        'title': '① "三图谱" 构建与\n课程标准重构',
        'subtitle': '对应：图—学断裂 / KQ①',
        'items': [
            '· 知识图谱（12模块·86知识点）',
            '· 能力图谱（127能力点）',
            '· 问题图谱（2300+提问·400+错误）',
            '· 三图谱融合 + Neo4j 建模',
            '· 课程标准修订（3模块·5项目·12任务）',
        ],
    },
    {
        'color_bg': C_RQ2, 'color_ec': C_RQ2_E,
        'title': '② "双智能体" 研发与\n场景嵌入',
        'subtitle': '对应：AI—用断裂 / KQ②',
        'items': [
            '· 私有化 LLM + RAG + LoRA 微调',
            '· 智学小助（代码伴学·路径推荐·归因）',
            '· AI 合规官（法规嵌入·预警·审查）',
            '· 仿真靶场（电商/招聘/政务/工业 4 类 · 10 级反爬）',
            '· 红蓝对抗实战模式',
        ],
    },
    {
        'color_bg': C_RQ3, 'color_ec': C_RQ3_E,
        'title': '③ "双师+双工单" 模式与\n五维评价·迁移推广',
        'subtitle': '对应：教—评断裂 / KQ③',
        'items': [
            '· "教师+AI" 双师 / "教学+企业" 双工单',
            '· 五维评价：代码·调试·协作·合规·伦理反思',
            '· 伦理 20% · 重大违规一票否决',
            '· 跨课程迁移指南 + 师资培训包',
            '· 学习分析与精准补偿',
        ],
    },
]

rc_centers = []
for i, d in enumerate(rc_data):
    x = rc_x0 + i*(rc_w + rc_gap)
    ax.add_patch(FancyBboxPatch((x, rc_y), rc_w, rc_h,
                                boxstyle="round,pad=0.02,rounding_size=0.7",
                                facecolor=d['color_bg'], edgecolor=d['color_ec'],
                                linewidth=1.8, zorder=3))
    # 标题
    ax.text(x + rc_w/2, rc_y + rc_h - 2.0, d['title'],
            ha='center', va='center', fontsize=11.3, fontweight='bold',
            color=d['color_ec'], zorder=4)
    # 对应关系
    ax.text(x + rc_w/2, rc_y + rc_h - 4.5, d['subtitle'],
            ha='center', va='center', fontsize=9.2, color='#566573',
            style='italic', zorder=4)
    # 分割线
    ax.add_patch(plt.Rectangle((x+2, rc_y + rc_h - 5.5), rc_w-4, 0.05,
                               facecolor=d['color_ec'], edgecolor='none',
                               alpha=0.5, zorder=4))
    # 子项
    for j, item in enumerate(d['items']):
        ax.text(x + 1.5, rc_y + rc_h - 7.0 - j*2.3, item,
                ha='left', va='center', fontsize=9.4, color='#1B2631', zorder=4)
    rc_centers.append((x + rc_w/2, rc_y + rc_h))

# 问题 -> 内容 一一对应（彩色细箭头）
for i in range(3):
    cx_top, cy_top = prob_centers[i]
    cx_bot, cy_bot = rc_centers[i]
    colors = [C_PROBLEM_E, C_RQ1_E, C_RQ2_E, C_RQ3_E]
    arrow(cx_top, prob_y, cx_bot, cy_bot,
          color=[C_RQ1_E, C_RQ2_E, C_RQ3_E][i], lw=1.2, mutation=13,
          style='-|>', connectionstyle="arc3,rad=0.0")

# 三模块底部汇聚箭头 -> 技术路径
merge_y = rc_y  # 底部
for i, (cx, _) in enumerate(rc_centers):
    arrow(cx, merge_y, cx, merge_y - 3.0,
          color=[C_RQ1_E, C_RQ2_E, C_RQ3_E][i], lw=1.3, mutation=14,
          style='-|>')

# =================================================================
# 第 4 层：技术路径层 —— 横向流程                         y ≈ 35—48
# =================================================================
layer_label(41, '实施路径\n技术路线', C_PATH_E)

path_y, path_h = 39.5, 5.5
path_nodes = [
    '①\n企业跟岗挖掘',
    '②\n三图谱建模\n(Neo4j)',
    '③\n资源与课程\n标准开发',
    '④\n双智能体\n训练部署',
    '⑤\n仿真靶场+\n红蓝对抗平台',
    '⑥\n多班级试点+\n跨校联合试用',
    '⑦\n学习分析与\n模式评价迭代',
    '⑧\n成果凝练\n推广结题',
]
pn_w = 10.2
pn_gap = (100 - 7 - len(path_nodes)*pn_w) / (len(path_nodes)-1)
pn_x0 = 7
pn_centers = []
for i, t in enumerate(path_nodes):
    x = pn_x0 + i*(pn_w + pn_gap)
    ax.add_patch(FancyBboxPatch((x, path_y), pn_w, path_h,
                                boxstyle="round,pad=0.02,rounding_size=0.5",
                                facecolor=C_PATH, edgecolor=C_PATH_E,
                                linewidth=1.5, zorder=3))
    ax.text(x + pn_w/2, path_y + path_h/2, t,
            ha='center', va='center', fontsize=9.4, color=C_PATH_E,
            fontweight='bold', zorder=4)
    pn_centers.append((x + pn_w/2, path_y, x + pn_w, path_y + path_h/2, x, path_y + path_h/2))

# 横向箭头串联
for i in range(len(path_nodes)-1):
    arrow(pn_centers[i][2]+0.3, pn_centers[i][3],
          pn_centers[i+1][4]-0.3, pn_centers[i+1][3],
          color=C_PATH_E, lw=1.6, mutation=15)

# 迭代回路
arrow(pn_centers[6][0], path_y, pn_centers[3][0], path_y,
      color='#D68910', lw=1.3, mutation=13, style='-|>',
      connectionstyle="arc3,rad=-0.35")
ax.text((pn_centers[3][0]+pn_centers[6][0])/2, path_y - 3.0,
        '↻ 迭代优化', ha='center', va='center',
        fontsize=9.2, color='#D68910', fontweight='bold', style='italic')

# =================================================================
# 第 5 层：成果层 —— 八位一体交付物                       y ≈ 16—28
# =================================================================
layer_label(22, '成果层\n八位一体', C_OUT_E)

out_y, out_h = 18, 8
out_items = [
    ('研究报告', '≥5千字'),
    ('课程标准', '2026版'),
    ('知识+能力\n+问题 三图谱', '3套'),
    ('智学小助\n智能体', '私有化'),
    ('AI 合规官\n智能体', '私有化'),
    ('仿真靶场+\n红蓝对抗', '1套'),
    ('教学数据集', '1套'),
    ('数字资源包', '≥300件'),
    ('教研论文', '≥1篇'),
    ('迁移指南+\n培训包', '≥1项'),
]
on_w = 8.3
on_gap = (100 - 7 - len(out_items)*on_w) / (len(out_items)-1)
on_x0 = 7
for i, (t, s) in enumerate(out_items):
    x = on_x0 + i*(on_w + on_gap)
    ax.add_patch(FancyBboxPatch((x, out_y), on_w, out_h,
                                boxstyle="round,pad=0.02,rounding_size=0.4",
                                facecolor=C_OUT, edgecolor=C_OUT_E,
                                linewidth=1.4, zorder=3))
    ax.text(x + on_w/2, out_y + out_h - 2.3, t,
            ha='center', va='center', fontsize=9.2, fontweight='bold',
            color='#7D6608', zorder=4)
    ax.text(x + on_w/2, out_y + 2.0, s,
            ha='center', va='center', fontsize=9, color=C_OUT_E,
            style='italic', zorder=4)

# 技术路径 -> 成果层 向下箭头
for cx, py, _, _, _, _ in pn_centers[4:]:  # 后几个节点
    pass
# 统一一条向下箭头：从技术路径中段 -> 成果层中心
arrow(50, path_y, 50, out_y + out_h, color=C_PATH_E, lw=1.8, mutation=18)

# =================================================================
# 时间轴                                                y ≈ 6—12
# =================================================================
tl_y = 9
ax.add_patch(plt.Rectangle((7, tl_y-0.15), 86, 0.3, facecolor='#34495E', zorder=3))

phases = [
    ('2026.6—7', '第一阶段\n跟岗挖掘·三图谱1.0', 14),
    ('2026.8—9', '第二阶段\n双智能体·靶场·资源', 36),
    ('2026.10—11', '第三阶段\n多班级试点·跨校试用', 60),
    ('2026.12', '第四阶段\n凝练·论文·推广·结题', 84),
]
for date, label, cx in phases:
    ax.plot(cx, tl_y, 'o', markersize=14, color='white',
            markeredgecolor='#34495E', markeredgewidth=2, zorder=4)
    ax.text(cx, tl_y + 2.8, date, ha='center', va='center',
            fontsize=10, fontweight='bold', color='#1B2631')
    ax.text(cx, tl_y - 3.6, label, ha='center', va='center',
            fontsize=9.3, color='#34495E')

# 时间轴标签
ax.text(3.3, tl_y, '时间轴', ha='center', va='center',
        fontsize=10.5, fontweight='bold', color='white',
        bbox=dict(boxstyle='round,pad=0.4', facecolor='#34495E', edgecolor='none'))

# =================================================================
# 底部注脚（图例）
# =================================================================
legend_elems = [
    mpatches.Patch(facecolor=C_PROBLEM, edgecolor=C_PROBLEM_E, label='问题层·三重断裂'),
    mpatches.Patch(facecolor=C_GOAL,    edgecolor=C_GOAL_E,    label='目标层·八位一体'),
    mpatches.Patch(facecolor=C_RQ1,     edgecolor=C_RQ1_E,    label='研究内容①·三图谱'),
    mpatches.Patch(facecolor=C_RQ2,     edgecolor=C_RQ2_E,    label='研究内容②·双智能体'),
    mpatches.Patch(facecolor=C_RQ3,     edgecolor=C_RQ3_E,    label='研究内容③·模式与评价'),
    mpatches.Patch(facecolor=C_PATH,    edgecolor=C_PATH_E,    label='实施路径·8 步闭环'),
    mpatches.Patch(facecolor=C_OUT,     edgecolor=C_OUT_E,     label='成果层·10 项交付'),
]
ax.legend(handles=legend_elems, loc='lower center', bbox_to_anchor=(0.5, -0.005),
          ncol=7, frameon=False, fontsize=9.2, handlelength=1.2, columnspacing=1.5)

plt.subplots_adjust(left=0.01, right=0.99, top=0.98, bottom=0.04)
out_path = '/workspace/technical_roadmap.png'
plt.savefig(out_path, dpi=200, bbox_inches='tight', facecolor='white')
out_path_pdf = '/workspace/technical_roadmap.pdf'
plt.savefig(out_path_pdf, bbox_inches='tight', facecolor='white')
print('Saved:', out_path)
print('Saved:', out_path_pdf)
