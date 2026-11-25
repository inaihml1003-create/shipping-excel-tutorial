import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import pandas as pd
import numpy as np

plt.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial']
plt.rcParams['axes.unicode_minus'] = False

os.makedirs("images", exist_ok=True)
print("Rendering polished, high-fidelity UI and workflow screenshots...")

# -----------------------------------------------------------------
# Fig 1-1: Excel Full UI Layout
# -----------------------------------------------------------------
fig, ax = plt.subplots(figsize=(12, 7.5), dpi=200)
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

# Outer frame
ax.add_patch(patches.Rectangle((1, 1), 98, 98, linewidth=1.5, edgecolor='#1E4E79', facecolor='#F3F3F3'))

# Title Bar
ax.add_patch(patches.Rectangle((1, 93), 98, 6, facecolor='#107C41'))
ax.text(3, 96, "快速存取工具列  [ 储存 | 复原 | 重做 ]", color='white', fontsize=10, fontweight='bold', va='center')
ax.text(50, 96, "航运实务资料前处理教学_示范档.xlsx - Excel", color='white', fontsize=11, fontweight='bold', ha='center', va='center')
ax.text(95, 96, "[ 最小化 | 还原 | 关闭 ]", color='white', fontsize=8.5, ha='center', va='center')

# Ribbon Tabs
ax.add_patch(patches.Rectangle((1, 88), 98, 5, facecolor='#E1DFDD', edgecolor='#D2D0CE'))
tabs = ["档案", "常用", "插入", "版面配置", "公式", "资料 (Data) [核心]", "校阅", "检视", "说明"]
x_t = 3
for t in tabs:
    is_active = "资料" in t
    w_t = len(t) * 1.6 + 3
    ax.add_patch(patches.Rectangle((x_t, 88), w_t, 5, facecolor='#FFFFFF' if is_active else '#E1DFDD', edgecolor='#C8C6C4'))
    ax.text(x_t + w_t/2, 90.5, t, color='#107C41' if is_active else '#323130', fontsize=9, fontweight='bold' if is_active else 'normal', ha='center', va='center')
    x_t += w_t + 1

# Ribbon Body
ax.add_patch(patches.Rectangle((1, 73), 98, 15, facecolor='#FFFFFF', edgecolor='#C8C6C4'))
groups = [
    ("取得及转换资料", ["从文字/CSV", "从Web", "从表格/范围"]),
    ("查询与连线", ["重新整理全部", "查询与连线"]),
    ("排序与筛选", ["A-Z 递增排序", "Z-A 递减排序", "筛选 (Filter)"]),
    ("资料工具 [核心前处理功能群]", ["资料剖析 (Text to Columns)", "移除重复项 (Remove Duplicates)", "快速填入 (Ctrl+E)", "资料验证 (Data Validation)"]),
    ("预测与大纲", ["模拟分析", "建立群组"])
]
x_g = 2
for g_title, g_items in groups:
    is_core = "核心" in g_title
    w_g = 36 if is_core else 15
    ax.add_patch(patches.Rectangle((x_g, 74), w_g, 13.5, facecolor='#F3F9F4' if is_core else '#FAFAFA', edgecolor='#107C41' if is_core else '#EDEBE9', linewidth=1.5 if is_core else 1))
    y_it = 84
    for it in g_items:
        ax.text(x_g + 1.2, y_it, f"- {it}", fontsize=8, color='#1E4E79' if is_core else '#323130', fontweight='bold' if is_core else 'normal')
        y_it -= 2.6
    ax.text(x_g + w_g/2, 74.8, g_title, fontsize=8, color='#107C41' if is_core else '#605E5C', ha='center', va='bottom', fontweight='bold')
    x_g += w_g + 0.8

# Formula Bar & Name Box
ax.add_patch(patches.Rectangle((1, 67), 98, 5.5, facecolor='#F3F2F1', edgecolor='#C8C6C4'))
ax.add_patch(patches.Rectangle((3, 68), 11, 3.5, facecolor='white', edgecolor='#8A8886'))
ax.text(8.5, 69.75, "C4", fontsize=10, fontweight='bold', ha='center', va='center')
ax.text(16, 69.75, "fx", fontsize=11, fontweight='bold', color='#605E5C', ha='center', va='center')
ax.add_patch(patches.Rectangle((19, 68), 78, 3.5, facecolor='white', edgecolor='#8A8886'))
ax.text(21, 69.75, '=TRIM(CLEAN(UPPER(B4)))   // 示范: 清除不可见字符与首尾空格并转换为大写', fontsize=9, color='#004E8C', fontfamily='Microsoft YaHei', va='center')

# Grid
ax.add_patch(patches.Rectangle((1, 7), 98, 59, facecolor='white', edgecolor='#C8C6C4'))
col_names = ["", "A (序号)", "B (提单号B/L)", "C (清洗后提单)", "D (箱号Container)", "E (箱型)", "F (起运港POL)", "G (目的港POD)", "H (海运费USD)"]
col_widths = [5, 9, 15, 16, 15, 8, 12, 12, 12]
x_c = 1
for idx, (c_name, c_w) in enumerate(zip(col_names, col_widths)):
    ax.add_patch(patches.Rectangle((x_c, 62), c_w, 4.5, facecolor='#E1DFDD', edgecolor='#C8C6C4'))
    ax.text(x_c + c_w/2, 64.2, c_name, fontsize=8, fontweight='bold', ha='center', va='center')
    x_c += c_w

grid_data = [
    ("1", "1", " COSU62918401 ", "COSU62918401", "cosu9823145", "40GP", "CNSHA", "DEHAM", "$2,450.00"),
    ("2", "2", "COSU62918402", "COSU62918402", "TEMU4819201", "20GP", "CNSHA", "NLSRT", "$1,200.00"),
    ("3", "3", "MSCU91823711", "MSCU91823711", "MSCU1928374", "40HC", "CNNBO", "SGSIN", "$1,800.00"),
    ("4", "4", "COSU62918402 [重复]", "COSU62918402", "TEMU4819201", "20GP", "CNSHA", "NLSRT", "$1,200.00"),
    ("5", "5", " OOCL82910394", "OOCL82910394", "OOCU9283019", "40HC", "CNSHA", "USLAX", "$3,100.00"),
    ("6", "6", "CMAU28190283 ", "CMAU28190283", "CMAU8172930", "20GP", "CNNBO", "FRLEH", "$1,650.00"),
    ("7", "7", "ONEY82910283", "ONEY82910283", "TGHU9182039", "40GP", "CNSHA", "USOAK", "$2,950.00"),
    ("8", "8", " EMCU92810394 ", "EMCU92810394", "EGLU1928304", "40HC", "CNNBO", "GRPIR", "$2,100.00"),
    ("9", "9", "HLCU82910384", "HLCU82910384", "HLXU9182736", "20GP", "CNDLC", "USNYC", "$2,200.00"),
    ("10", "10", "COSU62918403", "COSU62918403", "COSU1029384", "40HC", "CNSHA", "BEANR", "$2,350.00"),
]

y_r = 57.5
for row_tuple in grid_data:
    x_c = 1
    is_dup = "重复" in row_tuple[2]
    for idx, (val, c_w) in enumerate(zip(row_tuple, col_widths)):
        bg = '#FFE2E2' if (is_dup and idx > 0) else ('#E1DFDD' if idx == 0 else ('#EDF5E1' if idx == 3 else 'white'))
        ax.add_patch(patches.Rectangle((x_c, y_r), c_w, 4.5, facecolor=bg, edgecolor='#E1DFDD'))
        align = 'center' if idx in [0, 1, 5, 6, 7] else ('right' if idx == 8 else 'left')
        ax.text(x_c + (c_w/2 if align == 'center' else (c_w - 1.5 if align == 'right' else 1.5)), 
                y_r + 2.2, val, fontsize=8, va='center', ha=align, 
                color='#C00000' if is_dup else '#323130', fontweight='bold' if (idx == 3 or is_dup) else 'normal')
        x_c += c_w
    y_r -= 4.6

# Sheet tab & status bar
ax.add_patch(patches.Rectangle((1, 1), 98, 6, facecolor='#F3F2F1', edgecolor='#C8C6C4'))
sheet_names = ["30笔原始脏资料", "30笔清洗后标准资料", "基础代码参照表"]
x_st = 4
for s_idx, sn in enumerate(sheet_names):
    is_cur = s_idx == 0
    w_st = len(sn) * 1.6 + 2
    ax.add_patch(patches.Rectangle((x_st, 2), w_st, 4, facecolor='white' if is_cur else '#E1DFDD', edgecolor='#C8C6C4'))
    ax.text(x_st + w_st/2, 4, sn, fontsize=8, fontweight='bold' if is_cur else 'normal', color='#107C41' if is_cur else '#323130', ha='center', va='center')
    x_st += w_st + 1

ax.text(78, 4, "就绪 (Ready)   |   选取计数: 30   |   100% [ ---o--- ]", fontsize=8, color='#605E5C', va='center')

# Annotations
def draw_callout(x_start, y_start, x_text, y_text, text, color='#C00000'):
    ax.annotate(text, xy=(x_start, y_start), xytext=(x_text, y_text),
                arrowprops=dict(facecolor=color, edgecolor=color, arrowstyle='->', lw=1.5),
                bbox=dict(boxstyle='round,pad=0.3', facecolor='#FFF2CC', edgecolor=color, lw=1),
                fontsize=8, fontweight='bold', color='#B25900')

draw_callout(12, 91, 15, 98.5, "功能区标签页 (Ribbon Tabs)\n切换至「资料」选项卡", '#0070C0')
draw_callout(42, 80, 50, 85, "资料工具核心按键群:\n资料剖析、移除重复项、快速填入", '#107C41')
draw_callout(8.5, 69.75, 4, 75, "名称框 (Name Box)\n显示当前单元格座标", '#7030A0')
draw_callout(30, 69.75, 40, 75, "资料编辑列 (Formula Bar)\n检视真实公式与底层数据", '#004E8C')
draw_callout(20, 42, 28, 44, "异常记录 (首尾空格、重复录入)", '#C00000')

plt.tight_layout()
plt.savefig("images/fig_1_1_excel_interface_overview.png", dpi=200, bbox_inches='tight')
plt.close()


# -----------------------------------------------------------------
# Fig 2-1: Tidy Data vs Messy Data
# -----------------------------------------------------------------
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5.5), dpi=200)
for ax in [ax1, ax2]:
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

# Messy Data
ax1.set_title("[X] 常见航运混乱资料结构 (Messy Data / 难以机器分析)", fontsize=11, fontweight='bold', color='#C00000', pad=10)
ax1.add_patch(patches.Rectangle((2, 2), 96, 94, facecolor='#FFF7F7', edgecolor='#C00000', linewidth=1.5))
ax1.add_patch(patches.Rectangle((5, 78), 90, 10, facecolor='#E7E6E6', edgecolor='#A6A6A6'))
ax1.add_patch(patches.Rectangle((5, 62), 40, 14, facecolor='#FFD9D9', edgecolor='#C00000'))
ax1.add_patch(patches.Rectangle((47, 62), 48, 14, facecolor='#FFFFFF', edgecolor='#A6A6A6'))
ax1.add_patch(patches.Rectangle((5, 46), 90, 14, facecolor='#FFE699', edgecolor='#B25900'))
ax1.add_patch(patches.Rectangle((5, 30), 90, 14, facecolor='#FFFFFF', edgecolor='#A6A6A6'))
ax1.add_patch(patches.Rectangle((5, 14), 90, 14, facecolor='#E2EFDA', edgecolor='#385723'))

ax1.text(50, 83, "大洋物流2026年11月船期与海运费汇总表 (多行合并大表头)", fontsize=9, fontweight='bold', ha='center', color='#323130')
ax1.text(25, 69, "[警示] 跨行跨列合并\n(上海港直航组)", fontsize=8.5, ha='center', va='center', color='#C00000', fontweight='bold')
ax1.text(71, 69, "柜号/箱型挤在同格: TEMU192830(40HC)\n运费: $2300 含打托费", fontsize=7.5, ha='center', va='center')
ax1.text(50, 53, "[禁止] 小计行: 本港小计 12 柜 / $28,400 (严重破坏资料库连续性)", fontsize=8, ha='center', va='center', color='#B25900', fontweight='bold')
ax1.text(50, 37, "空格与不可见符号 / 日期写成 '11月5号左右' / 空白行隔开记录", fontsize=8, ha='center', va='center', color='#7030A0')
ax1.text(50, 21, "[禁止] 总计行: 统计总额 $158,000 (混入资料区域，导致枢纽分析表重复计算)", fontsize=8, ha='center', va='center', color='#385723', fontweight='bold')

# Tidy Data
ax2.set_title("[V] 标准整洁航运资料结构 (Tidy Data / 适于自动化分析)", fontsize=11, fontweight='bold', color='#107C41', pad=10)
ax2.add_patch(patches.Rectangle((2, 2), 96, 94, facecolor='#F4FBF7', edgecolor='#107C41', linewidth=1.5))
headers_tidy = ["提单号", "柜号", "箱型", "POL", "POD", "ETD", "海运费(USD)"]
w_tidy = 90 / len(headers_tidy)
x_th = 5
for h in headers_tidy:
    ax2.add_patch(patches.Rectangle((x_th, 80), w_tidy, 10, facecolor='#274E13', edgecolor='#FFFFFF'))
    ax2.text(x_th + w_tidy/2, 85, h, color='white', fontsize=8, fontweight='bold', ha='center', va='center')
    x_th += w_tidy

rows_tidy_sample = [
    ("COSU6291", "COSU1029", "40GP", "CNSHA", "DEHAM", "2026-11-05", "2450.00"),
    ("COSU6292", "TEMU4819", "20GP", "CNSHA", "NLSRT", "2026-11-06", "1200.00"),
    ("MSCU9182", "MSCU1928", "40HC", "CNNBO", "SGSIN", "2026-11-07", "1800.00"),
    ("OOCL8291", "OOCU9283", "40HC", "CNSHA", "USLAX", "2026-11-08", "3100.00"),
    ("CMAU2819", "CMAU8172", "20GP", "CNNBO", "FRLEH", "2026-11-09", "1650.00"),
]
y_tr = 70
for r_data in rows_tidy_sample:
    x_th = 5
    for c_val in r_data:
        ax2.add_patch(patches.Rectangle((x_th, y_tr), w_tidy, 8, facecolor='white', edgecolor='#D9D9D9'))
        ax2.text(x_th + w_tidy/2, y_tr + 4, c_val, fontsize=7.5, ha='center', va='center', color='#323130')
        x_th += w_tidy
    y_tr -= 8.5

ax2.add_patch(patches.Rectangle((5, 5), 90, 22, facecolor='#E2EFDA', edgecolor='#385723', linewidth=1))
ax2.text(50, 22, "* 整洁数据三原则 (Codd & Wickham Tidy Data Rules):", fontsize=8.5, fontweight='bold', ha='center', color='#274E13')
ax2.text(8, 16, "1. 每一列 (Column) 代表唯一的业务变量 (如纯粹的箱型、纯粹的港口代码)。", fontsize=7.5, color='#323130')
ax2.text(8, 11, "2. 每一行 (Row) 代表单次具体的运输观测纪录 (单一货柜或单一提单业务)。", fontsize=7.5, color='#323130')
ax2.text(8, 6, "3. 无跨行合并格、无夹杂小计合计行、无复合多重含义栏位、数值纯净无文字单位。", fontsize=7.5, color='#323130')

plt.tight_layout()
plt.savefig("images/fig_2_1_tidy_vs_messy_data.png", dpi=200, bbox_inches='tight')
plt.close()


# -----------------------------------------------------------------
# Fig 4-1: Text to Columns Wizard
# -----------------------------------------------------------------
fig, ax = plt.subplots(figsize=(10, 6), dpi=200)
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

ax.add_patch(patches.Rectangle((5, 5), 90, 90, facecolor='#F3F3F3', edgecolor='#0078D7', linewidth=2))
ax.add_patch(patches.Rectangle((5, 87), 90, 8, facecolor='#0078D7'))
ax.text(8, 91, "资料剖析精灵 - 步骤 3 之 2 (分界符号与复合栏位拆分)", color='white', fontsize=11, fontweight='bold', va='center')

ax.text(10, 81, "【应用情境】将「船名航次 (EVER GIVEN / 0118E)」依照斜杠 '/' 拆分为「船名」与「航次」两栏", fontsize=9.5, fontweight='bold', color='#1E4E79')

ax.add_patch(patches.Rectangle((10, 48), 38, 28, facecolor='white', edgecolor='#A6A6A6'))
ax.text(12, 72, "请勾选栏位分界符号 (Delimiters):", fontsize=8.5, fontweight='bold')
ax.text(15, 66, "[ ] Tab 键 (T)", fontsize=8)
ax.text(15, 61, "[ ] 分号 (S)", fontsize=8)
ax.text(15, 56, "[ ] 逗号 (C)", fontsize=8)
ax.text(15, 51, "[V] 其他 (O): [  /  ]  <-- 输入航运常用斜杠", fontsize=8.5, fontweight='bold', color='#C00000')

ax.add_patch(patches.Rectangle((52, 48), 38, 28, facecolor='white', edgecolor='#A6A6A6'))
ax.text(54, 72, "处理选项设置:", fontsize=8.5, fontweight='bold')
ax.text(56, 64, "[V] 将连续分界符号视为单一符号", fontsize=8)
ax.text(56, 56, "文字识别符号: [  \"  ]", fontsize=8)

ax.add_patch(patches.Rectangle((10, 15), 80, 29, facecolor='white', edgecolor='#A6A6A6'))
ax.text(12, 40, "资料预览 (Data Preview):", fontsize=8.5, fontweight='bold')

ax.add_patch(patches.Rectangle((12, 17), 38, 20, facecolor='#EDF5E1', edgecolor='#8A8886'))
ax.add_patch(patches.Rectangle((51, 17), 37, 20, facecolor='#E2EFDA', edgecolor='#8A8886'))

ax.text(31, 33, "栏位 1 (拆出船名)", fontsize=8, fontweight='bold', ha='center', color='#274E13')
ax.text(31, 28, "EVER GIVEN", fontsize=8, ha='center')
ax.text(31, 23, "COSCO GUANGZHOU", fontsize=8, ha='center')
ax.text(31, 18, "MSC OSCAR", fontsize=8, ha='center')

ax.text(69, 33, "栏位 2 (拆出航次代码)", fontsize=8, fontweight='bold', ha='center', color='#274E13')
ax.text(69, 28, "0118E", fontsize=8, ha='center')
ax.text(69, 23, "042W", fontsize=8, ha='center')
ax.text(69, 18, "2603W", fontsize=8, ha='center')

ax.text(62, 9, "[ < 上一步(B) ]   [ 下一步(N) > ]   [ 完成(F) ]   [ 取消 ]", fontsize=9, fontweight='bold', color='#323130')

plt.tight_layout()
plt.savefig("images/fig_4_1_text_to_columns.png", dpi=200, bbox_inches='tight')
plt.close()


# -----------------------------------------------------------------
# Fig 4-2: Remove Duplicates & Validation
# -----------------------------------------------------------------
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 6), dpi=200)

ax1.set_xlim(0, 100)
ax1.set_ylim(0, 100)
ax1.axis('off')

ax1.add_patch(patches.Rectangle((5, 5), 90, 90, facecolor='#F3F3F3', edgecolor='#0078D7', linewidth=2))
ax1.add_patch(patches.Rectangle((5, 87), 90, 8, facecolor='#0078D7'))
ax1.text(8, 91, "移除重复项 (Remove Duplicates)", color='white', fontsize=10.5, fontweight='bold', va='center')

ax1.text(8, 80, "若要删除重复值，请选定包含重复项的栏位:", fontsize=8.5)
ax1.text(8, 73, "[V] 我的资料具有标题 (My data has headers)", fontsize=8.5, fontweight='bold', color='#107C41')

ax1.add_patch(patches.Rectangle((8, 22), 84, 48, facecolor='white', edgecolor='#A6A6A6'))
cols_dup = [
    ("[ ] (全选)", "取消全选"),
    ("[V] 提单号 (B/L No.)", "* 关键唯一主键"),
    ("[V] 集装箱号 (Container No.)", "* 货柜物理唯一标识"),
    ("[ ] 箱型 (Type)", ""),
    ("[ ] 船名航次", ""),
    ("[ ] 装运港 / 目的港", "")
]
y_cd = 63
for c_name, c_hint in cols_dup:
    is_key = "*" in c_hint
    ax1.text(12, y_cd, c_name, fontsize=8, fontweight='bold' if is_key else 'normal', color='#C00000' if is_key else '#323130')
    if c_hint:
        ax1.text(55, y_cd, c_hint, fontsize=7.5, color='#B25900' if is_key else '#605E5C', fontweight='bold' if is_key else 'normal')
    y_cd -= 7.5

ax1.text(8, 14, "[要诀] 航运实务要诀: 依「提单号+柜号」作为判别基准，\n可精准清除重复录入，保留首次录入的合法单证。", fontsize=7.5, color='#1E4E79')
ax1.text(65, 8, "[ 确定 (OK) ]   [ 取消 ]", fontsize=8.5, fontweight='bold')

ax2.set_xlim(0, 100)
ax2.set_ylim(0, 100)
ax2.axis('off')

ax2.add_patch(patches.Rectangle((5, 5), 90, 90, facecolor='#F3F3F3', edgecolor='#107C41', linewidth=2))
ax2.add_patch(patches.Rectangle((5, 87), 90, 8, facecolor='#107C41'))
ax2.text(8, 91, "资料验证 (Data Validation) - 规范源头录入", color='white', fontsize=10.5, fontweight='bold', va='center')

ax2.text(10, 81, "[设定 (Settings)]   [输入讯息]   [错误警讯]", fontsize=8.5, fontweight='bold', color='#107C41')

ax2.add_patch(patches.Rectangle((8, 22), 84, 55, facecolor='white', edgecolor='#A6A6A6'))
ax2.text(12, 70, "验证准则 (Validation criteria):", fontsize=8.5, fontweight='bold')
ax2.text(12, 62, "允许 (Allow):  [ 清单 (List)                 [V] ]", fontsize=8.5, fontweight='bold', color='#004E8C')
ax2.text(12, 53, "[V] 忽略空白 (Ignore blank)", fontsize=8)
ax2.text(12, 46, "[V] 储存格内提供下拉式清单 (In-cell dropdown)", fontsize=8, fontweight='bold', color='#107C41')

ax2.text(12, 36, "来源 (Source):", fontsize=8.5, fontweight='bold')
ax2.add_patch(patches.Rectangle((12, 26), 76, 7, facecolor='#F2F2F2', edgecolor='#7F7F7F'))
ax2.text(14, 29.5, "=基础代码参照表!$B$2:$B$19  (或 20GP,40GP,40HC)", fontsize=8, fontfamily='Microsoft YaHei', color='#C00000', fontweight='bold')

ax2.text(8, 14, "[要诀] 航运实务要诀: 锁定港口代码与箱型选项，\n从源头杜绝 CNSHA 误打成 cn-sha、sha 等混乱资料。", fontsize=7.5, color='#1E4E79')
ax2.text(65, 8, "[ 确定 (OK) ]   [ 取消 ]", fontsize=8.5, fontweight='bold')

plt.tight_layout()
plt.savefig("images/fig_4_2_remove_duplicates_and_validation.png", dpi=200, bbox_inches='tight')
plt.close()


# -----------------------------------------------------------------
# Fig 4-3: Formula Cleaning Showcase
# -----------------------------------------------------------------
fig, ax = plt.subplots(figsize=(11, 5.5), dpi=200)
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

ax.add_patch(patches.Rectangle((2, 2), 96, 96, facecolor='#FFFFFF', edgecolor='#1F497D', linewidth=2))
ax.add_patch(patches.Rectangle((2, 88), 96, 10, facecolor='#1F497D'))
ax.text(50, 93, "Excel 航运数据清洗核心函数公式联动语法示范", color='white', fontsize=11, fontweight='bold', ha='center', va='center')

f_cards = [
    ("1. 强力清除不可见字符与空格", "=TRIM(CLEAN(B2))", "清除网页/EDI系统中常见的首尾多余空格与不可见ASCII控制字符 (0-31)。", "#EBF1F5", "#1F497D"),
    ("2. 文本大小写规范与提取", '=UPPER(TRIM(C2))  与  =LEFT(D2, 4)', "集装箱箱号前4位为所有者代码与箱型标记（如 COSU, MSCU），统一大写。", "#E2EFDA", "#274E13"),
    ("3. 混杂文本日期强力标准化", '=DATE(LEFT(G2,4), MID(G2,5,2), RIGHT(G2,2))', "将形如 '20261105' 纯文本日期转换为系统真日期，便于后续航程与延误运算。", "#FFF2CC", "#B25900"),
    ("4. 文本数值剥离与金额清洗", '=VALUE(SUBSTITUTE(SUBSTITUTE(H2, "$", ""), ",", ""))', "剔除运费中的美元符号 '$' 与千分位逗号 ','，转化为可进行求和运算的纯数值。", "#FCE4D6", "#C65911"),
    ("5. 智能跨表查询与异常捕获", '=XLOOKUP(E2, 港口代码表!A:A, 港口代码表!B:B, "未知港口")', "自动根据5位 UN/LOCODE 代码比对补全港口标准中文全称，容错处理缺漏港名。", "#F2DCDB", "#A61C1C")
]

y_fc = 72
for title, formula, desc, bg_c, bd_c in f_cards:
    ax.add_patch(patches.Rectangle((5, y_fc), 90, 14, facecolor=bg_c, edgecolor=bd_c, linewidth=1.2))
    ax.text(7, y_fc + 10, title, fontsize=8.5, fontweight='bold', color=bd_c)
    ax.text(48, y_fc + 10, f"公式: {formula}", fontsize=8.5, fontweight='bold', color='#004E8C', fontfamily='Microsoft YaHei')
    ax.text(7, y_fc + 3.5, f"实务说明: {desc}", fontsize=7.5, color='#323130')
    y_fc -= 16.5

plt.tight_layout()
plt.savefig("images/fig_4_3_formulas_cleaning_showcase.png", dpi=200, bbox_inches='tight')
plt.close()


# -----------------------------------------------------------------
# Fig 5-1: Practice Pipeline
# -----------------------------------------------------------------
fig, ax = plt.subplots(figsize=(12, 6), dpi=200)
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

ax.text(50, 95, "航运实务 300 笔情境资料前处理分析全景管线 (Data Preprocessing Pipeline)", fontsize=11.5, fontweight='bold', ha='center', color='#1F497D')

stages = [
    ("第 1 阶段: 结构诊断与去重", "1. 检查复合表头与空行\n2. 校验提单/柜号主键\n3. 执行「移除重复项」\n4. 过滤空工作单号", "#EBF1F5", "#1F497D"),
    ("第 2 阶段: 文本规范与清洗", "1. TRIM 清除前后空格\n2. CLEAN 消除乱码换行\n3. UPPER 统一大写代码\n4. 分列拆解「船名/航次」", "#E2EFDA", "#274E13"),
    ("第 3 阶段: 日期时间标准化", "1. 统一 2026-MM-DD\n2. DATEVALUE 转换纯日期\n3. 计算实际航程天数\n4. 判定 ATA 延误与滞期", "#FFF2CC", "#B25900"),
    ("第 4 阶段: 数值类型与费用换算", "1. 剔除 KGS/CBM 单位字样\n2. VALUE 转换计算数字\n3. 计算计费吨 RT=MAX(..)\n4. 滞期费与总应收核算", "#FCE4D6", "#C65911"),
    ("第 5 阶段: 代码对齐与品质封箱", "1. XLOOKUP 匹配中文港名\n2. 对齐国际船东联盟信息\n3. 设定下拉选单防二次污染\n4. 产出最终标准分析模型", "#F2DCDB", "#A61C1C")
]

x_pos = 2
for idx, (s_title, s_desc, bg_color, border_color) in enumerate(stages):
    ax.add_patch(patches.Rectangle((x_pos, 22), 17.5, 62, facecolor=bg_color, edgecolor=border_color, linewidth=2, linestyle='-'))
    ax.add_patch(patches.Rectangle((x_pos, 72), 17.5, 12, facecolor=border_color))
    ax.text(x_pos + 8.75, 78, s_title, color='white', fontsize=8, fontweight='bold', ha='center', va='center')
    ax.text(x_pos + 1.2, 66, s_desc, fontsize=7.5, color='#323130', va='top', linespacing=1.6)
    if idx < len(stages) - 1:
        ax.annotate('', xy=(x_pos + 19.5, 52), xytext=(x_pos + 17.7, 52),
                    arrowprops=dict(facecolor='#1F497D', edgecolor='#1F497D', width=2, headwidth=6, headlength=6))
    x_pos += 19.6

ax.add_patch(patches.Rectangle((2, 5), 96, 12, facecolor='#1F497D', edgecolor='#1F497D'))
ax.text(50, 11, "[目标达成] 从 ERP/EDI 导出混杂脏资料 -> 蜕变为零错误、标准统一、即开即用的高品质航运资料库", 
        color='white', fontsize=9.5, fontweight='bold', ha='center', va='center')

plt.tight_layout()
plt.savefig("images/fig_5_1_practice_pipeline.png", dpi=200, bbox_inches='tight')
plt.close()

print("Polished diagrams generated with zero warnings!")
