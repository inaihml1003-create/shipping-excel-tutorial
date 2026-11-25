import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_shipping_tutorial_pptx():
    prs = Presentation()
    
    # 16:9 Widescreen standard
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    # Blank slide layout
    blank_layout = prs.slide_layouts[6]
    
    # Color palette
    NAVY_DARK = RGBColor(10, 25, 47)      # #0A192F
    NAVY_CARD = RGBColor(17, 34, 64)      # #112240
    NAVY_BORDER = RGBColor(30, 58, 95)    # #1E3A5F
    TEAL_ACCENT = RGBColor(0, 168, 150)   # #00A896
    TEAL_LIGHT = RGBColor(72, 202, 228)   # #48CAE4
    WHITE = RGBColor(255, 255, 255)
    TEXT_MUTED = RGBColor(148, 163, 184)  # #94A3B8
    TEXT_LIGHT = RGBColor(226, 232, 240)  # #E2E8F0
    RED_ACCENT = RGBColor(231, 111, 81)   # #E76F51
    AMBER_ACCENT = RGBColor(244, 162, 97) # #F4A261
    GREEN_ACCENT = RGBColor(46, 204, 113) # #2ECC71
    
    def set_slide_background(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = NAVY_DARK
        bg.line.fill.background()
        return bg
        
    def add_header(slide, ch_tag, title_text, badge_text=""):
        # Header text box
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.9))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        p_tag = tf.paragraphs[0]
        p_tag.space_after = Pt(2)
        r_tag = p_tag.add_run()
        r_tag.text = ch_tag.upper()
        r_tag.font.name = "Microsoft YaHei"
        r_tag.font.size = Pt(10)
        r_tag.font.bold = True
        r_tag.font.color.rgb = TEAL_ACCENT
        
        p_title = tf.add_paragraph()
        r_title = p_title.add_run()
        r_title.text = title_text
        r_title.font.name = "Microsoft YaHei"
        r_title.font.size = Pt(20)
        r_title.font.bold = True
        r_title.font.color.rgb = WHITE
        
        if badge_text:
            r_sp = p_title.add_run()
            r_sp.text = "   "
            r_badge = p_title.add_run()
            r_badge.text = f"[{badge_text}]"
            r_badge.font.name = "Microsoft YaHei"
            r_badge.font.size = Pt(13)
            r_badge.font.bold = False
            r_badge.font.color.rgb = TEAL_LIGHT
            
    def add_footer(slide, current_idx, total_count=13):
        tb = slide.shapes.add_textbox(Inches(0.8), Inches(6.9), Inches(11.733), Inches(0.4))
        tf = tb.text_frame
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        r1 = p.add_run()
        r1.text = "航运行政人员 Excel 基础资料分析：资料前处理核心技能实务  |  智慧港航数字化教研组"
        r1.font.name = "Microsoft YaHei"
        r1.font.size = Pt(9)
        r1.font.color.rgb = TEXT_MUTED
        
        r2 = p.add_run()
        r2.text = f"                                                                               {current_idx:02d} / {total_count:02d}"
        r2.font.name = "Microsoft YaHei"
        r2.font.size = Pt(9)
        r2.font.bold = True
        r2.font.color.rgb = TEAL_LIGHT

    # =========================================================================
    # SLIDE 1: COVER
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    set_slide_background(s1)
    
    # Large Accent Card
    card1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.0), Inches(11.333), Inches(5.5))
    card1.fill.solid()
    card1.fill.fore_color.rgb = NAVY_CARD
    card1.line.color.rgb = TEAL_ACCENT
    card1.line.width = Pt(1.5)
    
    tb1 = s1.shapes.add_textbox(Inches(1.5), Inches(1.4), Inches(10.333), Inches(4.7))
    tf1 = tb1.text_frame
    tf1.word_wrap = True
    
    p0 = tf1.paragraphs[0]
    p0.space_after = Pt(10)
    r_sub0 = p0.add_run()
    r_sub0.text = "🚢 航运物流数字化运营系列课程 • MS EXCEL 365 / 2021 / 2019"
    r_sub0.font.name = "Microsoft YaHei"
    r_sub0.font.size = Pt(13)
    r_sub0.font.bold = True
    r_sub0.font.color.rgb = TEAL_LIGHT
    
    p_main = tf1.add_paragraph()
    p_main.space_after = Pt(12)
    r_m = p_main.add_run()
    r_m.text = "航运行政人员 Excel 基础资料分析\n资料前处理核心技能全景实战"
    r_m.font.name = "Microsoft YaHei"
    r_m.font.size = Pt(32)
    r_m.font.bold = True
    r_m.font.color.rgb = WHITE
    
    p_desc = tf1.add_paragraph()
    p_desc.space_after = Pt(28)
    r_d = p_desc.add_run()
    r_d.text = "专为航运行政初学者定制：从操作界面、整洁数据（Tidy Data）哲学到 30 笔示范清洗实战与 300 笔多场景情境题解。"
    r_d.font.name = "Microsoft YaHei"
    r_d.font.size = Pt(14)
    r_d.font.color.rgb = TEXT_LIGHT
    
    # 4 Highlights in cover
    highlights = [
        ("🖥️ 操作界面解析", "功能区/编辑列/网格/高频快捷键"),
        ("🧹 前处理基本常识", "GIGO 铁律 / 5大资料类型 / 整洁数据"),
        ("🚢 30 笔实务示范", "去重/分列/格式归一/跨表码表关联"),
        ("🎯 300 笔情境挑战", "集装箱流转/运费滞期/船期准点率题解")
    ]
    for idx, (h_title, h_sub) in enumerate(highlights):
        x = Inches(1.5 + idx * 2.6)
        y = Inches(4.5)
        box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(2.4), Inches(1.4))
        box.fill.solid()
        box.fill.fore_color.rgb = NAVY_DARK
        box.line.color.rgb = NAVY_BORDER
        
        btf = box.text_frame
        btf.word_wrap = True
        bp1 = btf.paragraphs[0]
        br1 = bp1.add_run()
        br1.text = h_title
        br1.font.name = "Microsoft YaHei"
        br1.font.size = Pt(11)
        br1.font.bold = True
        br1.font.color.rgb = WHITE
        
        bp2 = btf.add_paragraph()
        bp2.space_before = Pt(4)
        br2 = bp2.add_run()
        br2.text = h_sub
        br2.font.name = "Microsoft YaHei"
        br2.font.size = Pt(9.5)
        br2.font.color.rgb = TEXT_MUTED

    add_footer(s1, 1)

    # =========================================================================
    # SLIDE 2: AGENDA
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    set_slide_background(s2)
    add_header(s2, "Chapter Navigation", "教学大纲与学习路线图", "6 大核心模组全景图")
    
    modules = [
        ("第 1 章：Excel 介面全景与核心区域", "标题列、快速存取工具列、「资料」功能区群组、编辑列、名称框与初学者必备高频快捷键。", TEAL_ACCENT),
        ("第 2 章：前处理基本常识与整洁数据", "GIGO 垃圾进垃圾出铁律、5 大数据类型判别、Tidy Data 三大黄金原则与 4 类典型脏数据形态。", TEAL_LIGHT),
        ("第 3 章：30 笔实务示范全流程实操", "原始脏数据剖析、体检诊断清册、8 步标准清洗流水线以及清洗前后关键指标对照。", RGBColor(59, 130, 246)),
        ("第 4 章：关键功能选项与核心函数", "资料剖析、去重、验证、快速填入；TRIM, CLEAN, UPPER, DATE, XLOOKUP 与万能嵌套公式。", RGBColor(147, 51, 234)),
        ("第 5 章：300 笔多场景练习与题解", "集装箱流转、海运运费滞期、船期准点率三大业务场景架构，十大挑战任务手把手题解与标准答案。", RGBColor(236, 72, 153)),
        ("第 6 章：品质管控、自动化与资源", "源头表单刚性防错设计、Power Query 自动化展望、航运日常数据处理十条铁律与随附资源清单。", AMBER_ACCENT)
    ]
    
    for idx, (m_title, m_desc, m_color) in enumerate(modules):
        col = idx % 3
        row = idx // 3
        x = Inches(0.8 + col * 3.95)
        y = Inches(1.6 + row * 2.5)
        
        card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(3.75), Inches(2.2))
        card.fill.solid()
        card.fill.fore_color.rgb = NAVY_CARD
        card.line.color.rgb = m_color
        card.line.width = Pt(1.5)
        
        ctf = card.text_frame
        ctf.word_wrap = True
        ctf.margin_left = ctf.margin_right = Inches(0.2)
        ctf.margin_top = Inches(0.18)
        
        cp0 = ctf.paragraphs[0]
        cr0 = cp0.add_run()
        cr0.text = f"MODULE {idx+1:02d}"
        cr0.font.name = "Microsoft YaHei"
        cr0.font.size = Pt(9.5)
        cr0.font.bold = True
        cr0.font.color.rgb = m_color
        
        cp1 = ctf.add_paragraph()
        cp1.space_before = Pt(4)
        cp1.space_after = Pt(6)
        cr1 = cp1.add_run()
        cr1.text = m_title
        cr1.font.name = "Microsoft YaHei"
        cr1.font.size = Pt(13)
        cr1.font.bold = True
        cr1.font.color.rgb = WHITE
        
        cp2 = ctf.add_paragraph()
        cr2 = cp2.add_run()
        cr2.text = m_desc
        cr2.font.name = "Microsoft YaHei"
        cr2.font.size = Pt(10)
        cr2.font.color.rgb = TEXT_LIGHT
        
    add_footer(s2, 2)

    # =========================================================================
    # SLIDE 3: CHAPTER 1 - EXCEL UI & SHORTCUTS
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    set_slide_background(s3)
    add_header(s3, "Chapter 01", "Excel 操作介面全景解析与航运行政核心工作区", "工欲善其事，必先利其器")
    
    # Left: Interface overview diagram image
    img_ui_path = "images/fig_1_1_excel_interface_overview.png"
    if os.path.exists(img_ui_path):
        s3.shapes.add_picture(img_ui_path, Inches(0.8), Inches(1.5), width=Inches(7.2))
        
    # Right: High frequency shortcuts card
    right_card = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8.2), Inches(1.5), Inches(4.333), Inches(5.1))
    right_card.fill.solid()
    right_card.fill.fore_color.rgb = NAVY_CARD
    right_card.line.color.rgb = NAVY_BORDER
    
    rc_tf = right_card.text_frame
    rc_tf.word_wrap = True
    rc_tf.margin_left = rc_tf.margin_right = Inches(0.2)
    rc_tf.margin_top = Inches(0.2)
    
    rp0 = rc_tf.paragraphs[0]
    rp0.space_after = Pt(8)
    rr0 = rp0.add_run()
    rr0.text = "⚡ 航运行政必备高频快捷键速查"
    rr0.font.name = "Microsoft YaHei"
    rr0.font.size = Pt(12)
    rr0.font.bold = True
    rr0.font.color.rgb = TEAL_ACCENT
    
    sc_items = [
        ("Ctrl + Shift + ↓ / →", "一秒全选整列/整行连续数据，严禁滑鼠长距拖拽。"),
        ("Ctrl + E", "快速填入 (Flash Fill) 模式识别，智能拆分船名航次。"),
        ("Ctrl + H", "开启寻找及取代，批量清除美元符号 '$' 与不可见字符。"),
        ("Ctrl + 1", "开启储存档格式详细设定，规范短日期 YYYY-MM-DD。"),
        ("F4 键", "切换公式绝对引用 ($A$1)，跨表 XLOOKUP 锁定基准码表。")
    ]
    for k, v in sc_items:
        p_k = rc_tf.add_paragraph()
        p_k.space_before = Pt(6)
        r_k = p_k.add_run()
        r_k.text = f"▪ {k}"
        r_k.font.name = "Microsoft YaHei"
        r_k.font.size = Pt(10.5)
        r_k.font.bold = True
        r_k.font.color.rgb = WHITE
        
        p_v = rc_tf.add_paragraph()
        r_v = p_v.add_run()
        r_v.text = f"   {v}"
        r_v.font.name = "Microsoft YaHei"
        r_v.font.size = Pt(9.5)
        r_v.font.color.rgb = TEXT_MUTED

    add_footer(s3, 3)

    # =========================================================================
    # SLIDE 4: CHAPTER 2 - GIGO & 5 DATA TYPES
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    set_slide_background(s4)
    add_header(s4, "Chapter 02", "前处理核心常识：GIGO 铁律与五大资料类型", "Garbage In, Garbage Out")
    
    # GIGO Banner
    banner = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(11.733), Inches(0.9))
    banner.fill.solid()
    banner.fill.fore_color.rgb = RGBColor(50, 15, 15)
    banner.line.color.rgb = RED_ACCENT
    banner.line.width = Pt(1.5)
    
    btf = banner.text_frame
    btf.word_wrap = True
    btf.margin_top = Inches(0.12)
    bp = btf.paragraphs[0]
    br1 = bp.add_run()
    br1.text = "⚠️ GIGO 原则（垃圾进，垃圾出）："
    br1.font.name = "Microsoft YaHei"
    br1.font.size = Pt(12)
    br1.font.bold = True
    br1.font.color.rgb = RED_ACCENT
    
    br2 = bp.add_run()
    br2.text = " 如果输入的航运数据源存在前后空格、格式混杂或重复录入，任何高级公式与透视表都将得出完全失真的假结论。前处理是运费结算与报关零差错的第一道生命线！"
    br2.font.name = "Microsoft YaHei"
    br2.font.size = Pt(10.5)
    br2.font.color.rgb = TEXT_LIGHT
    
    # Table of 5 Data Types
    table_shape = s4.shapes.add_table(6, 4, Inches(0.8), Inches(2.6), Inches(11.733), Inches(3.9))
    table = table_shape.table
    table.columns[0].width = Inches(2.0)
    table.columns[1].width = Inches(1.8)
    table.columns[2].width = Inches(3.4)
    table.columns[3].width = Inches(4.533)
    
    headers = ["资料类型", "Excel 预设对齐", "航运单证对应示例", "初学者常见认知误区与严重后果"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(29, 53, 87)
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = h
        r.font.name = "Microsoft YaHei"
        r.font.size = Pt(10.5)
        r.font.bold = True
        r.font.color.rgb = WHITE
        
    dt_rows = [
        ("文字 (Text)", "靠左对齐 (Left)", "提单号 COSU6291、箱型 40HC、代码 CNSHA", "将纯数字型提单号作为数值存储，导致前面的前导零 '0' 被自动吞噬。"),
        ("数值 (Number)", "靠右对齐 (Right)", "货物毛重 18450、海运费 2450.00、计费吨 32.5", "在格中手打单位 '18450 kg'，导致转为文本，SUM 无法求和直接报 0。"),
        ("日期/时间 (Date)", "靠右对齐 (Right)", "预估开航日 ETD (2026-11-05)、截关时间", "输入 '2026.11.05' 或 '20261105'，系统视为文字，日期相减算航程直接报错。"),
        ("逻辑值 (Boolean)", "置中对齐 (Center)", "TRUE (海关已放行) / FALSE (查验扣留)", "输入带引号的文字 'TRUE'，导致 IF 条件判断逻辑失效。"),
        ("公式与错误值", "依结果而定", "=SUM(H2:H31)、#N/A、#VALUE!、#DIV/0!", "忽视报错隐患，将含错误值的表格直接提交，导致财务报销全盘瘫痪。")
    ]
    for row_idx, r_data in enumerate(dt_rows, start=1):
        for col_idx, text in enumerate(r_data):
            cell = table.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY_CARD if row_idx % 2 == 1 else RGBColor(14, 28, 52)
            p = cell.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = text
            r.font.name = "Microsoft YaHei"
            r.font.size = Pt(9.5)
            r.font.color.rgb = TEAL_LIGHT if col_idx == 0 else TEXT_LIGHT

    add_footer(s4, 4)

    # =========================================================================
    # SLIDE 5: CHAPTER 2 - TIDY DATA & MESSY DATA
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    set_slide_background(s5)
    add_header(s5, "Chapter 02", "整洁数据规范 (Tidy Data) 与航运典型脏污形态", "Codd & Wickham 黄金法则")
    
    # Left: Tidy vs Messy image
    img_tidy_path = "images/fig_2_1_tidy_vs_messy_data.png"
    if os.path.exists(img_tidy_path):
        s5.shapes.add_picture(img_tidy_path, Inches(0.8), Inches(1.5), width=Inches(6.8))
        
    # Right: 3 Rules card & 4 dirty data roots
    rc5 = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), Inches(1.5), Inches(4.733), Inches(5.1))
    rc5.fill.solid()
    rc5.fill.fore_color.rgb = NAVY_CARD
    rc5.line.color.rgb = NAVY_BORDER
    
    rc5_tf = rc5.text_frame
    rc5_tf.word_wrap = True
    rc5_tf.margin_left = rc5_tf.margin_right = Inches(0.25)
    rc5_tf.margin_top = Inches(0.2)
    
    p = rc5_tf.paragraphs[0]
    p.space_after = Pt(6)
    r = p.add_run()
    r.text = "★ 航运整洁数据三大黄金法则："
    r.font.name = "Microsoft YaHei"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = TEAL_ACCENT
    
    rules = [
        ("1. 每一列 (Column) 唯一变量", "装运港独占一列、目的港独占一列、船名与航次独立分列，杜绝复合含义。"),
        ("2. 每一行 (Row) 唯一事件", "单行记录单一提单或单只集装箱，具备清晰物理主键。"),
        ("3. 表头单一、无合并格", "第一行仅包含简洁栏位名称，全表杜绝跨行跨列合并储存档与中间小计行。")
    ]
    for r_t, r_d in rules:
        p_t = rc5_tf.add_paragraph()
        p_t.space_before = Pt(4)
        rt = p_t.add_run()
        rt.text = f"▪ {r_t}"
        rt.font.name = "Microsoft YaHei"
        rt.font.size = Pt(10.5)
        rt.font.bold = True
        rt.font.color.rgb = WHITE
        
        p_d = rc5_tf.add_paragraph()
        rd = p_d.add_run()
        rd.text = f"   {r_d}"
        rd.font.name = "Microsoft YaHei"
        rd.font.size = Pt(9.5)
        rd.font.color.rgb = TEXT_MUTED
        
    p_dirty = rc5_tf.add_paragraph()
    p_dirty.space_before = Pt(12)
    p_dirty.space_after = Pt(4)
    rdt = p_dirty.add_run()
    rdt.text = "⚠️ 航运 4 大脏污根源："
    rdt.font.name = "Microsoft YaHei"
    rdt.font.size = Pt(11)
    rdt.font.bold = True
    rdt.font.color.rgb = RED_ACCENT
    
    roots = [
        "① EDI/网页复制夹带不可见字符 Chr(10)/Chr(160)",
        "② 港口代码大小写与简称混杂 (cnsha vs SHA)",
        "③ 跨国日期格式混杂 (20261105 / 斜杠 / 点号)",
        "④ 货代多次改单导致的幽灵重复提单与货柜"
    ]
    for root in roots:
        pr = rc5_tf.add_paragraph()
        rr = pr.add_run()
        rr.text = root
        rr.font.name = "Microsoft YaHei"
        rr.font.size = Pt(9)
        rr.font.color.rgb = TEXT_LIGHT

    add_footer(s5, 5)

    # =========================================================================
    # SLIDE 6: CHAPTER 3 - 30 RECORDS WORKFLOW
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    set_slide_background(s6)
    add_header(s6, "Chapter 03", "30 笔航运实务示范档：资料前处理全流程示范", "标准化 8 步清洗工序")
    
    steps = [
        ("01", "建立副本与备份", "复制一份原始工作表，绝对禁止在唯一源数据上直接进行破坏性编辑。", TEAL_ACCENT),
        ("02", "双主键精准去重", "勾选「提单号+柜号」，精准识别并剔除 2 笔重复单证，保留 28 个唯一值。", TEAL_LIGHT),
        ("03", "文本与大小写清洗", "套用 TRIM+CLEAN+UPPER 组合公式，清除前后多余空格与不可见控制字符。", RGBColor(59, 130, 246)),
        ("04", "资料剖析拆分复合栏", "以斜杠 '/' 分割「船名航次」，拆解为独立的「船名」与「航次」两个变量。", RGBColor(147, 51, 234)),
        ("05", "日期格式标准化", "利用 DATEVALUE 替换点号与斜杠，统一规范为系统标准短日期 YYYY-MM-DD。", RGBColor(236, 72, 153)),
        ("06", "剥离单位与数值化", "SUBSTITUTE 剔除 'kg'、'$'，VALUE 转为可求和数字，设置货币与千分位格式。", AMBER_ACCENT),
        ("07", "XLOOKUP 匹配码表", "跨表关联《基础代码参照表》，自动将 CNSHA 翻译补全为「上海港 (Shanghai)」。", GREEN_ACCENT),
        ("08", "选择性贴上为值", "复制清洗结果【仅贴上数值】，切断动态公式依赖，彻底固化干净实体数据。", RGBColor(245, 158, 11))
    ]
    
    for idx, (num, title, desc, colr) in enumerate(steps):
        col = idx % 4
        row = idx // 4
        x = Inches(0.8 + col * 2.95)
        y = Inches(1.6 + row * 2.5)
        
        card = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(2.8), Inches(2.2))
        card.fill.solid()
        card.fill.fore_color.rgb = NAVY_CARD
        card.line.color.rgb = colr
        card.line.width = Pt(1.5)
        
        ctf = card.text_frame
        ctf.word_wrap = True
        ctf.margin_left = ctf.margin_right = Inches(0.18)
        ctf.margin_top = Inches(0.18)
        
        p0 = ctf.paragraphs[0]
        r0 = p0.add_run()
        r0.text = f"STEP {num}"
        r0.font.name = "Microsoft YaHei"
        r0.font.size = Pt(10)
        r0.font.bold = True
        r0.font.color.rgb = colr
        
        p1 = ctf.add_paragraph()
        p1.space_before = Pt(4)
        p1.space_after = Pt(6)
        r1 = p1.add_run()
        r1.text = title
        r1.font.name = "Microsoft YaHei"
        r1.font.size = Pt(12)
        r1.font.bold = True
        r1.font.color.rgb = WHITE
        
        p2 = ctf.add_paragraph()
        r2 = p2.add_run()
        r2.text = desc
        r2.font.name = "Microsoft YaHei"
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = TEXT_LIGHT

    add_footer(s6, 6)

    # =========================================================================
    # SLIDE 7: CHAPTER 3 - BEFORE VS AFTER
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    set_slide_background(s7)
    add_header(s7, "Chapter 03", "30 笔数据清洗前后对照实验室 (Before vs After)", "蜕变全景")
    
    # 4 Indicator metric cards
    metrics = [
        ("有效单证行数", "30 笔 (含2笔重复) ➔ 28 笔纯净", GREEN_ACCENT),
        ("运费字段求和", "文本报错 ➔ $61,250.00 ✓", TEAL_ACCENT),
        ("日期合规率", "33% 混杂 ➔ 100% YYYY-MM-DD", TEAL_LIGHT),
        ("港口中文解析", "仅存代码 ➔ 100% 自动对齐中文港名", AMBER_ACCENT)
    ]
    for idx, (m_lbl, m_val, m_c) in enumerate(metrics):
        x = Inches(0.8 + idx * 2.95)
        y = Inches(1.5)
        b = s7.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(2.8), Inches(0.85))
        b.fill.solid()
        b.fill.fore_color.rgb = NAVY_CARD
        b.line.color.rgb = m_c
        btf = b.text_frame
        btf.margin_top = Inches(0.1)
        p1 = btf.paragraphs[0]
        r1 = p1.add_run()
        r1.text = m_lbl
        r1.font.name = "Microsoft YaHei"
        r1.font.size = Pt(9)
        r1.font.color.rgb = TEXT_MUTED
        
        p2 = btf.add_paragraph()
        r2 = p2.add_run()
        r2.text = m_val
        r2.font.name = "Microsoft YaHei"
        r2.font.size = Pt(10.5)
        r2.font.bold = True
        r2.font.color.rgb = WHITE
        
    # Table of Before vs After sample
    t7_shape = s7.shapes.add_table(7, 5, Inches(0.8), Inches(2.55), Inches(11.733), Inches(4.1))
    t7 = t7_shape.table
    t7.columns[0].width = Inches(1.0)
    t7.columns[1].width = Inches(2.6)
    t7.columns[2].width = Inches(2.6)
    t7.columns[3].width = Inches(2.6)
    t7.columns[4].width = Inches(2.933)
    
    h7 = ["对比项", "原始脏数据形态 (Before)", "清洗后标准形态 (After)", "采用处理手段", "业务价值收益"]
    for i, h in enumerate(h7):
        cell = t7.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(29, 53, 87)
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = h
        r.font.name = "Microsoft YaHei"
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = WHITE
        
    comp_rows = [
        ("提单与箱号", "带前后空格、小写 (如 ' cosu9823 ')", "纯净大写 (COSU9823145)", "=TRIM(CLEAN(UPPER(..)))", "消灭匹配盲点，提高EDI对齐率"),
        ("重复单证", "整行完全重复 (如行4与行21)", "精准剔除重复项，保留唯一记录", "【资料】->【移除重复项】", "杜绝财务运费重复支付损失"),
        ("船名航次", "挤在同格 (EVER GIVEN/0118E)", "分立为「船名」与「航次」两列", "【资料剖析】按斜杠 '/' 拆分", "便于按航次单独进行航程与滞箱统计"),
        ("挂靠港口", "小写混写 (cnsha / deham)", "标准代码并补全「上海港」、「汉堡港」", "XLOOKUP 跨表匹配基础码表", "大幅提升物流跟单与客户沟通可读性"),
        ("开航日期", "混杂点号、斜杠与8位数字", "标准 ISO 格式 (2026-11-05)", "DATEVALUE + 嵌套 SUBSTITUTE", "支持直接以日期相减计算实际航程"),
        ("海运费用", "带文字单位 (如 'USD 1200'、'$2,450')", "纯数字型货币 (2450.00)，设千分位", "VALUE(SUBSTITUTE(..)) 剥离字符", "彻底根除求和时报 #VALUE! 的隐患")
    ]
    for row_idx, r_data in enumerate(comp_rows, start=1):
        for col_idx, text in enumerate(r_data):
            cell = t7.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY_CARD if row_idx % 2 == 1 else RGBColor(14, 28, 52)
            p = cell.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = text
            r.font.name = "Microsoft YaHei"
            r.font.size = Pt(9)
            r.font.color.rgb = TEAL_LIGHT if col_idx == 0 else (RED_ACCENT if col_idx == 1 else (GREEN_ACCENT if col_idx == 2 else TEXT_LIGHT))

    add_footer(s7, 7)

    # =========================================================================
    # SLIDE 8: CHAPTER 4 - EXCEL CORE FUNCTIONS
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    set_slide_background(s8)
    add_header(s8, "Chapter 04", "资料前处理功能选项深度操作攻略", "无需代码的前处理利器")
    
    # Left: Text to Columns image
    img_split_path = "images/fig_4_1_text_to_columns.png"
    if os.path.exists(img_split_path):
        s8.shapes.add_picture(img_split_path, Inches(0.8), Inches(1.5), width=Inches(5.7))
        
    # Right: Remove Duplicates & Validation image
    img_dedup_path = "images/fig_4_2_remove_duplicates_and_validation.png"
    if os.path.exists(img_dedup_path):
        s8.shapes.add_picture(img_dedup_path, Inches(6.8), Inches(1.5), width=Inches(5.7))

    add_footer(s8, 8)

    # =========================================================================
    # SLIDE 9: CHAPTER 4 - CLEANING FORMULAS
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    set_slide_background(s9)
    add_header(s9, "Chapter 04", "资料清洗核心函数公式武器库与万能嵌套模板", "自动化清洗神咒")
    
    # Left: Formula showcase image
    img_fm_path = "images/fig_4_3_formulas_cleaning_showcase.png"
    if os.path.exists(img_fm_path):
        s9.shapes.add_picture(img_fm_path, Inches(0.8), Inches(1.5), width=Inches(6.8))
        
    # Right: 3 Universal Nested Templates Cards
    rc9 = s9.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.8), Inches(1.5), Inches(4.733), Inches(5.1))
    rc9.fill.solid()
    rc9.fill.fore_color.rgb = NAVY_CARD
    rc9.line.color.rgb = NAVY_BORDER
    
    rc9_tf = rc9.text_frame
    rc9_tf.word_wrap = True
    rc9_tf.margin_left = rc9_tf.margin_right = Inches(0.2)
    rc9_tf.margin_top = Inches(0.2)
    
    p = rc9_tf.paragraphs[0]
    p.space_after = Pt(6)
    r = p.add_run()
    r.text = "🎯 航运实务三大万能嵌套公式模板："
    r.font.name = "Microsoft YaHei"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = TEAL_ACCENT
    
    templates = [
        ("① 航运文本降魔神咒：", 
         "=TRIM(CLEAN(UPPER(SUBSTITUTE(A2, CHAR(160), \"\"))))", 
         "彻底清除网页导出的 Chr(160) 不间断空格、隐藏控制符与前后多余空格，并统一为大写。"),
        ("② 混杂文本日期强力归一：", 
         "=IF(ISNUMBER(A2), A2, DATEVALUE(SUBSTITUTE(SUBSTITUTE(A2, \".\", \"-\"), \"/\", \"-\")))", 
         "兼容点号「2026.11.05」与斜杠「2026/11/05」，无缝转换为合法系统真日期。"),
        ("③ 运费金额剥离净化：", 
         "=VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(UPPER(A2), \"USD\", \"\"), \"$\", \"\"), \",\", \"\"))", 
         "剔除运费中的「$」、「USD」与千分位逗号，转化为真正可直接求和的纯数字。")
    ]
    for t_title, t_code, t_desc in templates:
        pt = rc9_tf.add_paragraph()
        pt.space_before = Pt(8)
        rt = pt.add_run()
        rt.text = t_title
        rt.font.name = "Microsoft YaHei"
        rt.font.size = Pt(10.5)
        rt.font.bold = True
        rt.font.color.rgb = AMBER_ACCENT
        
        pc = rc9_tf.add_paragraph()
        pc.space_before = Pt(2)
        rc = pc.add_run()
        rc.text = t_code
        rc.font.name = "Consolas"
        rc.font.size = Pt(8.5)
        rc.font.color.rgb = TEAL_LIGHT
        
        pd = rc9_tf.add_paragraph()
        pd.space_before = Pt(2)
        rd = pd.add_run()
        rd.text = t_desc
        rd.font.name = "Microsoft YaHei"
        rd.font.size = Pt(9)
        rd.font.color.rgb = TEXT_MUTED

    add_footer(s9, 9)

    # =========================================================================
    # SLIDE 10: CHAPTER 5 - 300 RECORDS PRACTICE ARCHITECTURE
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    set_slide_background(s10)
    add_header(s10, "Chapter 05", "300 笔多场景航运综合实战练习档业务架构", "三大航运实务场景")
    
    # Pipeline image
    img_pipe_path = "images/fig_5_1_practice_pipeline.png"
    if os.path.exists(img_pipe_path):
        s10.shapes.add_picture(img_pipe_path, Inches(0.8), Inches(1.5), width=Inches(11.733))
        
    # Scenario breakdown cards below
    sc_cards = [
        ("场景 A：集装箱流转与跟踪 (100笔)", "工作单号、主提单号、柜号、ISO 箱型、船东代码、铅封完整性。重点考核格式规范与主键去重。", RGBColor(59, 130, 246)),
        ("场景 B：海运运费与滞期费结算 (100笔)", "货物体积 (CBM)、毛重 (KGS)、计费吨 RT=MAX(CBM, KGS/1000)、海运费、超期滞箱费与应收总账。", GREEN_ACCENT),
        ("场景 C：船期准点率与港口作业 (100笔)", "船名航次剖析、UN/LOCODE 港口码表关联、ETD/ATD/ETA/ATA 航程运算与延误天数核算。", AMBER_ACCENT)
    ]
    for idx, (sc_t, sc_d, sc_c) in enumerate(sc_cards):
        x = Inches(0.8 + idx * 3.95)
        y = Inches(5.15)
        sc_b = s10.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(3.75), Inches(1.55))
        sc_b.fill.solid()
        sc_b.fill.fore_color.rgb = NAVY_CARD
        sc_b.line.color.rgb = sc_c
        sc_b.line.width = Pt(1.5)
        
        sctf = sc_b.text_frame
        sctf.word_wrap = True
        sctf.margin_left = sctf.margin_right = Inches(0.2)
        sctf.margin_top = Inches(0.12)
        
        p1 = sctf.paragraphs[0]
        r1 = p1.add_run()
        r1.text = sc_t
        r1.font.name = "Microsoft YaHei"
        r1.font.size = Pt(11)
        r1.font.bold = True
        r1.font.color.rgb = WHITE
        
        p2 = sctf.add_paragraph()
        p2.space_before = Pt(4)
        r2 = p2.add_run()
        r2.text = sc_d
        r2.font.name = "Microsoft YaHei"
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = TEXT_LIGHT

    add_footer(s10, 10)

    # =========================================================================
    # SLIDE 11: CHAPTER 5 - 10 CHALLENGES & SOLUTIONS
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    set_slide_background(s11)
    add_header(s11, "Chapter 05", "十大核心考核任务挑战与参考题解", "实训通关秘籍")
    
    t11_shape = s11.shapes.add_table(7, 4, Inches(0.8), Inches(1.5), Inches(11.733), Inches(5.1))
    t11 = t11_shape.table
    t11.columns[0].width = Inches(1.0)
    t11.columns[1].width = Inches(2.2)
    t11.columns[2].width = Inches(4.8)
    t11.columns[3].width = Inches(3.733)
    
    h11 = ["任务编号", "考核前处理模块", "标准操作步骤与核心公式", "参考答案与判定准则"]
    for i, h in enumerate(h11):
        cell = t11.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(29, 53, 87)
        p = cell.text_frame.paragraphs[0]
        r = p.add_run()
        r.text = h
        r.font.name = "Microsoft YaHei"
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = WHITE
        
    tasks_table_data = [
        ("任务 01", "多重主键精准去重", "选定 A1:V301 ->【资料】->【移除重复项】-> 仅勾选「工作单号」与「主提单号」", "精准识别并删除 3 笔重复项，数据行由 300 行收敛至 297 行。"),
        ("任务 02", "单证与箱号规范化", "套用公式 =TRIM(CLEAN(UPPER(D2))) 清洗前后空格与转为大写", "提单号与柜号全部转为标准大写，无任何残留隐形控制符。"),
        ("任务 03", "复合船名航次分列", "选中船名航次列 ->【资料剖析】-> 分隔符号勾选「其他」，输入斜杠 [/]", "成功拆解为独立的「船名」与「航次」两列单一属性变量。"),
        ("任务 04", "港口自动补全中文", "公式：=XLOOKUP(TRIM(UPPER(I2)), 参照表!$B$2:$B$19, 参照表!$C$2:$C$19)", "全部 297 笔记录成功解析出「上海港」、「汉堡港」等中文全称。"),
        ("任务 05", "日期标准化与航程", "计划航程 =ETA列 - ETD列  |  延误天数 =MAX(0, ATA列 - ETA列)", "平均计划航程约 23.5 天；超期到港返回实际延误天数。"),
        ("任务 06", "计费吨与运费核算", "计费吨 =MAX(CBM, ROUND(KGS/1000, 2))  |  总费用 = 基础运费 + 滞箱费", "297 笔有效业务应收总费用公式 =SUM(总费用列)，总额约 $445,000。")
    ]
    for row_idx, r_data in enumerate(tasks_table_data, start=1):
        for col_idx, text in enumerate(r_data):
            cell = t11.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = NAVY_CARD if row_idx % 2 == 1 else RGBColor(14, 28, 52)
            p = cell.text_frame.paragraphs[0]
            r = p.add_run()
            r.text = text
            r.font.name = "Microsoft YaHei"
            r.font.size = Pt(9.5)
            r.font.color.rgb = TEAL_LIGHT if col_idx == 0 else (WHITE if col_idx == 1 else (GREEN_ACCENT if col_idx == 3 else TEXT_LIGHT))

    add_footer(s11, 11)

    # =========================================================================
    # SLIDE 12: CHAPTER 6 - LONG-TERM DATA QUALITY CONTROL
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    set_slide_background(s12)
    add_header(s12, "Chapter 06", "航运数据品质长效管控与自动化进阶建议", "制度与技术双轮驱动")
    
    # Left Card: Form Defenses
    c12_l = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.5), Inches(5.7), Inches(5.1))
    c12_l.fill.solid()
    c12_l.fill.fore_color.rgb = NAVY_CARD
    c12_l.line.color.rgb = AMBER_ACCENT
    c12_l.line.width = Pt(1.5)
    
    tl = c12_l.text_frame
    tl.word_wrap = True
    tl.margin_left = tl.margin_right = Inches(0.25)
    tl.margin_top = Inches(0.2)
    
    p = tl.paragraphs[0]
    p.space_after = Pt(8)
    r = p.add_run()
    r.text = "🛡️ 源头表单设计防错三机制："
    r.font.name = "Microsoft YaHei"
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = AMBER_ACCENT
    
    defenses = [
        ("1. 刚性下拉选单约束", "对港口（UN/LOCODE）、箱型（ISO 6346）、船东代码设立「资料验证」下拉清单，杜绝自由手打乱拼。"),
        ("2. 字数与格式掩码阻断", "集装箱号强制限定 11 位，前 4 位必须为字母，未符合格式者系统直接跳出弹跳警告阻止提交。"),
        ("3. 拥抱 Ctrl+T 格式化表格", "将数据定义为「Excel 表格」，新增记录行时自动向下继承所有公式、验证规则与格式。")
    ]
    for dt, dd in defenses:
        pt = tl.add_paragraph()
        pt.space_before = Pt(6)
        rt = pt.add_run()
        rt.text = f"▪ {dt}"
        rt.font.name = "Microsoft YaHei"
        rt.font.size = Pt(11)
        rt.font.bold = True
        rt.font.color.rgb = WHITE
        
        pd = tl.add_paragraph()
        rd = pd.add_run()
        rd.text = f"   {dd}"
        rd.font.name = "Microsoft YaHei"
        rd.font.size = Pt(9.5)
        rd.font.color.rgb = TEXT_LIGHT

    # Right Card: 10 Golden Rules & Power Query
    c12_r = s12.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.5), Inches(5.733), Inches(5.1))
    c12_r.fill.solid()
    c12_r.fill.fore_color.rgb = NAVY_CARD
    c12_r.line.color.rgb = TEAL_ACCENT
    c12_r.line.width = Pt(1.5)
    
    tr = c12_r.text_frame
    tr.word_wrap = True
    tr.margin_left = tr.margin_right = Inches(0.25)
    tr.margin_top = Inches(0.2)
    
    p = tr.paragraphs[0]
    p.space_after = Pt(6)
    r = p.add_run()
    r.text = "📜 航运行政日常数据处理精选铁律："
    r.font.name = "Microsoft YaHei"
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = TEAL_ACCENT
    
    rules = [
        ("铁律 1 备份先行", "原始数据永远留有纯净底稿，严禁破坏性编辑。"),
        ("铁律 2 数文分离", "可计算数值格绝不附带 kg、usd 等文字单位。"),
        ("铁律 3 拒绝合并", "明细表内杜绝任何跨行跨列合并储存档。"),
        ("铁律 4 连续网格", "严禁插入空行隔开记录，小计交由透视表处理。"),
        ("铁律 5 固化数值", "复杂公式前处理后务必【选择性贴上为值】。")
    ]
    for rt, rd in rules:
        pt = tr.add_paragraph()
        pt.space_before = Pt(4)
        r = pt.add_run()
        r.text = f"▪ {rt}："
        r.font.name = "Microsoft YaHei"
        r.font.size = Pt(10)
        r.font.bold = True
        r.font.color.rgb = WHITE
        
        r2 = pt.add_run()
        r2.text = f" {rd}"
        r2.font.name = "Microsoft YaHei"
        r2.font.size = Pt(9.5)
        r2.font.color.rgb = TEXT_LIGHT
        
    p_pq = tr.add_paragraph()
    p_pq.space_before = Pt(12)
    rpq = p_pq.add_run()
    rpq.text = "🚀 进阶自动化展望：Power Query (M 语言)"
    rpq.font.name = "Microsoft YaHei"
    rpq.font.size = Pt(10.5)
    rpq.font.bold = True
    rpq.font.color.rgb = TEAL_LIGHT
    
    ppq2 = tr.add_paragraph()
    rpq2 = ppq2.add_run()
    rpq2.text = "掌握本手册技能后，推荐学习 Power Query。只需将船期更新文件扔进文件夹，点击【全部重新整理】即可全自动秒级清洗！"
    rpq2.font.name = "Microsoft YaHei"
    rpq2.font.size = Pt(9)
    rpq2.font.color.rgb = TEXT_MUTED

    add_footer(s12, 12)

    # =========================================================================
    # SLIDE 13: SUMMARY & RESOURCES
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    set_slide_background(s13)
    add_header(s13, "Course Completion", "课程总结与随附资源清单", "学以致用 • 赋能实务")
    
    # Large Card
    c13 = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.0), Inches(1.5), Inches(11.333), Inches(5.1))
    c13.fill.solid()
    c13.fill.fore_color.rgb = NAVY_CARD
    c13.line.color.rgb = TEAL_ACCENT
    c13.line.width = Pt(1.5)
    
    t13 = c13.text_frame
    t13.word_wrap = True
    t13.margin_left = t13.margin_right = Inches(0.4)
    t13.margin_top = Inches(0.3)
    
    p = t13.paragraphs[0]
    p.space_after = Pt(8)
    r = p.add_run()
    r.text = "🎉 恭喜完成航运 Excel 资料前处理全景培训！"
    r.font.name = "Microsoft YaHei"
    r.font.size = Pt(22)
    r.font.bold = True
    r.font.color.rgb = WHITE
    
    p_desc = t13.add_paragraph()
    p_desc.space_after = Pt(16)
    rd = p_desc.add_run()
    rd.text = "通过本套教学与实战练习，您已完整具备现代航运企业行政与数字化运营所需的【数据清洗、标准化、跨表关联与品质管控】全套核心操作技能。"
    rd.font.name = "Microsoft YaHei"
    rd.font.size = Pt(12)
    rd.font.color.rgb = TEXT_LIGHT
    
    # 3 Deliverable Boxes
    boxes = [
        ("📄 Word 排版主教材", "航运行政人员_Excel基础资料分析与资料前处理教学手册.docx", "含封面、图表目录、页码与参考文献之出版级完整手册 (1.47 MB)"),
        ("📊 30 笔实务示范档", "航运资料前处理_30笔示范档.xlsx", "内含原始脏数据、清洗后标准数据及分步公式对齐表"),
        ("🎯 300 笔情境实训档", "航运资料前处理_300笔情境练习档.xlsx", "涵盖集装箱流转、运费滞期、船期准点率三大场景与完整题解")
    ]
    for idx, (b_t, b_f, b_d) in enumerate(boxes):
        x = Inches(1.4 + idx * 3.55)
        y = Inches(3.2)
        box = s13.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(3.3), Inches(1.8))
        box.fill.solid()
        box.fill.fore_color.rgb = NAVY_DARK
        box.line.color.rgb = TEAL_LIGHT
        
        btf = box.text_frame
        btf.word_wrap = True
        btf.margin_left = btf.margin_right = Inches(0.18)
        btf.margin_top = Inches(0.15)
        
        bp1 = btf.paragraphs[0]
        br1 = bp1.add_run()
        br1.text = b_t
        br1.font.name = "Microsoft YaHei"
        br1.font.size = Pt(11)
        br1.font.bold = True
        br1.font.color.rgb = WHITE
        
        bp2 = btf.add_paragraph()
        bp2.space_before = Pt(4)
        br2 = bp2.add_run()
        br2.text = b_f
        br2.font.name = "Consolas"
        br2.font.size = Pt(9)
        br2.font.color.rgb = TEAL_ACCENT
        
        bp3 = btf.add_paragraph()
        bp3.space_before = Pt(4)
        br3 = bp3.add_run()
        br3.text = b_d
        br3.font.name = "Microsoft YaHei"
        br3.font.size = Pt(8.5)
        br3.font.color.rgb = TEXT_MUTED
        
    p_ref = t13.add_paragraph()
    p_ref.space_before = Pt(130)
    rr = p_ref.add_run()
    rr.text = "📚 参考文献与行业标准：国际海事组织 IMO FAL 公约 • ISO 6346 集装箱编码标准 • 联合国 UN/LOCODE 港口代码 • Hadley Wickham Tidy Data • 微软官方知识库"
    rr.font.name = "Microsoft YaHei"
    rr.font.size = Pt(9)
    rr.font.color.rgb = TEXT_MUTED

    add_footer(s13, 13)

    # Save presentation
    out_pptx_path = "航运行政人员_Excel基础资料分析与资料前处理教学手册.pptx"
    prs.save(out_pptx_path)
    print(f"Successfully generated PowerPoint file: {out_pptx_path}")

if __name__ == "__main__":
    create_shipping_tutorial_pptx()
