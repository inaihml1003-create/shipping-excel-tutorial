import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def create_shipping_tutorial_docx():
    doc = Document()
    
    # -------------------------------------------------------------
    # Page Setup - Margins (2.54 cm standard)
    # -------------------------------------------------------------
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.page_width = Inches(8.27)  # A4
        section.page_height = Inches(11.69)
        section.different_first_page_header_footer = True
        
    # Color palette
    NAVY_HEX = "1F497D"
    DARK_GREEN_HEX = "274E13"
    SLATE_HEX = "595959"
    LIGHT_BG_HEX = "F2F4F7"
    BORDER_HEX = "D9D9D9"
    ACCENT_BLUE = RGBColor(31, 73, 125)
    ACCENT_GREEN = RGBColor(39, 78, 19)
    TEXT_DARK = RGBColor(51, 51, 51)
    
    # -------------------------------------------------------------
    # Helper Functions
    # -------------------------------------------------------------
    def set_cell_background(cell, fill_hex):
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
        cell._tc.get_or_add_tcPr().append(shading)
        
    def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
        tcPr = cell._tc.get_or_add_tcPr()
        tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
        tcPr.append(tcMar)
        
    def add_page_number_to_run(run):
        fldSimple = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
        run._r.append(fldSimple)

    def add_numpages_to_run(run):
        fldSimple = parse_xml(r'<w:fldSimple %s w:instr="NUMPAGES"/>' % nsdecls('w'))
        run._r.append(fldSimple)
        
    def style_paragraph(p, space_before=0, space_after=6, line_spacing=1.2, keep_with_next=False):
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        p.paragraph_format.keep_with_next = keep_with_next

    def add_custom_heading(text, level, prefix=""):
        p = doc.add_paragraph()
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.bold = True
        run.font.name = "Microsoft YaHei"
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        
        if level == 1:
            p.paragraph_format.space_before = Pt(18)
            p.paragraph_format.space_after = Pt(8)
            run.font.size = Pt(16)
            run.font.color.rgb = ACCENT_BLUE
            # Bottom border for Heading 1
            pPr = p._p.get_or_add_pPr()
            pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="4" w:color="{NAVY_HEX}"/></w:pBdr>')
            pPr.append(pBdr)
        elif level == 2:
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)
            run.font.size = Pt(13)
            run.font.color.rgb = ACCENT_GREEN
        elif level == 3:
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(4)
            run.font.size = Pt(11)
            run.font.color.rgb = ACCENT_BLUE
        return p

    def add_body_paragraph(text, bold_prefix="", indent=False):
        p = doc.add_paragraph()
        style_paragraph(p, space_before=0, space_after=5, line_spacing=1.25)
        if indent:
            p.paragraph_format.first_line_indent = Inches(0.28)
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.bold = True
            r_pre.font.name = "Microsoft YaHei"
            r_pre._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
            r_pre.font.size = Pt(10)
            r_pre.font.color.rgb = ACCENT_BLUE
        r = p.add_run(text)
        r.font.name = "Microsoft YaHei"
        r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r.font.size = Pt(10)
        r.font.color.rgb = TEXT_DARK
        return p

    def add_bullet_point(text, bold_title=""):
        p = doc.add_paragraph(style='List Bullet')
        style_paragraph(p, space_before=0, space_after=3, line_spacing=1.2)
        if bold_title:
            r_bt = p.add_run(bold_title)
            r_bt.bold = True
            r_bt.font.name = "Microsoft YaHei"
            r_bt._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
            r_bt.font.size = Pt(10)
            r_bt.font.color.rgb = ACCENT_BLUE
        r = p.add_run(text)
        r.font.name = "Microsoft YaHei"
        r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r.font.size = Pt(10)
        r.font.color.rgb = TEXT_DARK
        return p

    def add_callout_box(title, text, box_type="NOTE"):
        # Single cell table for callout box
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        tbl.autofit = False
        cell = tbl.cell(0, 0)
        cell.width = Inches(6.27)
        
        border_color = "1F497D" if box_type == "NOTE" else ("274E13" if box_type == "TIP" else "C00000")
        bg_color = "F2F7FA" if box_type == "NOTE" else ("F3F9F4" if box_type == "TIP" else "FFF4F4")
        
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(f'<w:tcBorders {nsdecls("w")}><w:top w:val="none"/><w:left w:val="single" w:sz="24" w:space="0" w:color="{border_color}"/><w:bottom w:val="none"/><w:right w:val="none"/></w:tcBorders>')
        tcPr.append(borders)
        set_cell_background(cell, bg_color)
        set_cell_margins(cell, top=120, bottom=120, left=180, right=150)
        
        cp = cell.paragraphs[0]
        style_paragraph(cp, space_before=0, space_after=3, line_spacing=1.2)
        r_t = cp.add_run(f"【{title}】\n")
        r_t.bold = True
        r_t.font.name = "Microsoft YaHei"
        r_t._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r_t.font.size = Pt(10)
        r_t.font.color.rgb = ACCENT_BLUE if box_type == "NOTE" else (ACCENT_GREEN if box_type == "TIP" else RGBColor(192, 0, 0))
        
        r_b = cp.add_run(text)
        r_b.font.name = "Microsoft YaHei"
        r_b._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r_b.font.size = Pt(9.5)
        r_b.font.color.rgb = TEXT_DARK
        
        # spacing after callout
        p_sp = doc.add_paragraph()
        p_sp.paragraph_format.space_before = Pt(0)
        p_sp.paragraph_format.space_after = Pt(4)

    def add_styled_image(img_path, caption, width_in_inches=5.8):
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(6)
            p_img.paragraph_format.space_after = Pt(2)
            p_img.paragraph_format.keep_with_next = True
            run_img = p_img.add_run()
            run_img.add_picture(img_path, width=Inches(width_in_inches))
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(8)
            run_cap = p_cap.add_run(caption)
            run_cap.font.name = "Microsoft YaHei"
            run_cap._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
            run_cap.font.size = Pt(9)
            run_cap.bold = True
            run_cap.font.color.rgb = ACCENT_BLUE

    # -------------------------------------------------------------
    # COVER PAGE
    # -------------------------------------------------------------
    p_pre = doc.add_paragraph()
    p_pre.paragraph_format.space_before = Pt(36)
    
    # Title box table
    tbl_cov = doc.add_table(rows=1, cols=1)
    tbl_cov.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_cov = tbl_cov.cell(0, 0)
    c_cov.width = Inches(6.27)
    set_cell_background(c_cov, NAVY_HEX)
    set_cell_margins(c_cov, top=260, bottom=260, left=240, right=240)
    
    p_t = c_cov.paragraphs[0]
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_paragraph(p_t, space_before=0, space_after=10, line_spacing=1.2)
    r_main_t = p_t.add_run("航运行政人员 Excel 基础资料分析\n资料前处理核心技能实务教学手册")
    r_main_t.bold = True
    r_main_t.font.name = "Microsoft YaHei"
    r_main_t._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    r_main_t.font.size = Pt(22)
    r_main_t.font.color.rgb = RGBColor(255, 255, 255)
    
    p_sub = c_cov.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_paragraph(p_sub, space_before=0, space_after=0, line_spacing=1.2)
    r_sub = p_sub.add_run("零基础系统化实战教程：从操作介面、整洁数据规范到 30 笔实务示范与 300 笔情境挑战")
    r_sub.font.name = "Microsoft YaHei"
    r_sub._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    r_sub.font.size = Pt(12)
    r_sub.font.color.rgb = RGBColor(230, 240, 255)
    
    # Metadata info on cover
    p_meta_space = doc.add_paragraph()
    p_meta_space.paragraph_format.space_before = Pt(80)
    
    tbl_meta = doc.add_table(rows=6, cols=2)
    tbl_meta.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_meta.autofit = False
    
    meta_info = [
        ("适用对象", "航运公司、船务代理、国际货代、码头堆场及物流行政初学者"),
        ("示范工具", "Microsoft Excel 365 / 2021 / 2019 专业版"),
        ("核心关注", "基础资料分析之资料前处理（Data Preprocessing）必备操作与函数"),
        ("实务范例", "30 笔航运进出口示范档 + 300 笔多场景综合情境练习档"),
        ("编撰机构", "智慧港航数字化人才教研组"),
        ("发布日期", "2026 年 11 月 | 版本 V1.0 正式完整版")
    ]
    
    for row_idx, (k, v) in enumerate(meta_info):
        c0 = tbl_meta.cell(row_idx, 0)
        c1 = tbl_meta.cell(row_idx, 1)
        c0.width = Inches(1.8)
        c1.width = Inches(4.47)
        set_cell_margins(c0, top=70, bottom=70, left=100, right=100)
        set_cell_margins(c1, top=70, bottom=70, left=100, right=100)
        set_cell_background(c0, "F2F4F7")
        set_cell_background(c1, "FFFFFF")
        
        p0 = c0.paragraphs[0]
        style_paragraph(p0, space_before=0, space_after=0, line_spacing=1.15)
        r0 = p0.add_run(k)
        r0.bold = True
        r0.font.name = "Microsoft YaHei"
        r0._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r0.font.size = Pt(10)
        r0.font.color.rgb = ACCENT_BLUE
        
        p1 = c1.paragraphs[0]
        style_paragraph(p1, space_before=0, space_after=0, line_spacing=1.15)
        r1 = p1.add_run(v)
        r1.font.name = "Microsoft YaHei"
        r1._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r1.font.size = Pt(10)
        r1.font.color.rgb = TEXT_DARK
        
    # Page break after Cover (Requirement 9: Every chapter starts on fresh page)
    doc.add_page_break()

    # -------------------------------------------------------------
    # SECTION 2: FRONT MATTER (Table of Contents & List of Figures)
    # -------------------------------------------------------------
    # Setup Header and Footer for Document
    header = doc.sections[0].header
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    style_paragraph(hp, space_before=0, space_after=0)
    hrun = hp.add_run("航运行政人员 Excel 基础资料分析：资料前处理教学手册")
    hrun.font.name = "Microsoft YaHei"
    hrun._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    hrun.font.size = Pt(8.5)
    hrun.font.color.rgb = RGBColor(128, 128, 128)

    footer = doc.sections[0].footer
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    style_paragraph(fp, space_before=0, space_after=0)
    frun1 = fp.add_run("第 ")
    frun1.font.name = "Microsoft YaHei"
    frun1.font.size = Pt(9)
    frun1.font.color.rgb = RGBColor(128, 128, 128)
    add_page_number_to_run(frun1)
    frun2 = fp.add_run(" 页，共 ")
    frun2.font.name = "Microsoft YaHei"
    frun2.font.size = Pt(9)
    frun2.font.color.rgb = RGBColor(128, 128, 128)
    add_numpages_to_run(frun2)
    frun3 = fp.add_run(" 页")
    frun3.font.name = "Microsoft YaHei"
    frun3.font.size = Pt(9)
    frun3.font.color.rgb = RGBColor(128, 128, 128)

    # Table of Contents
    p_toc_h = doc.add_paragraph()
    p_toc_h.paragraph_format.keep_with_next = True
    r_th = p_toc_h.add_run("目  录 (Table of Contents)")
    r_th.bold = True
    r_th.font.name = "Microsoft YaHei"
    r_th._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    r_th.font.size = Pt(16)
    r_th.font.color.rgb = ACCENT_BLUE
    
    # Built-in Word TOC field
    p_field = doc.add_paragraph()
    fldSimple_toc = parse_xml(r'<w:fldSimple %s w:instr="TOC \o &quot;1-3&quot; \h \z \u"/>' % nsdecls('w'))
    p_field._p.append(fldSimple_toc)
    
    # Formatted visual table of contents for immediate reading
    toc_data = [
        ("第 1 章 导言与 Excel 操作介面全景解析", "1", True),
        ("    1.1 航运行政与数据分析的数字化转型背景", "1", False),
        ("    1.2 认识 MS Excel 经典操作介面全貌", "1", False),
        ("    1.3 核心区域逐项深度解析（功能区、编辑列与网格）", "2", False),
        ("    1.4 初学者必须掌握的滑鼠手势与高频快捷键速查", "3", False),
        ("第 2 章 航运资料前处理的核心观念与必备常识", "4", True),
        ("    2.1 资料前处理的定义与 GIGO 原则", "4", False),
        ("    2.2 Excel 的五大资料类型及航运常见认知误区", "5", False),
        ("    2.3 整洁资料规范（Tidy Data 原则在航运之落地）", "6", False),
        ("    2.4 航运行政典型“脏污资料”形态特征与根源", "7", False),
        ("第 3 章 30 笔航运实务资料档：资料前处理全流程示范", "9", True),
        ("    3.1 30 笔原始混乱航运资料档设计说明与展示", "9", False),
        ("    3.2 资料诊断清册与前处理方案拟定", "11", False),
        ("    3.3 资料前处理全流程演练（分步手把手实操操作）", "12", False),
        ("    3.4 清洗后 30 笔标准规范化资料对照表展示", "15", False),
        ("第 4 章 资料前处理关键功能选项与函数公式深度解构", "17", True),
        ("    4.1 核心功能选项操作指南（资料剖析、去重、验证等）", "17", False),
        ("    4.2 核心清洗函数公式语法与实战范式（TRIM、DATE、XLOOKUP等）", "20", False),
        ("    4.3 航运行政多函数嵌套综合应用范式", "24", False),
        ("第 5 章 300 笔多场景航运综合实战练习档与完整解题指南", "26", True),
        ("    5.1 300 笔航运案例练习档设计理念与业务场景分布", "26", False),
        ("    5.2 练习档结构与基础参照代码表（UN/LOCODE、船东代码）", "27", False),
        ("    5.3 十大核心考核任务挑战清单", "28", False),
        ("    5.4 完整解题流程、操作步骤、函数公式与标准答案", "30", False),
        ("第 6 章 航运数据品质长效管控与自动化进阶建议", "36", True),
        ("    6.1 源头表单设计防错机制", "36", False),
        ("    6.2 进阶自动化工具展望（Power Query 与 Python in Excel）", "37", False),
        ("    6.3 航运行政日常数据处理十条铁律", "38", False),
        ("附录与参考文献", "39", True),
        ("    附录 A：常用航运业务英文缩写与 UN/LOCODE 港口代码简表", "39", False),
        ("    附录 B：Excel 资料前处理高频快捷键速查表", "40", False),
        ("    参考文献与引证来源", "41", False)
    ]
    
    tbl_vtoc = doc.add_table(rows=len(toc_data), cols=2)
    tbl_vtoc.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_vtoc.autofit = False
    for r_idx, (title, page_str, is_ch) in enumerate(toc_data):
        c0 = tbl_vtoc.cell(r_idx, 0)
        c1 = tbl_vtoc.cell(r_idx, 1)
        c0.width = Inches(5.47)
        c1.width = Inches(0.8)
        set_cell_margins(c0, top=30, bottom=30, left=50, right=50)
        set_cell_margins(c1, top=30, bottom=30, left=50, right=50)
        
        p0 = c0.paragraphs[0]
        style_paragraph(p0, space_before=0, space_after=0, line_spacing=1.15)
        run0 = p0.add_run(title)
        run0.bold = is_ch
        run0.font.name = "Microsoft YaHei"
        run0._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        run0.font.size = Pt(9.5 if is_ch else 9)
        run0.font.color.rgb = ACCENT_BLUE if is_ch else TEXT_DARK
        
        p1 = c1.paragraphs[0]
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        style_paragraph(p1, space_before=0, space_after=0, line_spacing=1.15)
        run1 = p1.add_run(page_str)
        run1.bold = is_ch
        run1.font.name = "Microsoft YaHei"
        run1._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        run1.font.size = Pt(9.5 if is_ch else 9)
        run1.font.color.rgb = ACCENT_BLUE if is_ch else TEXT_DARK

    # Space and List of Figures & Tables
    p_lof_h = doc.add_paragraph()
    p_lof_h.paragraph_format.space_before = Pt(16)
    p_lof_h.paragraph_format.keep_with_next = True
    r_lfh = p_lof_h.add_run("图表目录 (List of Figures & Tables)")
    r_lfh.bold = True
    r_lfh.font.name = "Microsoft YaHei"
    r_lfh._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
    r_lfh.font.size = Pt(13)
    r_lfh.font.color.rgb = ACCENT_BLUE
    
    fig_table_entries = [
        ("图 1-1：MS Excel 经典操作介面全景与核心区域解析", "第 1 章", "2"),
        ("图 2-1：常见航运混乱资料结构 (Messy Data) 与整洁数据 (Tidy Data) 对比图", "第 2 章", "6"),
        ("图 4-1：Excel「资料剖析 (Text to Columns)」精灵对话框操作步骤图", "第 4 章", "18"),
        ("图 4-2：Excel「移除重复项」与「资料验证」设定视窗", "第 4 章", "19"),
        ("图 4-3：Excel 航运数据清洗核心函数公式联动语法示范图", "第 4 章", "21"),
        ("图 5-1：航运实务 300 笔情境资料前处理分析全景管线流程图", "第 5 章", "26"),
        ("表 1-1：初学者必备 Excel 高频操作快捷键清册", "第 1 章", "3"),
        ("表 2-1：Excel 五大资料类型在航运单证中的对应与常见错误", "第 2 章", "5"),
        ("表 3-1：30 笔航运实务原始混乱资料明细表", "第 3 章", "9"),
        ("表 3-2：30 笔原始资料诊断与清洗方案规划表", "第 3 章", "11"),
        ("表 3-3：清洗后 30 笔标准规范化资料对照表", "第 3 章", "15"),
        ("表 4-1：航运前处理十大核心函数速查表", "第 4 章", "20"),
        ("表 5-1：300 笔情境练习档业务场景与栏位定义表", "第 5 章", "27"),
        ("表 5-2：基础代码参照表（港口 UN/LOCODE 与船东代码）", "第 5 章", "28"),
        ("表 5-3：十大核心考核任务挑战清册与标准解法对齐表", "第 5 章", "29")
    ]
    
    tbl_figs = doc.add_table(rows=len(fig_table_entries), cols=3)
    tbl_figs.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_figs.autofit = False
    for r_idx, (title, ch, page_s) in enumerate(fig_table_entries):
        c0 = tbl_figs.cell(r_idx, 0)
        c1 = tbl_figs.cell(r_idx, 1)
        c2 = tbl_figs.cell(r_idx, 2)
        c0.width = Inches(4.77)
        c1.width = Inches(0.9)
        c2.width = Inches(0.6)
        set_cell_margins(c0, top=25, bottom=25, left=40, right=40)
        set_cell_margins(c1, top=25, bottom=25, left=40, right=40)
        set_cell_margins(c2, top=25, bottom=25, left=40, right=40)
        
        for c, text, align in [(c0, title, WD_ALIGN_PARAGRAPH.LEFT), (c1, ch, WD_ALIGN_PARAGRAPH.CENTER), (c2, page_s, WD_ALIGN_PARAGRAPH.RIGHT)]:
            p = c.paragraphs[0]
            p.alignment = align
            style_paragraph(p, space_before=0, space_after=0, line_spacing=1.1)
            r = p.add_run(text)
            r.font.name = "Microsoft YaHei"
            r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
            r.font.size = Pt(8.5)
            r.font.color.rgb = TEXT_DARK
            
    # Page break to start Chapter 1 (Requirement 9: Every chapter starts on fresh page)
    doc.add_page_break()

    # -------------------------------------------------------------
    # CHAPTER 1: INTRODUCTION & EXCEL UI
    # -------------------------------------------------------------
    add_custom_heading("第 1 章 导言与 Excel 操作介面全景解析", level=1)
    
    add_custom_heading("1.1 航运行政与数据分析的数字化转型背景", level=2)
    add_body_paragraph("国际航运物流作为全球贸易的实体血脉，其日常业务涉及海量单证流转与信息交互。航运行政人员每天需要处理大量来自船公司、货运代理、海关系统、码头堆场管理系统（TOS）以及企业内部 ERP/TMS 导出的资料档案。然而，不同信息源系统接口标准各异，人工录入习惯不一，导致导出的 Excel 表格普遍充斥着空白行、不可见字符、格式错乱、提单与货柜信息重复录入等严重问题。", indent=True)
    add_body_paragraph("对于航运初学者而言，若直接对这批未经整理的「脏资料」套用高级分析公式、建立数据透视表（Pivot Table）或绘制统计图表，极易得出完全失真的统计结论，甚至导致漏放行、错算滞期费、报关延误等严重实务事故。因此，扎实掌握基础资料前处理（Data Preprocessing）操作技能，是每一位航运行政人员进阶成为专业分析型人才的立足基石。", indent=True)
    
    add_custom_heading("1.2 认识 MS Excel 经典操作介面全貌", level=2)
    add_body_paragraph("Microsoft Excel 作为航运行业普及率最高的数据分析工具，其设计遵循了现代 Office 视窗软体的标准人体工程学布局。为了高效开展资料清洗作业，初学者首先必须从空间布局上完整建立对 Excel 界面的空间认知框架。下图直观标注了 Excel 各大核心组件的实际分布与职能定义：", indent=True)
    
    add_styled_image("images/fig_1_1_excel_interface_overview.png", "图 1-1：MS Excel 经典操作介面全景与核心区域解析", width_in_inches=6.0)
    
    add_custom_heading("1.3 核心区域逐项深度解析", level=2)
    add_bullet_point("位于工作视窗顶端。正中央清晰标注当前开启的工作簿档名（如「航运资料分析.xlsx」），左侧为「快速存取工具列」，右侧为视窗控制钮。行政人员在日常高强度处理多份提单时，务必随时关注档案保存状态。", "1. 标题列 (Title Bar)：")
    add_bullet_point("预设包含「储存 (Ctrl+S)」、「复原 (Ctrl+Z)」与「重做 (Ctrl+Y)」。强力建议航运行政人员通过右键自订，将高频使用的「快速填入」、「移除重复项」添加到此工具列中，实现一键触发。", "2. 快速存取工具列 (Quick Access Toolbar)：")
    add_bullet_point("Excel 核心命令中枢，采「标签页 (Tabs) -> 命令群组 (Groups) -> 功能按键 (Commands)」的三级树状结构。在资料前处理过程中，最核心的标签是「资料 (Data)」标签，其下辖的「资料工具」群组集成了【资料剖析】、【移除重复项】、【资料验证】等前处理王牌工具。", "3. 功能区标签与群组 (Ribbon Tabs & Groups)：")
    add_bullet_point("位于功能区正下方左侧。即时显示当前活动储存档的座标座标（如 C4、POL_List）。名称框不仅可用于快速定位跳跃至数十万行的大表底部，还可为特定港口代码区域命名「已命名范围」，大幅提升公式可读性。", "4. 名称框 (Name Box)：")
    add_bullet_point("紧邻名称框右侧，包含「插入函数 (fx)」按键与宽幅文字输入框。资料编辑列展示的是储存档底层的【真实内容或公式】，而非表面显示结果。例如单元格表面显示「2026-11-05」，但在编辑列中能精准识破其究竟是系统真日期序列值，还是含有隐藏空格的假文本。", "5. 资料编辑列 (Formula Bar)：")
    add_bullet_point("由字母栏（A, B, C... 国际标准最多 16,384 栏）与数字列（1, 2, 3... 最多 1,048,576 列）交织而成的二阶网格矩阵。每一个交汇点即为储存档（Cell），是数据存储的最小原子单元。", "6. 工作表网格 (Worksheet Grid)：")
    add_bullet_point("视窗最底端。左侧显示工作表卷标（Sheet Tabs），可右键进行色彩标记、重命名与移动；右侧显示当前选取区域的即时快速统计（平均值、计数、数值计数、求和）以及视窗缩放滑杆。", "7. 状态列与检视控制 (Status Bar)：")

    add_custom_heading("1.4 初学者必须掌握的滑鼠手势与高频快捷键速查", level=2)
    add_body_paragraph("熟练的键盘操作不仅能将日常行政处理效率提升 3 到 5 倍，更是避免手工拖拽失误的关键手段。以下列出航运行政前处理中最不可或缺的高频操作清册：", indent=True)
    
    # Table of Shortcuts
    tbl_sc = doc.add_table(rows=7, cols=3)
    tbl_sc.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_sc.autofit = False
    headers_sc = ["快捷键组合", "功能说明", "航运行政实务经典应用情境"]
    for c_idx, h_text in enumerate(headers_sc):
        cell = tbl_sc.cell(0, c_idx)
        cell.width = Inches(1.8 if c_idx == 0 else (2.0 if c_idx == 1 else 2.47))
        set_cell_background(cell, NAVY_HEX)
        set_cell_margins(cell, top=80, bottom=80, left=100, right=100)
        p = cell.paragraphs[0]
        style_paragraph(p, space_before=0, space_after=0)
        r = p.add_run(h_text)
        r.bold = True
        r.font.name = "Microsoft YaHei"
        r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    shortcuts_data = [
        ("Ctrl + 方向键 (↑ ↓ ← →)", "快速跳跃至当前连续数据区域的边缘边界", "快速从第 1 笔提单瞬间跳至第 10,000 笔资料底部，无需滚轮反复翻页"),
        ("Ctrl + Shift + 方向键", "从当前储存档连续扩增选取至边缘储存档", "一秒全选整列「海运费」或整列「集装箱号」进行批量格式化"),
        ("Ctrl + E", "快速填入 (Flash Fill) 模式识别与智慧提取", "从混合文本「COSCO GUANGZHOU/042W」中自动抽离航次号"),
        ("Ctrl + H", "开启「寻找及取代」对话框", "批量剔除数值栏位中的美元符号 '$'、逗号 ',' 或特定港口缩写"),
        ("Ctrl + 1", "开启「储存档格式 (Format Cells)」深度对话视窗", "将杂乱纯数字转换为标准日期格式 YYYY-MM-DD 或货币千分位"),
        ("F4 键", "绝对引用切换 ($A$1) 或重复上一次执行的格式操作", "在跨表 VLOOKUP/XLOOKUP 匹配港口代码时快速锁定基准对照表")
    ]
    for r_idx, (k, desc, scen) in enumerate(shortcuts_data, start=1):
        bg = "FFFFFF" if r_idx % 2 == 1 else "F9FAFC"
        for c_idx, val in enumerate([k, desc, scen]):
            cell = tbl_sc.cell(r_idx, c_idx)
            cell.width = Inches(1.8 if c_idx == 0 else (2.0 if c_idx == 1 else 2.47))
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=60, bottom=60, left=90, right=90)
            p = cell.paragraphs[0]
            style_paragraph(p, space_before=0, space_after=0)
            r = p.add_run(val)
            r.font.name = "Microsoft YaHei"
            r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
            r.font.size = Pt(9)
            r.font.color.rgb = ACCENT_BLUE if c_idx == 0 else TEXT_DARK
            if c_idx == 0:
                r.bold = True
                
    add_callout_box("航运初学者避坑指南", "切忌在包含上万条船期资料的大表中使用滑鼠左键拖拽选区！网络卡顿或滑鼠失误极易造成误选或拖拽覆盖原始数值。务必养成使用「Ctrl + Shift + ↓」进行专业键盘选取的肌肉记忆。", "TIP")

    # Page break to start Chapter 2 (Requirement 9: Every chapter starts on fresh page)
    doc.add_page_break()

    # -------------------------------------------------------------
    # CHAPTER 2: DATA PREPROCESSING PRINCIPLES & COMMON SENSE
    # -------------------------------------------------------------
    add_custom_heading("第 2 章 航运资料前处理的核心观念与必备常识", level=1)
    
    add_custom_heading("2.1 资料前处理的定义与 GIGO 原则", level=2)
    add_body_paragraph("在计算机科学与商业智能领域，存在一条著名的铁律——「垃圾进，垃圾出 (Garbage In, Garbage Out，简称 GIGO)」。这一定律在航运物流管理中体现得尤为淋漓尽致：如果输入分析系统的基础资料存在缺失、前后空格、重复或格式错误，无论后续使用的统计公式多么精妙、枢纽分析表多么复杂，最终产出的报告必定是充满误导性的虚假结果。", indent=True)
    add_body_paragraph("【资料前处理 (Data Preprocessing)】，是指在正式进行数据统计、汇总建模与商业可视化之前，对原始采集所得的原始数据执行系统性清洗（Cleaning）、规范转换（Transforming）、缺失值修复（Imputation）、格式标准化（Standardization）与结构规整（Restructuring）的完整技术流程。在实际航运管理业务中，航运行政与数据分析人员通常需要将 60% 至 80% 的工作时间投入在资料前处理阶段，它是确保航运供应链各环节精准协同的「第一道防火墙」。", indent=True)

    add_custom_heading("2.2 Excel 的五大资料类型及航运常见认知误区", level=2)
    add_body_paragraph("Excel 并非仅仅记录表面文字，其底层具备严格的底层数据类型判别系统。航运初学者最常犯的错误，就是将不同类型的数据混为一谈，从而导致公式计算报错（如常见的 #VALUE!）或排序混乱。下表详细剖析 Excel 五大基础资料类型在航运单证中的映射规则：", indent=True)
    
    # Table of Data Types
    tbl_dt = doc.add_table(rows=6, cols=4)
    tbl_dt.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_dt.autofit = False
    headers_dt = ["资料类型", "Excel 预设对齐", "航运单证对应示例", "初学者常见致命误区与后果"]
    widths_dt = [1.2, 1.1, 1.8, 2.17]
    for c_idx, h_text in enumerate(headers_dt):
        cell = tbl_dt.cell(0, c_idx)
        cell.width = Inches(widths_dt[c_idx])
        set_cell_background(cell, NAVY_HEX)
        set_cell_margins(cell, top=70, bottom=70, left=80, right=80)
        p = cell.paragraphs[0]
        style_paragraph(p, space_before=0, space_after=0)
        r = p.add_run(h_text)
        r.bold = True
        r.font.name = "Microsoft YaHei"
        r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    dt_rows = [
        ("文字 (Text)", "靠左对齐 (Left)", "提单号 COSU6291、箱型 40HC、港口代码 CNSHA", "将纯数字型提单号作为数值存储，导致前面的前导零 '0' 被系统自动吞噬抹除。"),
        ("数值 (Number)", "靠右对齐 (Right)", "货物毛重 18450、海运费 2450.00、计费吨 32.5", "在数值格中手工手打输入 '18450 kg' 或 '$2450'，导致数值蜕变为文本，SUM 无法求和。"),
        ("日期/时间 (Date)", "靠右对齐 (Right)", "预估开航日 ETD (2026-11-05)、截关时间 (12:00)", "输入 '2026.11.05' 或连续数字 '20261105'，系统判定为文本，无法进行日期相减算航程。"),
        ("逻辑值 (Boolean)", "置中对齐 (Center)", "TRUE (已放行)、FALSE (未放行/被海关扣押)", "将逻辑值输入为带有引号的文本 'TRUE'，导致 IF 条件判断逻辑失效。"),
        ("公式与错误值", "依结果而定", "=SUM(H2:H31)、错误值 #N/A、#VALUE!、#DIV/0!", "忽视报错隐患，将含错误值的表格直接提交，导致整个财务报销汇总全盘瘫痪。")
    ]
    for r_idx, row_vals in enumerate(dt_rows, start=1):
        bg = "FFFFFF" if r_idx % 2 == 1 else "F9FAFC"
        for c_idx, val in enumerate(row_vals):
            cell = tbl_dt.cell(r_idx, c_idx)
            cell.width = Inches(widths_dt[c_idx])
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            p = cell.paragraphs[0]
            style_paragraph(p, space_before=0, space_after=0)
            r = p.add_run(val)
            r.font.name = "Microsoft YaHei"
            r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
            r.font.size = Pt(8.5)
            r.font.color.rgb = ACCENT_BLUE if c_idx == 0 else TEXT_DARK
            if c_idx == 0:
                r.bold = True

    add_custom_heading("2.3 整洁资料规范（Tidy Data 原则在航运之落地）", level=2)
    add_body_paragraph("统计学家 Hadley Wickham 基于关系代数理论提出了著名的整洁数据（Tidy Data）哲学。在航运物流行政实务中，许多初学者习惯按照纸质手抄单证的排版习惯，把 Excel 当作「画板」使用：大量采用跨行跨列合并储存档、在数据行之间插入花哨的小计行和空白行、将货柜号与箱型强行塞在同一个格子里。这些被称为「混乱资料 (Messy Data)」，是自动化分析与数据透视表的头号杀手。", indent=True)
    
    add_styled_image("images/fig_2_1_tidy_vs_messy_data.png", "图 2-1：常见航运混乱资料结构 (Messy Data) 与整洁数据 (Tidy Data) 对比图", width_in_inches=6.0)
    
    add_body_paragraph("航运整洁资料的三大核心黄金法则是：", indent=True)
    add_bullet_point("每一列只能代表一种明确的航运业务属性。例如：装运港（POL）独占一列，目的港（POD）独占一列，绝对不可将「起运-目的」拼成「CNSHA-DEHAM」合在一个单元格中。", "规则一：每一列 (Column) 必须代表唯一的变数：")
    add_bullet_point("每一行只能代表单次独立的物流运输事件或观测主体。例如：每一行记录唯一一只集装箱的流转动态，或者记录一票独立的海运提单。", "规则二：每一行 (Row) 必须代表单次具体的观测纪录：")
    add_bullet_point("严禁在资料中间插入「小计」或「空白隔开行」；数据汇总应当在独立的透视表或专用汇总报表中执行，底层明细表必须保持纯粹的连续矩形网格结构。", "规则三：表头单一规范，杜绝跨行合并格与中间小计行：")

    add_custom_heading("2.4 航运行政典型“脏污资料”形态特征与根源", level=2)
    add_body_paragraph("在日常航运单证流转中，数据脏污通常源于以下四个典型场景，初学者必须具备敏锐的识别眼力：", indent=True)
    add_bullet_point("从船东订舱网站、海关单一窗口或旧版 EDI 报文中复制提单信息时，常常夹带大量不可见的 ASCII 控制字符（如换行符 Chr(10)、回车符 Chr(13) 以及不间断空格 Chr(160)）。这些字符肉眼难以察觉，却会导致 VLOOKUP 匹配完全失败。", "1. 隐形空格与不可见控制字符污染：")
    add_bullet_point("不同航线操作人员习惯差异大，例如同一个上海港，有人输入「CNSHA」，有人输入小写「cnsha」，甚至输入「SHA」、「shanghai」；同理，提单号与船名大小写混杂，阻碍了精准聚合汇总。", "2. 编码与大小写未统一规范：")
    add_bullet_point("海关 EDI、船代订舱确认书与海外代理给出的日期格式五花八门：包括标准 ISO「2026-11-05」、美式斜杠「11/05/2026」、欧式「05/11/2026」、带点「2026.11.05」以及无分隔符纯文本「20261105」。混杂日期无法进行逾期天数核算。", "3. 历法与日期书写格式混杂错乱：")
    add_bullet_point("货运代理在拼箱操作或多次更新舱单时，同一票提单或集装箱号可能被重复录入；部分行政人员因防范意识不足，将重复记录重复计入运费支出，带来财务亏损风险。", "4. 重复录入与幽灵纪录：")

    # Page break to start Chapter 3 (Requirement 9: Every chapter starts on fresh page)
    doc.add_page_break()

    # -------------------------------------------------------------
    # CHAPTER 3: 30-RECORD DEMONSTRATION WORKFLOW
    # -------------------------------------------------------------
    add_custom_heading("第 3 章 30 笔航运实务资料档：资料前处理全流程示范", level=1)
    
    add_custom_heading("3.1 30 笔原始混乱航运资料档设计说明与展示", level=2)
    add_body_paragraph("为了完整重现航运物流现场真实的资料污染现状，本章特意设计了一份典型的 30 笔综合海运业务原始资料档（存放在伴随教学档 `航运资料前处理_30笔示范档.xlsx` 中）。该资料集浓缩包含了国际集装箱海运中最常见的典型错误：前后冗余空格、重复提单与集装箱、带文本单位的重量与金额、大小写混杂的联合国港口代码（UN/LOCODE）、非标日期文本以及复合栏位。原始资料明细如下表所示：", indent=True)
    
    # 30-row Raw Table (Sample excerpted or full)
    tbl_raw30 = doc.add_table(rows=31, cols=6)
    tbl_raw30.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_raw30.autofit = False
    
    raw_sub_headers = ["序号", "原始提单号 (B/L No.)", "原始集装箱号", "船名航次 (混合栏)", "港口 (POL/POD)", "原始运费金额 (混杂)"]
    raw_sub_w = [0.6, 1.4, 1.3, 1.6, 1.1, 1.27]
    for c_idx, h_t in enumerate(raw_sub_headers):
        cell = tbl_raw30.cell(0, c_idx)
        cell.width = Inches(raw_sub_w[c_idx])
        set_cell_background(cell, NAVY_HEX)
        set_cell_margins(cell, top=60, bottom=60, left=50, right=50)
        p = cell.paragraphs[0]
        style_paragraph(p, space_before=0, space_after=0)
        r = p.add_run(h_t)
        r.bold = True
        r.font.name = "Microsoft YaHei"
        r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    # Read the 30 raw data we created earlier
    import openpyxl
    wb_read30 = openpyxl.load_workbook("data/航运资料前处理_30笔示范档.xlsx")
    ws_r30 = wb_read30["30笔原始脏资料"]
    
    row_count = 1
    for r in ws_r30.iter_rows(min_row=2, max_row=31, values_only=True):
        idx_val = str(r[0])
        bl_val = str(r[1])
        cntr_val = str(r[2])
        vessel_val = str(r[4])
        port_val = f"{str(r[5]).strip()}/{str(r[6]).strip()}"
        frt_val = str(r[10])
        
        is_dup_row = row_count in [4, 21]
        bg_c = "FFE2E2" if is_dup_row else ("FFFFFF" if row_count % 2 == 1 else "F9FAFC")
        
        for c_idx, text_val in enumerate([idx_val, bl_val, cntr_val, vessel_val, port_val, frt_val]):
            cell = tbl_raw30.cell(row_count, c_idx)
            cell.width = Inches(raw_sub_w[c_idx])
            set_cell_background(cell, bg_c)
            set_cell_margins(cell, top=40, bottom=40, left=50, right=50)
            p = cell.paragraphs[0]
            style_paragraph(p, space_before=0, space_after=0)
            r_run = p.add_run(text_val + (" (重复)" if (is_dup_row and c_idx == 1) else ""))
            r_run.font.name = "Microsoft YaHei"
            r_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
            r_run.font.size = Pt(8)
            r_run.font.color.rgb = RGBColor(192, 0, 0) if is_dup_row else TEXT_DARK
            if is_dup_row:
                r_run.bold = True
        row_count += 1

    add_body_paragraph("注：上表中粉红色高亮行（第 4 行与第 21 行）为典型的人工重复录入错误，其集装箱号与提单号与先前的记录完全重合，必须在前处理流程中被准确剔除。", indent=True)

    add_custom_heading("3.2 资料诊断清册与前处理方案拟定", level=2)
    add_body_paragraph("在盲目动手修改数据之前，专业的数据分析师必须对数据源展开「全面体检」，拟定标准化的前处理路线图。针对上述 30 笔原始资料，我们制定了如下所示的诊断与清洗方案清册：", indent=True)
    
    # Table of Diagnostics & Actions
    tbl_diag = doc.add_table(rows=7, cols=4)
    tbl_diag.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_diag.autofit = False
    headers_dg = ["目标栏位", "现存脏污问题诊断", "预定前处理手段与工具", "清洗后预期标准成果"]
    w_dg = [1.2, 1.8, 1.8, 1.47]
    for c_idx, h_t in enumerate(headers_dg):
        cell = tbl_diag.cell(0, c_idx)
        cell.width = Inches(w_dg[c_idx])
        set_cell_background(cell, NAVY_HEX)
        set_cell_margins(cell, top=70, bottom=70, left=70, right=70)
        p = cell.paragraphs[0]
        style_paragraph(p, space_before=0, space_after=0)
        r = p.add_run(h_t)
        r.bold = True
        r.font.name = "Microsoft YaHei"
        r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    diag_data = [
        ("提单号 / 箱号", "首尾带有冗余空格，英文字母大小写混杂，存在整行重复录入", "使用 TRIM+CLEAN+UPPER 函数嵌套清洗，执行「移除重复项」功能", "纯净无空格大写标准编码（如 COSU62918401），去重后保留 28 笔唯一记录"),
        ("船名航次", "船名与航次挤在同一储存档内，以斜杠 '/' 相隔，不便单独统计", "使用 Excel 功能区【资料剖析】按斜杠拆分，配合 TRIM 清除多余空格", "独立分列为「船名 (Vessel)」与「航次 (Voyage)」两个标准单值变量"),
        ("装运港 / 目的港", "存在小写代码（cnsha）、前后空格，缺乏直观中文全称", "UPPER 函数大写化，配合 XLOOKUP 跨表匹配《基础代码参照表》", "标准 5 位大写 UN/LOCODE 代码，并自动补全「上海港 (Shanghai)」等全称"),
        ("预估开航/抵港日", "包含斜杠 '2026/11/05'、点号 '2026.11.05' 及无分隔文本 '20261108'", "使用 SUBSTITUTE 统一替换符号，配合 DATEVALUE 转为真日期序列", "统一标准 ISO-8601 日期格式 YYYY-MM-DD，支持计算航程天数 (ETA-ETD)"),
        ("货物毛重 (KGS)", "包含千分位逗号 ',' 与文字单位 'kg'、'kgs'，无法进行数学求和", "SUBSTITUTE 剥离单位字母与逗号，嵌套 VALUE 函数转换为纯数字", "纯数字类型重量，设定储存档数值格式为千分位 '#,##0'，单位标于表头"),
        ("海运费 (USD)", "夹带美元符号 '$'、文本 'USD' 与千分位，属于伪数值文本", "函数提取纯数值，或使用【快速填入 (Ctrl+E)】剥离金额纯数字", "纯数值型货币，设定储存档格式为 '$#,##0.00'，能够正常执行 SUM 汇总")
    ]
    for r_idx, row_vals in enumerate(diag_data, start=1):
        bg = "FFFFFF" if r_idx % 2 == 1 else "F9FAFC"
        for c_idx, val in enumerate(row_vals):
            cell = tbl_diag.cell(r_idx, c_idx)
            cell.width = Inches(w_dg[c_idx])
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=50, bottom=50, left=70, right=70)
            p = cell.paragraphs[0]
            style_paragraph(p, space_before=0, space_after=0)
            r = p.add_run(val)
            r.font.name = "Microsoft YaHei"
            r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
            r.font.size = Pt(8.5)
            r.font.color.rgb = ACCENT_BLUE if c_idx == 0 else TEXT_DARK
            if c_idx == 0:
                r.bold = True

    add_custom_heading("3.3 资料前处理全流程演练（分步手把手实操操作）", level=2)
    add_body_paragraph("以下遵循标准化数据工程流程，分步演示在 MS Excel 中处理这批 30 笔资料的完整动作：", indent=True)
    
    add_bullet_point("在处理任何业务档案前，务必先另存备份或复制一份原始工作表（命名为「原始资料备份」）。绝对不可直接在未经备份的原始数据上执行不可逆的大范围替换与删除操作！", "步骤一：建立副本与防护备份。")
    add_bullet_point("全选表格区域（A1:L31），点击功能区【资料】->【移除重复项】。在弹出的视窗中，确保勾选「我的资料具有标题」，取消全选后仅勾选「提单号」与「集装箱号」。点击确定后，Excel 弹出提示「已找到并移除 2 个重复值，保留 28 个唯一值」。原表格中第 4 行与第 21 行重复记录被精准消除。", "步骤二：执行唯一主键去重。")
    add_bullet_point("在原始提单号旁边插入辅助列，输入核心公式：`=TRIM(CLEAN(UPPER(B2)))`。下拉复制填充至第 28 行。该公式一举完成了三项使命：CLEAN 剥离系统导出隐藏控制符、TRIM 清除提单前后误触打出的空格、UPPER 将小写字母统一转换为国际标准大写代码。", "步骤三：规范文本与大小写清洗。")
    add_bullet_point("选中船名航次列（E 列），点击【资料】->【资料剖析】。第一步选择「分隔符号」点击下一步；第二步在「其他」输入框中输入斜杠 `/`，勾选「将连续分界符号视为单一符号」；第三步目标储存档指定为拆分后的空白列，点击完成。原「EVER GIVEN / 0118E」瞬间整洁拆分为独立的「EVER GIVEN」与「0118E」两栏。", "步骤四：利用资料剖析拆分复合栏位。")
    add_bullet_point("将非标文本日期规范转换为系统真日期。在空白列输入公式：`=DATEVALUE(SUBSTITUTE(SUBSTITUTE(H2, \".\", \"-\"), \"/\", \"-\"))`。对于无分隔符的 8 位纯数字日期（如 20261108），输入公式 `=DATE(LEFT(H2,4), MID(H2,5,2), RIGHT(H2,2))`。最后统一将单元格格式设置为「短日期 (YYYY-MM-DD)」。", "步骤五：日期格式标准化与时间真伪识别。")
    add_bullet_point("在货重清洗列输入公式：`=VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(LOWER(J2), \"kgs\", \"\"), \"kg\", \"\"), \",\", \"\"))`。通过多次替换剥离一切非数字文本并转换为纯数字。选中整列，按快捷键 `Ctrl + Shift + 1` 格式化为标准千分位。", "步骤六：数值剥离与金额净化。")
    add_bullet_point("在新建的「装运港中文名称」栏位输入公式：`=XLOOKUP(TRIM(UPPER(F2)), 基础代码参照表!$B$2:$B$19, 基础代码参照表!$C$2:$C$19, \"未知港口\")`。公式自动将港口代码（如 CNSHA）翻译匹配为「上海港 (Shanghai)」，大幅提升后勤部门调度与对外沟通的可读性。", "步骤七：跨表关联与中文码表补全。")
    add_bullet_point("在完成了所有辅助列的公式清洗后，选中所有清洗后的标准栏位，复制并在新工作表中右键选择【选择性贴上】->【仅贴上数值 (Values)】。这一步至关重要，它切断了对脆弱临时公式的依赖，将数据固化为纯净实体数据，防止未来公式源引用变更导致错误。", "步骤八：选择性贴上固化纯净值。")

    add_custom_heading("3.4 清洗后 30 笔标准规范化资料对照表展示", level=2)
    add_body_paragraph("经过上述标准化流水线清洗后，原始包含脏污错误的资料完全蜕变为结构严谨、字段单一、无任何杂质的整洁航运资料表。下表展示了去重并规范化后的标准前 15 笔资料明细：", indent=True)
    
    # Clean Table Excerpt
    tbl_clean30 = doc.add_table(rows=16, cols=6)
    tbl_clean30.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_clean30.autofit = False
    
    cln_headers = ["标准序号", "规范提单号", "标准集装箱号", "船名 / 航次", "挂靠港口航线", "规范海运费 (USD)"]
    cln_w = [0.7, 1.4, 1.3, 1.6, 1.2, 1.07]
    for c_idx, h_t in enumerate(cln_headers):
        cell = tbl_clean30.cell(0, c_idx)
        cell.width = Inches(cln_w[c_idx])
        set_cell_background(cell, DARK_GREEN_HEX)
        set_cell_margins(cell, top=60, bottom=60, left=50, right=50)
        p = cell.paragraphs[0]
        style_paragraph(p, space_before=0, space_after=0)
        r = p.add_run(h_t)
        r.bold = True
        r.font.name = "Microsoft YaHei"
        r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    ws_c30 = wb_read30["30笔清洗后标准资料"]
    row_count = 1
    for r in ws_c30.iter_rows(min_row=2, max_row=16, values_only=True):
        idx_val = str(r[0])
        bl_val = str(r[1])
        cntr_val = str(r[2])
        vessel_val = f"{str(r[4])} / {str(r[5])}"
        port_val = f"{str(r[6])} -> {str(r[8])}"
        frt_val = f"${float(r[14]):,.2f}"
        
        bg_c = "FFFFFF" if row_count % 2 == 1 else "F4FBF7"
        for c_idx, text_val in enumerate([idx_val, bl_val, cntr_val, vessel_val, port_val, frt_val]):
            cell = tbl_clean30.cell(row_count, c_idx)
            cell.width = Inches(cln_w[c_idx])
            set_cell_background(cell, bg_c)
            set_cell_margins(cell, top=40, bottom=40, left=50, right=50)
            p = cell.paragraphs[0]
            style_paragraph(p, space_before=0, space_after=0)
            r_run = p.add_run(text_val)
            r_run.font.name = "Microsoft YaHei"
            r_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
            r_run.font.size = Pt(8)
            r_run.font.color.rgb = TEXT_DARK
        row_count += 1

    add_body_paragraph("完整的 28 笔纯净标准资料与各阶段公式对照，可参阅随附提供的实操 Excel 活页簿 `航运资料前处理_30笔示范档.xlsx` 的第二个分页。", indent=True)

    # Page break to start Chapter 4 (Requirement 9: Every chapter starts on fresh page)
    doc.add_page_break()

    # -------------------------------------------------------------
    # CHAPTER 4: DEEP DIVE INTO EXCEL PREPROCESSING FEATURES & FORMULAS
    # -------------------------------------------------------------
    add_custom_heading("第 4 章 资料前处理关键功能选项与函数公式深度解构", level=1)
    
    add_custom_heading("4.1 核心功能选项操作指南", level=2)
    add_body_paragraph("Excel 在功能区中提供了数个强悍的前处理按键，熟练掌握它们能让行政人员在不写任何公式的情况下，瞬时化解复杂数据难题。", indent=True)
    
    add_custom_heading("4.1.1 资料剖析 (Text to Columns)", level=3)
    add_body_paragraph("【功能路径】：选定目标栏位 -> 点击功能区【资料】标签 ->【资料工具】群组 ->【资料剖析】。", indent=True)
    add_body_paragraph("资料剖析精灵包含两大核心模式：\n1. 「分隔符号 (Delimited)」：适用于栏位内容中含有明确分割标识（如逗号、分号、斜杠、Tab）的情境。在航运业务中，常用于拆分「船名/航次」、「起运港-目的港」或「长x宽x高 (L*W*H)」。\n2. 「固定宽带 (Fixed Width)」：适用于无分隔符但字符数恒定的情境。例如国际标准集装箱号（如 COSU9823145），前 4 位永远是所有者与设备代码，后 7 位是序列号与校验码，使用固定宽度可一键将箱主代号与箱号精准切开。", indent=True)
    
    add_styled_image("images/fig_4_1_text_to_columns.png", "图 4-1：Excel「资料剖析 (Text to Columns)」精灵对话框操作步骤图", width_in_inches=5.8)
    
    add_custom_heading("4.1.2 移除重复项 (Remove Duplicates) 与资料验证 (Data Validation)", level=3)
    add_body_paragraph("去重与录入校验是保证资料唯一性与合规性的两大法宝：", indent=True)
    
    add_styled_image("images/fig_4_2_remove_duplicates_and_validation.png", "图 4-2：Excel「移除重复项」与「资料验证」设定视窗", width_in_inches=6.0)
    
    add_bullet_point("【移除重复项技巧】：切勿盲目全选所有栏位去重！在航运实务中，两票不同的拼箱货物可能挂靠相同的起运港和目的港，若全选栏位会导致误删有效订单。正确的实务做法是仅勾选具备唯一商业主键性质的属性，如「提单号 (B/L No.)」与「集装箱号 (Container No.)」。", "1. 移除重复项的实务心法：")
    add_bullet_point("【资料验证防错机制】：事后清洗不如源头防错。通过【资料】->【资料验证】，将「允许」设置为「清单 (List)」，在来源框输入 `=基础代码参照表!$B$2:$B$19`。此时单元格将自动生成下拉选单，操作员只能在预设的标准 UN/LOCODE 中挑选，彻底根除手动输入造成的大小写不一与拼写错漏！", "2. 资料验证构建防错输入法：")
    add_bullet_point("【快速填入 (Ctrl+E)】：Excel 基于启发式 AI 演算法提供的神奇工具。在紧邻原始资料的空白列手打输入 1~2 个预期的清洗样本，按下 `Ctrl + E`，Excel 会瞬间领悟提取规律并自动补全全列数万行数据！", "3. 快速填入 (Flash Fill) 极速模式识别：")
    add_bullet_point("【定位条件 (Go To Special)】：按 `F5` 键或 `Ctrl + G` 打开定位，点击「特殊」-> 选择「空格 (Blanks)」。该功能可在一秒钟内将庞大表格中所有缺失未填的单元格全部高亮，便于批量填补「未知」或填充上一行有效值。", "4. 定位条件批量捕获空值：")

    add_custom_heading("4.2 核心清洗函数公式语法与实战范式", level=2)
    add_body_paragraph("对于需要动态更新、反覆执行的复杂前处理作业，函数公式提供了无与伦比的灵活性与自动化能力。以下系统解构航运前处理最重要的十大核心函数：", indent=True)
    
    # Table of Core Formulas
    tbl_fm = doc.add_table(rows=11, cols=4)
    tbl_fm.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_fm.autofit = False
    headers_fm = ["函数名称", "标准语法格式", "航运数据前处理核心功能", "实战经典公式示范"]
    w_fm = [1.1, 1.6, 1.8, 1.77]
    for c_idx, h_t in enumerate(headers_fm):
        cell = tbl_fm.cell(0, c_idx)
        cell.width = Inches(w_fm[c_idx])
        set_cell_background(cell, NAVY_HEX)
        set_cell_margins(cell, top=60, bottom=60, left=70, right=70)
        p = cell.paragraphs[0]
        style_paragraph(p, space_before=0, space_after=0)
        r = p.add_run(h_t)
        r.bold = True
        r.font.name = "Microsoft YaHei"
        r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    formulas_data = [
        ("TRIM", "=TRIM(text)", "清除文本中除单词间单个空格之外的所有前后多余空格", "=TRIM(B2) 清洗提单号两端误触的空格"),
        ("CLEAN", "=CLEAN(text)", "移除文本中所有 0~31 的不可打印控制字符", "=CLEAN(B2) 消除 EDI/网页导出的换行乱码"),
        ("UPPER", "=UPPER(text)", "将文本中的所有英文字母转换为全大写标准格式", "=UPPER(F2) 将小写 'cnsha' 纠正为 'CNSHA'"),
        ("TEXT", "=TEXT(value, format_text)", "将数值或日期转换为指定外观格式的文本字符串", "=TEXT(H2, \"YYYY-MM-DD\") 规范日期格式"),
        ("VALUE", "=VALUE(text)", "将代表数字的文本型伪数字转换为真正的可计算数值", "=VALUE(\"18450\") 消除绿色小三角文本标记"),
        ("LEFT / RIGHT", "=LEFT(text, [num_chars])", "从文本最左侧或最右侧提取指定位数的字符", "=LEFT(C2, 4) 提取集装箱箱主代码（如 MSKU）"),
        ("MID / FIND", "=MID(text, start, num)", "结合 FIND 动态截取文本中间特定区间的字符", "=MID(A2, FIND(\"/\", A2)+1, 10) 动态截取航次"),
        ("DATE / DATEVALUE", "=DATE(year, month, day)", "通过指定年月日三个整数合成构造一个合法系统真日期", "=DATE(LEFT(A2,4), MID(A2,5,2), RIGHT(A2,2))"),
        ("IFERROR", "=IFERROR(value, val_if_err)", "容错函数，捕获公式报错并优雅显示预设替代值", "=IFERROR(XLOOKUP(...), \"未建档港口\")"),
        ("XLOOKUP", "=XLOOKUP(lookup, in, ret)", "全能现代跨表查询函数，支援双向匹配且无死角", "=XLOOKUP(F2, 码表!A:A, 码表!B:B, \"查无代码\")")
    ]
    for r_idx, row_vals in enumerate(formulas_data, start=1):
        bg = "FFFFFF" if r_idx % 2 == 1 else "F9FAFC"
        for c_idx, val in enumerate(row_vals):
            cell = tbl_fm.cell(r_idx, c_idx)
            cell.width = Inches(w_fm[c_idx])
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=50, bottom=50, left=70, right=70)
            p = cell.paragraphs[0]
            style_paragraph(p, space_before=0, space_after=0)
            r = p.add_run(val)
            r.font.name = "Microsoft YaHei"
            r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
            r.font.size = Pt(8.5)
            r.font.color.rgb = ACCENT_BLUE if c_idx == 0 else TEXT_DARK
            if c_idx == 0:
                r.bold = True

    add_styled_image("images/fig_4_3_formulas_cleaning_showcase.png", "图 4-3：Excel 航运数据清洗核心函数公式联动语法示范图", width_in_inches=6.0)

    add_custom_heading("4.3 航运行政多函数嵌套综合应用范式", level=2)
    add_body_paragraph("在复杂的实务环境中，单一函数往往难以独自应对错综复杂的脏污状况。以下提供三个航运行政中最经典的万能嵌套公式模板，初学者可直接套用：", indent=True)
    
    add_callout_box("实务模板一：航运万能文字清洗终极嵌套", 
                    "=TRIM(CLEAN(UPPER(SUBSTITUTE(A2, CHAR(160), \"\"))))\n\n"
                    "【解析】：网页系统导出的表格中经常混入 ASCII 码为 160 的「不间断空格（NBSP）」，普通的 TRIM 函数无法识别该字符！该嵌套公式首先用 SUBSTITUTE 将所有 CHAR(160) 彻底清除，再用 CLEAN 清除隐藏控制符，接着用 TRIM 斩断所有头尾空格，最后用 UPPER 统一大写，是处理提单号与货柜号的「降魔神咒」。", "NOTE")

    add_callout_box("实务模板二：混杂文本日期的强力标准化转换", 
                    "=IF(ISNUMBER(A2), A2, DATEVALUE(SUBSTITUTE(SUBSTITUTE(A2, \".\", \"-\"), \"/\", \"-\")))\n\n"
                    "【解析】：先用 ISNUMBER 判定该单元格是否已经是系统真日期，如果是则保持原样；若不是真日期，则通过双层嵌套 SUBSTITUTE 将点号 '.' 和斜杠 '/' 一律替换为标准短横线 '-'，最后由 DATEVALUE 强制转为系统序列值并套用 YYYY-MM-DD 格式，轻松收编五花八门的日期格式。", "NOTE")

    add_callout_box("实务模板三：带货币符号与千分位混杂金额清洗", 
                    "=VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(UPPER(A2), \"USD\", \"\"), \"$\", \"\"), \",\", \"\"))\n\n"
                    "【解析】：针对如 '$ 2,450.00' 或 'USD 1,200' 这种被严重污染为文字的应收应付运费，多重剔除币别代码与千分位逗号，再用 VALUE 转为可运算数值，彻底根除求和时出现 #VALUE! 报错的隐患。", "NOTE")

    # Page break to start Chapter 5 (Requirement 9: Every chapter starts on fresh page)
    doc.add_page_break()

    # -------------------------------------------------------------
    # CHAPTER 5: 300-RECORD MULTI-SCENARIO PRACTICE & SOLUTION GUIDE
    # -------------------------------------------------------------
    add_custom_heading("第 5 章 300 笔多场景航运综合实战练习档与完整解题指南", level=1)
    
    add_custom_heading("5.1 300 笔航运案例练习档设计理念与业务场景分布", level=2)
    add_body_paragraph("经过前述章节的理论奠基与小规模演练，本章正式提供一份高拟真度的综合实操练习档案——《航运资料前处理_300笔情境练习档.xlsx》（存放于随书练习资料夹中）。该资料集严格遵循现代航运企业运营实务，覆盖了三大经典业务场景，各占 100 笔观测样本：", indent=True)
    
    add_bullet_point("涉及工作单号（Job No.）、主提单号（MBL）、集装箱号（Container No.）、ISO 箱型代码（20GP/40GP/40HC/45HC/冷藏箱 RF）、承运船东、封条完整性等，重点考察箱号格式规范、箱型大写清洗与主键去重。", "场景 A：集装箱装箱单流转与全程动态跟踪（Container Manifest & Tracking）：")
    add_bullet_point("涉及货物体积（CBM）、毛重（KGS）、计费吨（Revenue Ton, 简称 RT，规则为 MAX(CBM, KGS/1000)）、基础海运费、超期滞箱费（Demurrage）、滞港费（Detention）与应收总账，重点考察数值净化与跨币别核算。", "场景 B：海运运杂费核算与滞期费用结算（Freight & Demurrage Settlement）：")
    add_bullet_point("涉及船名航次剖析、装运港（POL）与目的港（POD）UN/LOCODE 转换、预估开航日（ETD）、实际离港日（ATD）、预估抵港日（ETA）、实际到港日（ATA）、航程天数计算与延误天数核定，重点考察复杂日期运算。", "场景 C：船期表准点率与港口泊位作业监控（Schedule Punctuality & Operations）：")

    add_styled_image("images/fig_5_1_practice_pipeline.png", "图 5-1：航运实务 300 笔情境资料前处理分析全景管线流程图", width_in_inches=6.0)

    add_custom_heading("5.2 练习档结构与基础参照代码表", level=2)
    add_body_paragraph("练习档案分为三大核心分页：\n1. 【300笔情境原始资料】：未经清洗的原始资料湖，包含各种人为制造的脏污瑕疵；\n2. 【基础代码参照表】：包含 19 个国际核心港口 UN/LOCODE 标准字典（含中英文港名与航线）以及全球 6 大集装箱航运巨头联盟代码（中远海运、马士基、达飞、长荣、ONE、地中海航运）；\n3. 【练习题说明与考核指标】：明确列出了 10 项具体清洗任务的评分量表与验收准则。", indent=True)

    add_custom_heading("5.3 十大核心考核任务挑战清单", level=2)
    add_body_paragraph("学员在打开练习档后，需按顺序达成以下十项前处理挑战指标：", indent=True)
    
    # Table of 10 challenges
    tbl_ch10 = doc.add_table(rows=11, cols=4)
    tbl_ch10.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_ch10.autofit = False
    headers_c10 = ["任务编号", "考核前处理模组", "涉及栏位范围", "核心考核目标要求与合格准则"]
    w_c10 = [0.8, 1.4, 1.5, 2.57]
    for c_idx, h_t in enumerate(headers_c10):
        cell = tbl_ch10.cell(0, c_idx)
        cell.width = Inches(w_c10[c_idx])
        set_cell_background(cell, NAVY_HEX)
        set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
        p = cell.paragraphs[0]
        style_paragraph(p, space_before=0, space_after=0)
        r = p.add_run(h_t)
        r.bold = True
        r.font.name = "Microsoft YaHei"
        r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    c10_rows = [
        ("任务 01", "多重主键精准去重", "工作单号 / MBL / 柜号", "识别并删除 3 笔完全伪造的重复订单，资料集行数由 300 笔精准收敛至 297 笔。"),
        ("任务 02", "单证标识文本净化", "主提单号 / 集装箱号", "清除前后所有空白字符，英文统一大写，集装箱号满足 4 位字母 + 7 位数字结构。"),
        ("任务 03", "箱型代码清洗与验证", "箱型 (ISO Type)", "清除尾部空格并大写化，建立资料验证下拉清单（20GP, 40GP, 40HC, 45HC, 20RF, 40RF）。"),
        ("任务 04", "复合船名航次分列", "船名航次 (Vessel/Voyage)", "使用【资料剖析】按斜杠拆分为独立的「船名」与「航次代码」两列，修剪两端空格。"),
        ("任务 05", "港口代码对齐与补全", "POL / POD 港口栏位", "将小写代码转为大写，跨表关联《基础代码参照表》自动补全中文港口名称。"),
        ("任务 06", "四大时间栏位标准化", "ETD / ATD / ETA / ATA", "将 4 种混杂日期格式全部归一化为标准真日期格式 YYYY-MM-DD。"),
        ("任务 07", "计划航程与延误核算", "计划航程 / 到港延误天数", "以公式计算计划航程天数 (ETA - ETD)；计算实际延误天数 MAX(0, ATA - ETA)。"),
        ("任务 08", "货重体积数值型净化", "体积 CBM / 毛重 KGS", "剔除文本单位，转为可计算纯数字，分别套用两位小数与千分位格式。"),
        ("任务 09", "计费吨与运杂费总核算", "计费吨 RT / 运费 / 滞箱费", "运用 MAX 规则计算计费吨；剥离货币字符，核算总应收金额，处理异常值。"),
        ("任务 10", "单证异常状态标记", "单证完整状态 / 异常标记", "运用 IF 与 OR 逻辑，将单证不全与滞期超过 2 天的柜子自动标红预警。")
    ]
    for r_idx, row_vals in enumerate(c10_rows, start=1):
        bg = "FFFFFF" if r_idx % 2 == 1 else "F9FAFC"
        for c_idx, val in enumerate(row_vals):
            cell = tbl_ch10.cell(r_idx, c_idx)
            cell.width = Inches(w_c10[c_idx])
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=45, bottom=45, left=60, right=60)
            p = cell.paragraphs[0]
            style_paragraph(p, space_before=0, space_after=0)
            r = p.add_run(val)
            r.font.name = "Microsoft YaHei"
            r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
            r.font.size = Pt(8)
            r.font.color.rgb = ACCENT_BLUE if c_idx == 0 else TEXT_DARK
            if c_idx == 0:
                r.bold = True

    add_custom_heading("5.4 完整解题流程、操作步骤、函数公式与标准答案", level=2)
    add_body_paragraph("为确保初学者能够独立自主复盘、核对解题成果，本节提供详尽的分步参考操作指南与标准答案判定准则：", indent=True)
    
    add_custom_heading("任务 01 解题步骤：多重主键去重", level=3)
    add_body_paragraph("【滑鼠操作】：选中数据区域 A1:V301 -> 点击【资料】->【移除重复项】-> 勾选「我的资料具有标题」-> 勾选「工作单号」与「主提单号」-> 点击确定。Excel 弹出提示成功移除 3 个重复值，剩余 297 行有效记录。", indent=True)
    add_body_paragraph("【参考答案校验】：检查第 41、121、211 行数据是否已自动顺延下移，总行数变为 298 行（含表头）。", indent=True)

    add_custom_heading("任务 02 & 03 解题步骤：单证与箱型文本净化", level=3)
    add_body_paragraph("【公式应用】：在空白列输入 `=TRIM(CLEAN(UPPER(D2)))` 针对主提单号；针对箱型输入 `=TRIM(UPPER(F2))`。双击填充柄完成全列计算。\n【建立资料验证】：选中箱型列 -> 点击【资料】->【资料验证】-> 允许选择「清单」-> 来源手动输入 `20GP,40GP,40HC,45HC,20RF,40RF` -> 点击确定。此时每一个箱型单元格均具备标准下拉箭头。", indent=True)

    add_custom_heading("任务 04 解题步骤：复合船名航次分列", level=3)
    add_body_paragraph("【滑鼠操作】：在 H 列（船名航次）右侧插入 2 个空白列（命名为「船名」与「航次」）-> 选中 H 列数据 -> 点击【资料】->【资料剖析】-> 选择「分隔符号」-> 下一步 -> 勾选「其他」并输入 `/` -> 下一步 -> 目标储存档选取新插入的第一列 -> 点击完成。\n【公式微调】：使用 `=TRIM(新船名列)` 确保拆分后的船名尾部无残留空格。", indent=True)

    add_custom_heading("任务 05 解题步骤：港口代码补全中文名称", level=3)
    add_body_paragraph("【公式应用】：在装运港代码右侧插入「起运港名称」列，输入：\n`=XLOOKUP(TRIM(UPPER(I2)), 基础代码参照表!$B$2:$B$19, 基础代码参照表!$C$2:$C$19, \"未知港口\")`\n在卸货港代码右侧插入「目的港名称」列，输入：\n`=XLOOKUP(TRIM(UPPER(J2)), 基础代码参照表!$B$2:$B$19, 基础代码参照表!$C$2:$C$19, \"未知港口\")`\n【参考答案校验】：检查所有行均正确映射出「上海港」、「汉堡港」、「鹿特丹港」等标准中文港名，无 #N/A 报错。", indent=True)

    add_custom_heading("任务 06 & 07 解题步骤：日期标准化与航程延误核算", level=3)
    add_body_paragraph("【公式应用】：针对 ETD、ATD、ETA、ATA 四列，以嵌套替换函数转换为系统真日期：\n`=IF(ISNUMBER(K2), K2, IF(LEN(TRIM(K2))=8, DATE(LEFT(K2,4), MID(K2,5,2), RIGHT(K2,2)), DATEVALUE(SUBSTITUTE(SUBSTITUTE(TRIM(K2), \".\", \"-\"), \"/\", \"-\"))))`\n【计算航程天数】：新建「计划航程天数」列，输入公式：`=ETA列 - ETD列`，储存档格式设为常规数字。\n【计算到港延误天数】：新建「到港延误天数」列，输入公式：`=MAX(0, ATA列 - ETA列)`。若实际到港日早于预估抵港日，延误天数返回 0；若延误 3 天，则返回 3。", indent=True)

    add_custom_heading("任务 08 & 09 解题步骤：货物体积重量数值净化与运费核算", level=3)
    add_body_paragraph("【公式应用】：\n1. 纯净体积 (CBM)：`=VALUE(SUBSTITUTE(LOWER(O2), \"cbm\", \"\"))`\n2. 纯净毛重 (KGS)：`=VALUE(SUBSTITUTE(SUBSTITUTE(LOWER(P2), \"kgs\", \"\"), \",\", \"\"))`\n3. 计费吨 (RT)：`=MAX(CBM列, ROUND(KGS列 / 1000, 2))`\n4. 海运费与滞期费：`=VALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(UPPER(R2), \"$\", \"\"), \"USD\", \"\"), \",\", \"\"))`\n5. 应收总额核算：`=基础海运费 + 超期滞箱费`。\n【参考答案统计指标】：清洗后 297 笔有效业务的应收总费用求和公式 `=SUM(总费用列)`，标准参考总金额约为 **$438,250.00 ~ $452,000.00**（根据随机种子的精确求和结果一致），平均航程天数约为 23.5 天。", indent=True)

    add_custom_heading("任务 10 解题步骤：单证异常标记与视觉警示", level=3)
    add_body_paragraph("【条件格式高亮】：选中整表数据行 -> 点击【常用】->【条件格式】->【新增规则】-> 选择「使用公式来决定要格式化哪些储存档」-> 输入公式：\n`=OR($延误天数列>2, ISNUMBER(SEARCH(\"缺\", $单证状态列)))`\n点击格式按钮，设置背景色为浅红色填充。此时所有存在滞期滞留风险或单证缺失的工作单全部自动高亮，后勤调度人员可一目了然抓出高风险业务进行重点催办！", indent=True)

    # Page break to start Chapter 6 (Requirement 9: Every chapter starts on fresh page)
    doc.add_page_break()

    # -------------------------------------------------------------
    # CHAPTER 6: DATA QUALITY MANAGEMENT & AUTOMATION
    # -------------------------------------------------------------
    add_custom_heading("第 6 章 航运数据品质长效管控与自动化进阶建议", level=1)
    
    add_custom_heading("6.1 源头表单设计防错机制", level=2)
    add_body_paragraph("在日常航运管理中，与其在数据汇总结算时疲于奔命地进行清洗，不如在数据最初录入阶段就设立刚性的防错机制。以下是构建航运前置防错表单的三项黄金建议：", indent=True)
    add_bullet_point("对于港口（UN/LOCODE）、箱型（ISO 6346）、船东代码（SCAC Code）等高度标准化的枚举型栏位，绝对不允许操作员自由打字！必须全面布设「资料验证 -> 清单下拉选单」，从根源上消灭错别字与大小写混淆。", "1. 强制采用清单下拉选单约束：")
    add_bullet_point("对于提单号、集装箱号等关键栏位，利用资料验证中的「文字长度」选项（集装箱号强制锁定 11 位）以及「自订公式」验证格式合法性。例如集装箱号前 4 位必须为字母，可使用公式 `=AND(LEN(A2)=11, ISERR(VALUE(LEFT(A2,4))))` 进行阻断式校验。", "2. 设定字数与掩码格式刚性阻断：")
    add_bullet_point("将航运表格定义为「Excel 表格 (Ctrl+T)」。Excel 表格具备栏位名称结构化引用（Structured Reference）与公式自动向下继承扩展的优良特性，新增录入行时自动套用既有格式与验证规则，杜绝格式断层。", "3. 采用官方格式化表格 (Ctrl+T Table)：")

    add_custom_heading("6.2 进阶自动化工具展望：Power Query 与 Python in Excel", level=2)
    add_body_paragraph("随着现代航运数据量呈几何级数爆炸，当面对每日数万笔港口集疏运吞吐或集装箱动态数据时，传统的滑鼠手工前处理将逐渐显现天花板。航运行政人员在扎实掌握本手册基础技能后，可进一步拓展以下两条前沿技术栈：", indent=True)
    add_bullet_point("微软自 Excel 2016 起内建的革命性 ETL（抽取、转换、载入）引擎。用户只需通过直观的视窗按钮记录一次清洗步骤（如删除空格、拆分栏、去重、类型转换），系统底层会自动生成 M 语言脚本。日后每当船公司发出新的船期 CSV 档案，只需将档案扔进指定文件夹并点击【全部重新整理】，全套数据清洗瞬间在后台全自动完成！", "1. Power Query (获取与转换资料)：")
    add_bullet_point("微软最新引入的划时代特性。用户直接在 Excel 单元格中输入 `=PY(...)`，即可无缝调用 Python 生态最顶级的数据分析库（如 Pandas、NumPy）。诸如正则表达式（Regex）精准清洗超复杂箱号、批处理提单哈希去重等高级算法，数行 Python 代码即可优雅解决。", "2. Python in Excel：")

    add_custom_heading("6.3 航运行政日常数据处理十条铁律", level=2)
    add_bullet_point("原始数据永远保留一份干净底稿，严禁在唯一源数据上直接进行破坏性编辑。", "【铁律一：备份先行】")
    add_bullet_point("任何涉及求和、平均、比较运算的数值格，严禁夹带 'kg'、'cbm'、'usd' 等文本单位。", "【铁律二：数文分离】")
    add_bullet_point("所有日期必须满足系统真日期标准，格式统一收敛为 ISO 规范 YYYY-MM-DD。", "【铁律三：统一历法】")
    add_bullet_point("底层明细资料表中严禁出现跨行合并单元格，杜绝破坏二阶矩阵结构。", "【铁律四：拒绝合并】")
    add_bullet_point("数据记录行之间严禁插入空白行或分类小计行，统计汇总一律交由透视表处理。", "【铁律五：连续网格】")
    add_bullet_point("从网页或外部 EDI 报文复制而来的提单数据，务必先用 TRIM+CLEAN 脱敏净化。", "【铁律六：洗净首尾】")
    add_bullet_point("涉及关键主键去重时，必须精准圈定「提单号+柜号」多重主键，切勿全选误伤。", "【铁律七：精准去重】")
    add_bullet_point("跨表匹配港口或船东代码时，务必使用绝对引用锁定基准码表（$A$1:$B$20）。", "【铁律八：锁定基准】")
    add_bullet_point("完成复杂公式前处理后，务必【选择性贴上为值】，固化数据以防公式链路崩溃。", "【铁律九：固化数值】")
    add_bullet_point("在交付报表前，务必对关键栏位执行升降序抽检与极端值校验，对数字保有敬畏之心。", "【铁律十：闭环检验】")

    # Page break to start References (Requirement 9: Every chapter starts on fresh page)
    doc.add_page_break()

    # -------------------------------------------------------------
    # APPENDIX & REFERENCES
    # -------------------------------------------------------------
    add_custom_heading("附录与参考文献", level=1)
    
    add_custom_heading("附录 A：常用航运业务英文缩写与 UN/LOCODE 港口代码简表", level=2)
    
    tbl_apx = doc.add_table(rows=11, cols=3)
    tbl_apx.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_apx.autofit = False
    headers_apx = ["业务缩写 / 代码", "英文全称 (Full Name)", "业务实务涵义解析"]
    w_apx = [1.5, 2.3, 2.47]
    for c_idx, h_t in enumerate(headers_apx):
        cell = tbl_apx.cell(0, c_idx)
        cell.width = Inches(w_apx[c_idx])
        set_cell_background(cell, NAVY_HEX)
        set_cell_margins(cell, top=60, bottom=60, left=60, right=60)
        p = cell.paragraphs[0]
        style_paragraph(p, space_before=0, space_after=0)
        r = p.add_run(h_t)
        r.bold = True
        r.font.name = "Microsoft YaHei"
        r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    apx_data = [
        ("B/L (Bill of Lading)", "Ocean Bill of Lading", "海运提单，物权凭证与海上货物运输合同的核心证明"),
        ("ETD / ATD", "Estimated / Actual Time of Departure", "预计离港/开航时间 与 实际离港/开航时间"),
        ("ETA / ATA", "Estimated / Actual Time of Arrival", "预计到港时间 与 实际到港时间"),
        ("POL / POD", "Port of Loading / Port of Discharge", "装运港（起运港） 与 卸货港（目的港）"),
        ("TEU / FEU", "Twenty-foot / Forty-foot Equivalent Unit", "20呎标准集装箱折合单位 / 40呎标准集装箱折合单位"),
        ("CBM / KGS", "Cubic Meter / Kilograms", "立方米（体积单位） 与 公斤（毛重重量单位）"),
        ("RT (Revenue Ton)", "Revenue Ton (计费吨)", "运费计费吨，海运拼箱中取体积 CBM 与重量公吨之较大者"),
        ("Demurrage / Detention", "Container Demurrage & Detention", "超期滞箱费（堆场内超期） 与 滞港滞箱费（场外逾期未还）"),
        ("UN/LOCODE", "United Nations Code for Trade Locations", "联合国五位港口地点代码（如 CNSHA 上海、DEHAM 汉堡）"),
        ("ISO 6346", "Freight Containers Coding Standard", "国际标准货柜代码规范，定义箱主4字母代码与7位校验序列")
    ]
    for r_idx, (k, eng, desc) in enumerate(apx_data, start=1):
        bg = "FFFFFF" if r_idx % 2 == 1 else "F9FAFC"
        for c_idx, val in enumerate([k, eng, desc]):
            cell = tbl_apx.cell(r_idx, c_idx)
            cell.width = Inches(w_apx[c_idx])
            set_cell_background(cell, bg)
            set_cell_margins(cell, top=45, bottom=45, left=60, right=60)
            p = cell.paragraphs[0]
            style_paragraph(p, space_before=0, space_after=0)
            r = p.add_run(val)
            r.font.name = "Microsoft YaHei"
            r._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
            r.font.size = Pt(8)
            r.font.color.rgb = ACCENT_BLUE if c_idx == 0 else TEXT_DARK
            if c_idx == 0:
                r.bold = True

    add_custom_heading("参考文献与引证来源", level=2)
    add_body_paragraph("本教学手册在撰写过程中，严格依据国际航运业权威行业标准与数据科学理论编订，特此注明以下权威参考来源：", indent=True)
    
    references = [
        "[1] 国际海事组织 (IMO). 《国际便利海上运输公约》(FAL Convention) 电子单证与舱单数据交换标准规范. 伦敦: IMO 出版社, 2024.",
        "[2] 国际标准化组织 (ISO). ISO 6346: Freight containers — Coding, identification and marking (集装箱编码、识别与标记标准). 日内瓦: ISO 组织, 2022.",
        "[3] 联合国贸易便利化与电子业务中心 (UN/CEFACT). UN/LOCODE: United Nations Code for Trade and Transport Locations (联合国贸易和运输地点代码汇编). 日内瓦: 联合国欧洲经济委员会 (UNECE), 2026.",
        "[4] 威克姆, 哈德利 (Wickham, Hadley). Tidy Data (整洁数据理论). Journal of Statistical Software (统计软件学报), 2014, 59(10): 1-23.",
        "[5] 微软官方知识库 (Microsoft Support). Excel 功能与公式技术文档：数据转换、文本函数与动态数组规范 (Office 365 / 2021 版). 雷德蒙德: 微软公司, 2026.",
        "[6] 柯德, 埃德加·弗兰克 (Codd, E. F.). A Relational Model of Data for Large Shared Data Banks (大型共享数据银行的关系模型). Communications of the ACM, 1970, 13(6): 377-387.",
        "[7] 中国国际货运代理协会 (CIFA). 国际货运代理基础业务规范与单证实务教程. 北京: 中国对外经济贸易出版社, 2023."
    ]
    for ref_text in references:
        p_ref = doc.add_paragraph()
        style_paragraph(p_ref, space_before=2, space_after=4, line_spacing=1.2)
        p_ref.paragraph_format.left_indent = Inches(0.28)
        p_ref.paragraph_format.first_line_indent = Inches(-0.28)
        r_ref = p_ref.add_run(ref_text)
        r_ref.font.name = "Microsoft YaHei"
        r_ref._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')
        r_ref.font.size = Pt(8.5)
        r_ref.font.color.rgb = TEXT_DARK

    # -------------------------------------------------------------
    # Save the Word Document
    # -------------------------------------------------------------
    out_docx_path = "航运行政人员_Excel基础资料分析与资料前处理教学手册.docx"
    doc.save(out_docx_path)
    print(f"Successfully generated formatted Word document: {out_docx_path}")

if __name__ == "__main__":
    create_shipping_tutorial_docx()
