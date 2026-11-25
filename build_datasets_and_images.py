import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as patches
import pandas as pd
import numpy as np
import random
from datetime import datetime

# Font config
plt.rcParams['font.sans-serif'] = ['SimHei', 'Microsoft YaHei', 'DejaVu Sans', 'Arial']
plt.rcParams['axes.unicode_minus'] = False

os.makedirs("images", exist_ok=True)
os.makedirs("data", exist_ok=True)

print("Starting generation of datasets and images...")

# -------------------------------------------------------------
# 1. Generate 30-row demonstration Excel file
# -------------------------------------------------------------
wb30 = openpyxl.Workbook()

# Sheet 1: Raw Dirty Data
ws_raw = wb30.active
ws_raw.title = "30笔原始脏资料"

headers_30 = [
    "序号", "提单号 (B/L No.)", "集装箱号 (Container No.)", "箱型 (Type)", 
    "船名航次 (Vessel/Voyage)", "装运港 (POL)", "目的港 (POD)", 
    "预估开航日 (ETD)", "预估抵港日 (ETA)", "货重 (Gross Wt)", "运费金额 (Ocean Freight)", "备注 (Remarks)"
]

raw_data_30 = [
    (1, " COSU62918401 ", "cosu9823145", "40GP ", "cosco guangzhou / 042w", "CNSHA", "DEHAM", "2026-11-05", "2026/12/10", "18,450 kg", "$2,450.00", " 已放行 "),
    (2, "COSU62918402", "TEMU4819201", "20gp", "EVER GIVEN/0118E", " cnsha ", "nlsrt", "2026/11/06", "2026.12.12", "12200", "USD 1200", "电放"),
    (3, "MSCU91823711", "MSCU1928374", "40HC", "MSC OSCAR / 2603W", "CNNBO", "SGSIN", "2026.11.07", "2026-11-15", "24,100 KG", "1800", "特快提单"),
    (4, "COSU62918402", "TEMU4819201", "20gp", "EVER GIVEN/0118E", " cnsha ", "nlsrt", "2026/11/06", "2026.12.12", "12200", "USD 1200", "电放"), # Duplicate
    (5, "OOCL82910394", "OOCU9283019", "40hc", "OOCL POLAND / 088E", "cnsha", "USLAX", "20261108", "2026-11-24", "19800", "$3,100", "退税报关"),
    (6, "CMAU28190283", "CMAU8172930", "20GP", "cma cgm marco polo/091w", "CNNBO", "FRLEH", "2026-11-09", "2026-12-18", "14500 kg", "USD 1,650.00", "加急查验"),
    (7, "ONEY82910283", "TGHU9182039", "40GP", "ONE APUS/0064E", "CNSHA", "USOAK", "2026/11/10", "2026/11/28", "21000", "2950", "危品需温控"),
    (8, "EMCU92810394", "EGLU1928304", "40HC ", "EVER GLORY / 0192W", "CNNBO", "GRPIR", "2026.11.11", "2026.12.05", "22400 KGS", "$2,100.00", "正本"),
    (9, "HLCU82910384", "HLXU9182736", "20GP", "YANTIAN EXPRESS / 2641E", "CNDLC", "USNYC", "2026-11-12", "2026-12-20", "15300", "USD 2200", "中转港预报"),
    (10, "COSU62918403", "COSU1029384", "40hc", "COSCO PRIDE/031W", "CNSHA", "BEANR", "20261113", "2026-12-16", "23500kg", "2,350", "装箱未结"),
    (11, "MSCU91823712", "MEDU8192034", "40HC", "msc diana / 2611E", "CNSHA", "ITGOA", "2026-11-14", "2026/12/19", "25100", "$2,600", "正本提单已寄"),
    (12, "SITC28190384", "SITU9182034", "20GP", "SITC NINGBO / 2645S", "CNNBO", "JPTYO", "2026/11/15", "2026-11-19", "9800 KG", "USD 450", "直航快速"),
    (13, "KMTC82910384", "KMTU8192039", "20gp", "KMTC PUSAN / 2618N", "CNQDG", "KRPUS", "2026.11.16", "2026.11.18", "11200", "380", "近洋加班船"),
    (14, "COSU62918404", "TCLU9182038", "40GP ", "cosco shipping universe / 012w", "CNSHA", "GBFXT", "2026-11-17", "2026-12-22", "18900 kgs", "$2,500.00", "大客户约价"),
    (15, "YMLU82910384", "YMLU9182034", "40HC", "YM WELLHEAD / 038E", "CNNBO", "USLGB", "20261118", "2026/12/03", "20500", "USD 3250", "需做ACI"),
    (16, "WHLU92810384", "WHLU8192038", "20GP", "WAN HAI 512 / S102", "CNXMN", "SGSIN", "2026-11-19", "2026-11-26", "13100", "750", "待改单"),
    (17, "COSU62918405", "CSNU9182034", "40hc", "COSCO DEVELOPMENT/024W", "cnsha", "deham", "2026/11/20", "2026.12.25", "22100 KG", "$2,400", "双清包税"),
    (18, "MSCU91823713", "MSCU8291034", "40HC", "MSC ISABELLA / 2615W", "CNNBO", "ESVLC", "2026.11.21", "2026-12-27", "24800", "USD 2450.00", "已索赔"),
    (19, "OOCL82910395", "OOLU9182034", "20gp", "OOCL BANGKOK / 041E", "CNSHA", "THBKK", "2026-11-22", "2026/11/29", "10500 kg", "680", "需熏蒸证"),
    (20, "CMAU28190284", "CMAU9182034", "40GP", "cma cgm jacques saade/035w", "CNSHA", "FRLEH", "20261123", "2026-12-30", "19200", "$2,700", "已申领原产地证"),
    (21, "SITC28190384", "SITU9182034", "20GP", "SITC NINGBO / 2645S", "CNNBO", "JPTYO", "2026/11/15", "2026-11-19", "9800 KG", "USD 450", "直航快速"), # Duplicate
    (22, "ZIMU82910384", "ZIMU9182034", "40HC", "ZIM ROTTERDAM / 081E", "CNQDG", "USSAV", "2026-11-24", "2026/12/28", "21500 KGS", "USD 3,600.00", "电商专线"),
    (23, "HLCU82910385", "HLXU8192034", "20GP ", "BERLIN EXPRESS / 019W", "CNNBO", "DEHAM", "2026/11/25", "2026.12.31", "14800", "1600", "保函已收"),
    (24, "ONEY82910284", "NYKU9182034", "40hc", "ONE TRUST / 0012E", "cnsha", "ustiw", "2026.11.26", "2026-12-14", "22900 kg", "$3,300.00", "换单迟缓"),
    (25, "PILU82910384", "PCIU9182034", "20gp", "KOTA PELEPAS / 0281S", "CNNBO", "MYPKG", "2026-11-27", "2026/12/04", "12600", "520", "拼箱直达"),
    (26, "COSU62918406", "TGHU8192034", "40HC", "cosco shipping nebula / 019w", "CNSHA", "NLAMS", "20261128", "2027-01-02", "23800", "USD 2550", "港区已进重"),
    (27, "EMCU92810395", "EGLU8192034", "40GP", "EVER GOLDEN/0210E", "CNQDG", "USLAX", "2026/11/29", "2026-12-16", "18500 KG", "$3,050.00", "海关查验完毕"),
    (28, "MSCU91823714", "TCLU8192034", "20GP", "MSC GULSUN / 2620W", "CNNBO", "EGSUZ", "2026-11-30", "2026.12.18", "13700", "1450", "苏伊士过闸"),
    (29, "KMTC82910385", "KMTU9182034", "40hc", "KMTC INCHEON / 2622N", "CNSHA", "KRINC", "2026.12.01", "2026/12/03", "20100 kgs", "USD 720.00", "冷链监控"),
    (30, "COSU62918407", "COSU8192034", "20GP ", "COSCO GLORY / 015W", " cnsha ", "deham", "2026-12-02", "2027-01-08", "15900", "1500", "危险品三类")
]

# Write headers & styling
header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
thin_border = Border(
    left=Side(style='thin', color='D3D3D3'),
    right=Side(style='thin', color='D3D3D3'),
    top=Side(style='thin', color='D3D3D3'),
    bottom=Side(style='thin', color='D3D3D3')
)

ws_raw.append(headers_30)
for col_num in range(1, len(headers_30) + 1):
    cell = ws_raw.cell(row=1, column=col_num)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

for row_idx, r in enumerate(raw_data_30, start=2):
    ws_raw.append(r)
    for col_idx in range(1, len(r) + 1):
        c = ws_raw.cell(row=row_idx, column=col_idx)
        c.border = thin_border
        c.font = Font(name="Calibri", size=10)
        if col_idx in [1, 4, 6, 7, 8, 9]:
            c.alignment = Alignment(horizontal="center", vertical="center")
        elif col_idx in [10, 11]:
            c.alignment = Alignment(horizontal="right", vertical="center")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")

# Sheet 2: Standard Clean Data
ws_clean = wb30.create_sheet(title="30笔清洗后标准资料")
clean_headers = [
    "标准序号", "提单号", "集装箱号", "箱型代码", "船名", "航次", 
    "起运港代码", "起运港名称", "目的港代码", "目的港名称", 
    "预计开航日(ETD)", "预计抵港日(ETA)", "航程天数", 
    "货重(KG)", "海运费(USD)", "处理状态", "备注"
]

port_names = {
    "CNSHA": "上海 (Shanghai)", "CNNBO": "宁波 (Ningbo)", "CNQDG": "青岛 (Qingdao)", 
    "CNDLC": "大连 (Dalian)", "CNXMN": "厦门 (Xiamen)", "DEHAM": "汉堡 (Hamburg)", 
    "NLSRT": "鹿特丹 (Rotterdam)", "SGSIN": "新加坡 (Singapore)", "USLAX": "洛杉矶 (Los Angeles)", 
    "FRLEH": "勒阿弗尔 (Le Havre)", "USOAK": "奥克兰 (Oakland)", "GRPIR": "比雷埃夫斯 (Piraeus)", 
    "USNYC": "纽约 (New York)", "BEANR": "安特卫普 (Antwerp)", "ITGOA": "热那亚 (Genoa)", 
    "JPTYO": "东京 (Tokyo)", "KRPUS": "釜山 (Busan)", "GBFXT": "费利克斯托 (Felixstowe)", 
    "USLGB": "长滩 (Long Beach)", "ESVLC": "巴伦西亚 (Valencia)", "THBKK": "曼谷 (Bangkok)", 
    "USSAV": "萨凡纳 (Savannah)", "USTIW": "塔科马 (Tacoma)", "MYPKG": "巴生港 (Port Klang)", 
    "NLAMS": "阿姆斯特丹 (Amsterdam)", "EGSUZ": "苏伊士 (Suez Canal)", "KRINC": "仁川 (Incheon)"
}

ws_clean.append(clean_headers)
clean_header_fill = PatternFill(start_color="274E13", end_color="274E13", fill_type="solid")
for col_num in range(1, len(clean_headers) + 1):
    cell = ws_clean.cell(row=1, column=col_num)
    cell.fill = clean_header_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

seen_containers = set()
clean_rows = []
for r in raw_data_30:
    bl = str(r[1]).strip().upper()
    cntr = str(r[2]).strip().upper()
    if cntr in seen_containers:
        continue # skip duplicate
    seen_containers.add(cntr)
    
    ctype = str(r[3]).strip().upper()
    vessel_raw = str(r[4]).strip()
    if "/" in vessel_raw:
        parts = vessel_raw.split("/")
        vname = parts[0].strip().title()
        voy = parts[1].strip().upper()
    else:
        vname = vessel_raw.title()
        voy = "-"
        
    pol = str(r[5]).strip().upper()
    pod = str(r[6]).strip().upper()
    pol_name = port_names.get(pol, pol)
    pod_name = port_names.get(pod, pod)
    
    def parse_d(dstr):
        dstr = str(dstr).strip().replace(".", "-").replace("/", "-")
        if len(dstr) == 8 and dstr.isdigit():
            dstr = f"{dstr[:4]}-{dstr[4:6]}-{dstr[6:]}"
        return dstr
        
    etd = parse_d(r[7])
    eta = parse_d(r[8])
    
    try:
        d1 = datetime.strptime(etd, "%Y-%m-%d")
        d2 = datetime.strptime(eta, "%Y-%m-%d")
        transit = (d2 - d1).days
    except:
        transit = "-"
        
    wstr = str(r[9]).lower().replace(",", "").replace("kgs", "").replace("kg", "").strip()
    try:
        wt = float(wstr)
    except:
        wt = 0.0
        
    fstr = str(r[10]).upper().replace("$", "").replace("USD", "").replace(",", "").strip()
    try:
        frt = float(fstr)
    except:
        frt = 0.0
        
    remarks = str(r[11]).strip()
    status = "有效放行"
    
    clean_rows.append([
        len(clean_rows) + 1, bl, cntr, ctype, vname, voy,
        pol, pol_name, pod, pod_name,
        etd, eta, transit, wt, frt, status, remarks
    ])

for row_idx, r in enumerate(clean_rows, start=2):
    ws_clean.append(r)
    for col_idx in range(1, len(r) + 1):
        c = ws_clean.cell(row=row_idx, column=col_idx)
        c.border = thin_border
        c.font = Font(name="Calibri", size=10)
        if col_idx in [1, 2, 3, 4, 6, 7, 9, 11, 12, 13, 16]:
            c.alignment = Alignment(horizontal="center", vertical="center")
        elif col_idx in [14, 15]:
            c.alignment = Alignment(horizontal="right", vertical="center")
            if col_idx == 14:
                c.number_format = '#,##0'
            elif col_idx == 15:
                c.number_format = '$#,##0.00'
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")

# Auto-adjust column widths
for ws in [ws_raw, ws_clean]:
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)

wb30.save("data/航运资料前处理_30笔示范档.xlsx")
print("Saved data/航运资料前处理_30笔示范档.xlsx successfully!")


# -------------------------------------------------------------
# 2. Generate 300-row Practice Excel file
# -------------------------------------------------------------
wb300 = openpyxl.Workbook()
ws300_raw = wb300.active
ws300_raw.title = "300笔情境原始资料"

headers_300 = [
    "行号", "业务类型 (Scenario)", "工作单号 (Job No.)", "主提单号 (MBL)", "集装箱号 (Container)",
    "箱型 (ISO Type)", "承运船东 (Carrier)", "船名航次 (Vessel/Voyage)", "装运港代码 (POL)",
    "卸货港代码 (POD)", "预估开航日 (ETD)", "实际离港日 (ATD)", "预估抵港日 (ETA)",
    "实际到港日 (ATA)", "货物体积 (CBM)", "货物毛重 (KGS)", "计费吨 (Revenue Ton)",
    "基础海运费 (USD)", "超期滞箱费 (USD)", "总应收费用 (USD)", "单证完整状态", "异常标记"
]

ws300_raw.append(headers_300)
h300_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
for col_num in range(1, len(headers_300) + 1):
    cell = ws300_raw.cell(row=1, column=col_num)
    cell.fill = h300_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

random.seed(20261125)

carriers = [("COSCO", "中远海运"), ("MAERSK", "马士基"), ("CMA-CGM", "达飞海运"), ("EVERGREEN", "长荣海运"), ("ONE", "海洋网联"), ("MSC", "地中海航运")]
pols = ["CNSHA", "CNNBO", "CNQDG", "CNSZX", "CNXMN", "CNDLC"]
pods = ["DEHAM", "NLSRT", "USLAX", "USNYC", "SGSIN", "JPTYO", "KRPUS", "FRLEH", "GBFXT", "ITGOA", "MYPKG", "AEJEA", "AUMEL"]
scenarios = [
    "A-集装箱流转与跟踪", 
    "B-运杂费与滞期结算", 
    "C-船期准点率与港口作业"
]
types = ["20GP", "40GP", "40HC", "45HC", "20RF", "40RF"]

rows_300 = []
for i in range(1, 301):
    sc = scenarios[i % 3]
    job_no = f"SHP2026{i:04d}"
    c_code, c_name = random.choice(carriers)
    
    bl = f"{c_code[:4]}{random.randint(1000000, 9999999)}"
    if random.random() < 0.2:
        bl = f" {bl}  "
    elif random.random() < 0.15:
        bl = bl.lower()
        
    cntr_prefix = c_code[:3] + "U"
    cntr_num = random.randint(1000000, 9999999)
    cntr = f"{cntr_prefix}{cntr_num}"
    if random.random() < 0.15:
        cntr = cntr.lower()
    if random.random() < 0.1:
        cntr = f"  {cntr} "
        
    iso_t = random.choice(types)
    if random.random() < 0.2:
        iso_t = iso_t.lower() + " "
        
    vessel = f"{c_name}探索号 / {random.randint(100, 999)}W"
    if random.random() < 0.2:
        vessel = vessel.replace(" / ", "/").lower()
        
    pol = random.choice(pols)
    pod = random.choice(pods)
    if random.random() < 0.15:
        pol = f" {pol.lower()} "
    if random.random() < 0.15:
        pod = pod.lower()
        
    day_etd = (i % 28) + 1
    etd_obj = datetime(2026, 11, day_etd)
    transit = random.randint(12, 35)
    eta_obj = etd_obj + pd.Timedelta(days=transit)
    
    delay_dep = random.choice([0, 0, 0, 1, 2, -1])
    atd_obj = etd_obj + pd.Timedelta(days=delay_dep)
    
    delay_arr = random.choice([0, 0, 1, 3, 5, 0, -1])
    ata_obj = eta_obj + pd.Timedelta(days=delay_arr)
    
    def fmt_date_dirty(dt):
        fmt = random.choice(["%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"])
        return dt.strftime(fmt)
        
    etd_s = fmt_date_dirty(etd_obj)
    atd_s = fmt_date_dirty(atd_obj)
    eta_s = fmt_date_dirty(eta_obj)
    ata_s = fmt_date_dirty(ata_obj)
    
    cbm = round(random.uniform(8.0, 68.0), 2)
    kgs = round(random.uniform(5000, 26000), 1)
    rt = max(cbm, round(kgs / 1000.0, 2))
    base_frt = round(rt * random.uniform(35.0, 85.0), 2)
    demurrage_days = max(0, delay_arr - 2)
    demurrage_fee = demurrage_days * 120.0
    total_fee = base_frt + demurrage_fee
    
    cbm_s = f"{cbm} CBM" if random.random() < 0.3 else str(cbm)
    kgs_s = f"{kgs:,} KGS" if random.random() < 0.3 else str(kgs)
    rt_s = f"{rt} RT" if random.random() < 0.3 else str(rt)
    frt_s = f"${base_frt:,.2f}" if random.random() < 0.4 else f"USD {base_frt}" if random.random() < 0.3 else str(base_frt)
    dem_s = f"${demurrage_fee:.2f}" if random.random() < 0.3 else str(demurrage_fee)
    tot_s = f"${total_fee:,.2f}" if random.random() < 0.4 else str(total_fee)
    
    doc_status = random.choice(["已归档", "缺装箱单", "电放已核", "待产地证", "正本待寄", "海关放行"])
    anomaly = "正常"
    if demurrage_fee > 0:
        anomaly = f"滞期超期{demurrage_days}天"
    elif "缺" in doc_status or "待" in doc_status:
        anomaly = "单证待补"
        
    rows_300.append([
        i, sc, job_no, bl, cntr, iso_t, c_code, vessel, pol, pod,
        etd_s, atd_s, eta_s, ata_s, cbm_s, kgs_s, rt_s,
        frt_s, dem_s, tot_s, doc_status, anomaly
    ])

# deliberate duplicates
rows_300[40] = list(rows_300[39])
rows_300[40][0] = 41
rows_300[120] = list(rows_300[119])
rows_300[120][0] = 121
rows_300[210] = list(rows_300[209])
rows_300[210][0] = 211

for row_idx, r in enumerate(rows_300, start=2):
    ws300_raw.append(r)
    for col_idx in range(1, len(r) + 1):
        c = ws300_raw.cell(row=row_idx, column=col_idx)
        c.border = thin_border
        c.font = Font(name="Calibri", size=9)
        if col_idx in [1, 2, 6, 7, 9, 10, 11, 12, 13, 14, 21, 22]:
            c.alignment = Alignment(horizontal="center", vertical="center")
        elif col_idx in [15, 16, 17, 18, 19, 20]:
            c.alignment = Alignment(horizontal="right", vertical="center")
        else:
            c.alignment = Alignment(horizontal="left", vertical="center")

# Sheet 2: Reference port and carrier lookups
ws_ref = wb300.create_sheet(title="基础代码参照表")
ws_ref.append(["代码类型", "业务代码", "中文全称", "英文全称/国家", "主要航线/属性"])
ref_headers_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
for c_idx in range(1, 6):
    cell = ws_ref.cell(row=1, column=c_idx)
    cell.fill = ref_headers_fill
    cell.font = header_font
    cell.alignment = Alignment(horizontal="center", vertical="center")

ref_rows = [
    ["港口 (POL)", "CNSHA", "上海港", "Shanghai, China", "华东干线港"],
    ["港口 (POL)", "CNNBO", "宁波舟山港", "Ningbo, China", "华东干线港"],
    ["港口 (POL)", "CNQDG", "青岛港", "Qingdao, China", "北方主枢纽"],
    ["港口 (POL)", "CNSZX", "深圳盐田/蛇口", "Shenzhen, China", "华南主枢纽"],
    ["港口 (POL)", "CNXMN", "厦门港", "Xiamen, China", "东南沿海"],
    ["港口 (POL)", "CNDLC", "大连港", "Dalian, China", "东北主枢纽"],
    ["港口 (POD)", "DEHAM", "汉堡港", "Hamburg, Germany", "欧洲基本港"],
    ["港口 (POD)", "NLSRT", "鹿特丹港", "Rotterdam, Netherlands", "欧洲基本港"],
    ["港口 (POD)", "USLAX", "洛杉矶港", "Los Angeles, USA", "美西基本港"],
    ["港口 (POD)", "USNYC", "纽约港", "New York, USA", "美东基本港"],
    ["港口 (POD)", "SGSIN", "新加坡港", "Singapore", "东南亚中转枢纽"],
    ["港口 (POD)", "JPTYO", "东京港", "Tokyo, Japan", "日韩航线"],
    ["港口 (POD)", "KRPUS", "釜山港", "Busan, South Korea", "东北亚中转港"],
    ["港口 (POD)", "FRLEH", "勒阿弗尔港", "Le Havre, France", "欧洲大西洋"],
    ["港口 (POD)", "GBFXT", "费利克斯托港", "Felixstowe, UK", "英国主港"],
    ["港口 (POD)", "ITGOA", "热那亚港", "Genoa, Italy", "地中海基本港"],
    ["港口 (POD)", "MYPKG", "巴生港", "Port Klang, Malaysia", "东南亚枢纽"],
    ["港口 (POD)", "AEJEA", "杰贝阿里港", "Jebel Ali, UAE", "中东主枢纽"],
    ["港口 (POD)", "AUMEL", "墨尔本港", "Melbourne, Australia", "澳洲主港"],
    ["船东 (Carrier)", "COSCO", "中远海运", "COSCO Shipping Lines", "Ocean Alliance"],
    ["船东 (Carrier)", "MAERSK", "马士基航运", "Maersk Line", "Gemini Cooperation"],
    ["船东 (Carrier)", "CMA-CGM", "达飞轮船", "CMA CGM Group", "Ocean Alliance"],
    ["船东 (Carrier)", "EVERGREEN", "长荣海运", "Evergreen Marine Corp", "Ocean Alliance"],
    ["船东 (Carrier)", "ONE", "海洋网联船务", "Ocean Network Express", "THE Alliance/Premier"],
    ["船东 (Carrier)", "MSC", "地中海航运", "Mediterranean Shipping Co", "Independent"]
]
for r in ref_rows:
    ws_ref.append(r)
    for c_idx in range(1, 6):
        cell = ws_ref.cell(row=ws_ref.max_row, column=c_idx)
        cell.border = thin_border
        cell.font = Font(name="Calibri", size=10)
        cell.alignment = Alignment(horizontal="center" if c_idx in [1, 2] else "left", vertical="center")

# Auto-fit columns
for ws in [ws300_raw, ws_ref]:
    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 2, 11)

wb300.save("data/航运资料前处理_300笔情境练习档.xlsx")
print("Saved data/航运资料前处理_300笔情境练习档.xlsx successfully!")


# -------------------------------------------------------------
# 3. Generate High-Res Instructional Figures
# -------------------------------------------------------------
print("Generating UI and workflow illustrations...")

# Figure 1-1: MS Excel UI Overview Diagram
fig, ax = plt.subplots(figsize=(12, 7.5), dpi=200)
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

rect_win = patches.Rectangle((1, 1), 98, 98, linewidth=2, edgecolor='#1E4E79', facecolor='#F3F3F3')
ax.add_patch(rect_win)

rect_tb = patches.Rectangle((1, 93), 98, 6, linewidth=1, edgecolor='#107C41', facecolor='#107C41')
ax.add_patch(rect_tb)
ax.text(3, 96, "快速存取工具列  [ 💾  ↩  ↪ ]", color='white', fontsize=11, fontweight='bold', va='center')
ax.text(50, 96, "航运资料分析_工作表1.xlsx - Excel", color='white', fontsize=12, fontweight='bold', ha='center', va='center')
ax.text(95, 96, "—  ❐  ✕", color='white', fontsize=12, ha='center', va='center')

rect_tabs = patches.Rectangle((1, 88), 98, 5, facecolor='#E1DFDD', edgecolor='#D2D0CE')
ax.add_patch(rect_tabs)
tabs = ["档案 (File)", "常用 (Home)", "插入 (Insert)", "版面配置 (Page Layout)", "公式 (Formulas)", "资料 (Data) ★", "校阅 (Review)", "检视 (View)", "说明 (Help)"]
x_t = 3
for t in tabs:
    is_active = "资料" in t
    bg_c = '#FFFFFF' if is_active else '#E1DFDD'
    fg_c = '#107C41' if is_active else '#323130'
    w_t = len(t) * 1.5 + 2
    r_tab = patches.Rectangle((x_t, 88), w_t, 5, facecolor=bg_c, edgecolor='#C8C6C4')
    ax.add_patch(r_tab)
    ax.text(x_t + w_t/2, 90.5, t, color=fg_c, fontsize=9.5, fontweight='bold' if is_active else 'normal', ha='center', va='center')
    x_t += w_t + 1

rect_ribbon = patches.Rectangle((1, 74), 98, 14, facecolor='#FFFFFF', edgecolor='#C8C6C4')
ax.add_patch(rect_ribbon)

groups = [
    ("取得及转换资料\n(Get & Transform)", ["从文字/CSV", "从Web", "从资料表"]),
    ("排序与筛选\n(Sort & Filter)", ["A-Z 递增", "Z-A 递减", "筛选 (Filter)"]),
    ("资料工具 (核心前处理) ★\n(Data Tools)", ["资料剖析 (Text to Columns)", "移除重复项 (Remove Duplicates)", "快速填入 (Ctrl+E)", "资料验证 (Data Validation)"]),
    ("预测与大纲\n(Forecast & Outline)", ["假设分析", "建立群组"])
]
x_g = 2
for g_title, g_items in groups:
    is_core = "核心" in g_title
    w_g = 23 if not is_core else 34
    r_grp = patches.Rectangle((x_g, 75), w_g, 12.5, facecolor='#F3F9F4' if is_core else '#FAFAFA', edgecolor='#107C41' if is_core else '#EDEBE9', linewidth=1.5 if is_core else 1)
    ax.add_patch(r_grp)
    y_it = 84
    for it in g_items:
        ax.text(x_g + 1.5, y_it, f"▪ {it}", fontsize=8.5, color='#1E4E79' if is_core else '#323130', fontweight='bold' if is_core else 'normal')
        y_it -= 2.6
    ax.text(x_g + w_g/2, 75.8, g_title, fontsize=8, color='#107C41' if is_core else '#605E5C', ha='center', va='bottom', fontweight='bold')
    x_g += w_g + 1.2

rect_fb = patches.Rectangle((1, 68), 98, 5, facecolor='#F3F2F1', edgecolor='#C8C6C4')
ax.add_patch(rect_fb)
r_nb = patches.Rectangle((3, 69), 12, 3.2, facecolor='white', edgecolor='#8A8886')
ax.add_patch(r_nb)
ax.text(9, 70.6, "C4", fontsize=10, fontweight='bold', ha='center', va='center')
ax.text(18, 70.6, "fx", fontsize=11, fontweight='bold', color='#605E5C', ha='center', va='center')
r_fi = patches.Rectangle((21, 69), 76, 3.2, facecolor='white', edgecolor='#8A8886')
ax.add_patch(r_fi)
ax.text(23, 70.6, '=TRIM(CLEAN(UPPER(B4)))  // 正在清洗并规范提单编号文本', fontsize=9.5, color='#004E8C', fontfamily='monospace', va='center')

rect_grid = patches.Rectangle((1, 7), 98, 60, facecolor='white', edgecolor='#C8C6C4')
ax.add_patch(rect_grid)

col_names = ["", "A (序号)", "B (提单号B/L)", "C (清洗后提单)", "D (箱号Container)", "E (箱型)", "F (起运港POL)", "G (目的港POD)", "H (海运费USD)"]
col_widths = [5, 9, 15, 16, 15, 8, 12, 12, 12]
x_c = 1
for idx, (c_name, c_w) in enumerate(zip(col_names, col_widths)):
    r_ch = patches.Rectangle((x_c, 63), c_w, 4, facecolor='#E1DFDD', edgecolor='#C8C6C4')
    ax.add_patch(r_ch)
    ax.text(x_c + c_w/2, 65, c_name, fontsize=8.5, fontweight='bold', ha='center', va='center')
    x_c += c_w

grid_data = [
    ("1", "1", " COSU62918401 ", "COSU62918401", "cosu9823145", "40GP", "CNSHA", "DEHAM", "$2,450.00"),
    ("2", "2", "COSU62918402", "COSU62918402", "TEMU4819201", "20GP", "CNSHA", "NLSRT", "$1,200.00"),
    ("3", "3", "MSCU91823711", "MSCU91823711", "MSCU1928374", "40HC", "CNNBO", "SGSIN", "$1,800.00"),
    ("4", "4", "COSU62918402 (重)", "COSU62918402", "TEMU4819201", "20GP", "CNSHA", "NLSRT", "$1,200.00"),
    ("5", "5", " OOCL82910394", "OOCL82910394", "OOCU9283019", "40HC", "CNSHA", "USLAX", "$3,100.00"),
    ("6", "6", "CMAU28190283 ", "CMAU28190283", "CMAU8172930", "20GP", "CNNBO", "FRLEH", "$1,650.00"),
    ("7", "7", "ONEY82910283", "ONEY82910283", "TGHU9182039", "40GP", "CNSHA", "USOAK", "$2,950.00"),
    ("8", "8", " EMCU92810394 ", "EMCU92810394", "EGLU1928304", "40HC", "CNNBO", "GRPIR", "$2,100.00"),
    ("9", "9", "HLCU82910384", "HLCU82910384", "HLXU9182736", "20GP", "CNDLC", "USNYC", "$2,200.00"),
    ("10", "10", "COSU62918403", "COSU62918403", "COSU1029384", "40HC", "CNSHA", "BEANR", "$2,350.00"),
]

y_r = 58
for row_tuple in grid_data:
    x_c = 1
    is_dup = "重" in row_tuple[2]
    for idx, (val, c_w) in enumerate(zip(row_tuple, col_widths)):
        bg = '#FFE2E2' if (is_dup and idx > 0) else ('#E1DFDD' if idx == 0 else ('#EDF5E1' if idx == 3 else 'white'))
        r_cell = patches.Rectangle((x_c, y_r), c_w, 4.5, facecolor=bg, edgecolor='#E1DFDD')
        ax.add_patch(r_cell)
        align = 'center' if idx in [0, 1, 5, 6, 7] else ('right' if idx == 8 else 'left')
        ax.text(x_c + (c_w/2 if align == 'center' else (c_w - 1.5 if align == 'right' else 1.5)), 
                y_r + 2.2, val, fontsize=8, va='center', ha=align, 
                color='#C00000' if is_dup else '#323130', fontweight='bold' if (idx == 3 or is_dup) else 'normal')
        x_c += c_w
    y_r -= 4.7

rect_sb = patches.Rectangle((1, 1), 98, 6, facecolor='#F3F2F1', edgecolor='#C8C6C4')
ax.add_patch(rect_sb)
sheet_names = ["30笔原始脏资料", "30笔清洗后标准资料", "参照代码对照表"]
x_st = 4
for s_idx, sn in enumerate(sheet_names):
    is_cur = s_idx == 0
    w_st = len(sn) * 1.5 + 2
    r_st = patches.Rectangle((x_st, 2), w_st, 4, facecolor='white' if is_cur else '#E1DFDD', edgecolor='#C8C6C4')
    ax.add_patch(r_st)
    ax.text(x_st + w_st/2, 4, sn, fontsize=8, fontweight='bold' if is_cur else 'normal', color='#107C41' if is_cur else '#323130', ha='center', va='center')
    x_st += w_st + 1

ax.text(78, 4, "就绪 (Ready)   |   选取计数: 30   |   100% [ ———●—— ]", fontsize=8, color='#605E5C', va='center')

def draw_callout(x_start, y_start, x_text, y_text, text, color='#C00000'):
    ax.annotate(text, xy=(x_start, y_start), xytext=(x_text, y_text),
                arrowprops=dict(facecolor=color, edgecolor=color, arrowstyle='->', lw=1.5),
                bbox=dict(boxstyle='round,pad=0.3', facecolor='#FFF2CC', edgecolor=color, lw=1),
                fontsize=8.5, fontweight='bold', color='#B25900')

draw_callout(12, 91, 15, 99, "功能区标签页 (Ribbon Tabs)\n切换至「资料」选项卡", '#0070C0')
draw_callout(42, 80, 52, 85, "资料工具核心按键群:\n资料剖析、移除重复项、快速填入", '#107C41')
draw_callout(9, 70.6, 3, 76, "名称框 (Name Box)\n显示当前单元格座标", '#7030A0')
draw_callout(30, 70.6, 40, 76, "资料编辑列 (Formula Bar)\n检视真实公式与底层数据", '#004E8C')
draw_callout(20, 42, 28, 45, "重复与脏资料 (前后空格、格式混杂)", '#C00000')

plt.tight_layout()
plt.savefig("images/fig_1_1_excel_interface_overview.png", dpi=200, bbox_inches='tight')
plt.close()
print("Generated images/fig_1_1_excel_interface_overview.png")

# Figure 2-1: Tidy Data vs Messy Data
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 5.5), dpi=200)

for ax in [ax1, ax2]:
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 100)
    ax.axis('off')

ax1.set_title("❌ 常见航运混乱资料结构 (Messy Data / 难以机器分析)", fontsize=11, fontweight='bold', color='#C00000', pad=10)
r1 = patches.Rectangle((2, 2), 96, 94, facecolor='#FFF7F7', edgecolor='#C00000', linewidth=1.5)
ax1.add_patch(r1)

bad_draws = [
    patches.Rectangle((5, 78), 90, 10, facecolor='#E7E6E6', edgecolor='#A6A6A6'),
    patches.Rectangle((5, 62), 40, 14, facecolor='#FFD9D9', edgecolor='#C00000'),
    patches.Rectangle((47, 62), 48, 14, facecolor='#FFFFFF', edgecolor='#A6A6A6'),
    patches.Rectangle((5, 46), 90, 14, facecolor='#FFE699', edgecolor='#B25900'),
    patches.Rectangle((5, 30), 90, 14, facecolor='#FFFFFF', edgecolor='#A6A6A6'),
    patches.Rectangle((5, 14), 90, 14, facecolor='#E2EFDA', edgecolor='#385723')
]
for p in bad_draws:
    ax1.add_patch(p)

ax1.text(50, 83, "大洋物流2026年11月船期与海运费汇总表 (合并储存档多行表头)", fontsize=9, fontweight='bold', ha='center', color='#323130')
ax1.text(25, 69, "⚠️ 跨行跨列合并\n(上海港直航组)", fontsize=8.5, ha='center', va='center', color='#C00000', fontweight='bold')
ax1.text(71, 69, "柜号/箱型挤在同格: TEMU192830(40HC)\n运费: $2300 含打托费", fontsize=7.5, ha='center', va='center')
ax1.text(50, 53, "⛔ 小计行: 本港小计 12 柜 / $28,400 (破坏资料库连续性)", fontsize=8, ha='center', va='center', color='#B25900', fontweight='bold')
ax1.text(50, 37, "空格与隐藏空白字符 / 日期写成 '11月5号左右' / 空白行隔开", fontsize=8, ha='center', va='center', color='#7030A0')
ax1.text(50, 21, "⛔ 总计行: 统计总额 $158,000 (不可直接建立枢纽分析表)", fontsize=8, ha='center', va='center', color='#385723', fontweight='bold')

ax2.set_title("✅ 标准整洁航运资料结构 (Tidy Data / 适于自动化分析)", fontsize=11, fontweight='bold', color='#107C41', pad=10)
r2 = patches.Rectangle((2, 2), 96, 94, facecolor='#F4FBF7', edgecolor='#107C41', linewidth=1.5)
ax2.add_patch(r2)

headers_tidy = ["提单号", "柜号", "箱型", "POL", "POD", "ETD", "海运费(USD)"]
w_tidy = 90 / len(headers_tidy)
x_th = 5
for h in headers_tidy:
    r_th = patches.Rectangle((x_th, 80), w_tidy, 10, facecolor='#274E13', edgecolor='#FFFFFF')
    ax2.add_patch(r_th)
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
        r_c = patches.Rectangle((x_th, y_tr), w_tidy, 8, facecolor='white', edgecolor='#D9D9D9')
        ax2.add_patch(r_c)
        ax2.text(x_th + w_tidy/2, y_tr + 4, c_val, fontsize=7.5, ha='center', va='center', color='#323130')
        x_th += w_tidy
    y_tr -= 8.5

r_rules = patches.Rectangle((5, 5), 90, 22, facecolor='#E2EFDA', edgecolor='#385723', linewidth=1)
ax2.add_patch(r_rules)
ax2.text(50, 22, "★ 整洁数据三原则 (Codd & Wickham Tidy Data Rules):", fontsize=8.5, fontweight='bold', ha='center', color='#274E13')
ax2.text(8, 16, "1. 每一列 (Column) 代表唯一的业务变量 (如纯粹的箱型、纯粹的港口代码)。", fontsize=7.5, color='#323130')
ax2.text(8, 11, "2. 每一行 (Row) 代表单次具体的运输观测纪录 (单一货柜或单一提单业务)。", fontsize=7.5, color='#323130')
ax2.text(8, 6, "3. 无跨行合并格、无夹杂小计合计行、无复合多重含义栏位、数值纯净无文字单位。", fontsize=7.5, color='#323130')

plt.tight_layout()
plt.savefig("images/fig_2_1_tidy_vs_messy_data.png", dpi=200, bbox_inches='tight')
plt.close()
print("Generated images/fig_2_1_tidy_vs_messy_data.png")

# Figure 4-1: Text to Columns Wizard Diagram
fig, ax = plt.subplots(figsize=(10, 6), dpi=200)
ax.set_xlim(0, 100)
ax.set_ylim(0, 100)
ax.axis('off')

dlg = patches.Rectangle((5, 5), 90, 90, facecolor='#F3F3F3', edgecolor='#0078D7', linewidth=2)
ax.add_patch(dlg)

dlg_h = patches.Rectangle((5, 87), 90, 8, facecolor='#0078D7')
ax.add_patch(dlg_h)
ax.text(8, 91, "资料剖析精灵 - 步骤 3 之 2 (分界符号与复合栏位拆分)", color='white', fontsize=11, fontweight='bold', va='center')

ax.text(10, 81, "【应用情境】将「船名航次 (EVER GIVEN / 0118E)」依照斜杠 '/' 拆分为「船名」与「航次」两栏", fontsize=9.5, fontweight='bold', color='#1E4E79')

box_delim = patches.Rectangle((10, 48), 38, 28, facecolor='white', edgecolor='#A6A6A6')
ax.add_patch(box_delim)
ax.text(12, 72, "请勾选栏位分界符号 (Delimiters):", fontsize=8.5, fontweight='bold')
ax.text(15, 66, "□ Tab 键 (T)", fontsize=8)
ax.text(15, 61, "□ 分号 (S)", fontsize=8)
ax.text(15, 56, "□ 逗号 (C)", fontsize=8)
ax.text(15, 51, "☑ 其他 (O): [  /  ]  <-- 输入航运常用斜杠", fontsize=8.5, fontweight='bold', color='#C00000')

box_opt = patches.Rectangle((52, 48), 38, 28, facecolor='white', edgecolor='#A6A6A6')
ax.add_patch(box_opt)
ax.text(54, 72, "处理选项设置:", fontsize=8.5, fontweight='bold')
ax.text(56, 64, "☑ 将连续分界符号视为单一符号", fontsize=8)
ax.text(56, 56, "文字识别符号: [  \"  ]", fontsize=8)

box_prev = patches.Rectangle((10, 15), 80, 29, facecolor='white', edgecolor='#A6A6A6')
ax.add_patch(box_prev)
ax.text(12, 40, "资料预览 (Data Preview):", fontsize=8.5, fontweight='bold')

patches_col1 = patches.Rectangle((12, 17), 38, 20, facecolor='#EDF5E1', edgecolor='#8A8886')
patches_col2 = patches.Rectangle((51, 17), 37, 20, facecolor='#E2EFDA', edgecolor='#8A8886')
ax.add_patch(patches_col1)
ax.add_patch(patches_col2)

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
print("Generated images/fig_4_1_text_to_columns.png")

# Figure 4-2: Remove Duplicates & Data Validation Dialogue
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 6), dpi=200)

ax1.set_xlim(0, 100)
ax1.set_ylim(0, 100)
ax1.axis('off')

dlg1 = patches.Rectangle((5, 5), 90, 90, facecolor='#F3F3F3', edgecolor='#0078D7', linewidth=2)
ax1.add_patch(dlg1)
dlg1_h = patches.Rectangle((5, 87), 90, 8, facecolor='#0078D7')
ax1.add_patch(dlg1_h)
ax1.text(8, 91, "移除重复项 (Remove Duplicates)", color='white', fontsize=10.5, fontweight='bold', va='center')

ax1.text(8, 80, "若要删除重复值，请选定包含重复项的栏位:", fontsize=8.5)
ax1.text(8, 73, "☑ 我的资料具有标题 (My data has headers)", fontsize=8.5, fontweight='bold', color='#107C41')

box_cols = patches.Rectangle((8, 22), 84, 48, facecolor='white', edgecolor='#A6A6A6')
ax1.add_patch(box_cols)
cols_dup = [
    ("□ (全选)", "取消全选"),
    ("☑ 提单号 (B/L No.)", "★ 关键唯一主键"),
    ("☑ 集装箱号 (Container No.)", "★ 货柜物理唯一标识"),
    ("□ 箱型 (Type)", ""),
    ("□ 船名航次", ""),
    ("□ 装运港 / 目的港", "")
]
y_cd = 63
for c_name, c_hint in cols_dup:
    is_key = "★" in c_hint
    ax1.text(12, y_cd, c_name, fontsize=8, fontweight='bold' if is_key else 'normal', color='#C00000' if is_key else '#323130')
    if c_hint:
        ax1.text(55, y_cd, c_hint, fontsize=7.5, color='#B25900' if is_key else '#605E5C', fontweight='bold' if is_key else 'normal')
    y_cd -= 7.5

ax1.text(8, 14, "💡 航运实务要诀: 依「提单号+柜号」作为判别基准，\n可精准清除重复录入，保留首次录入的合法单证。", fontsize=7.5, color='#1E4E79')
ax1.text(65, 8, "[ 确定 (OK) ]   [ 取消 ]", fontsize=8.5, fontweight='bold')

ax2.set_xlim(0, 100)
ax2.set_ylim(0, 100)
ax2.axis('off')

dlg2 = patches.Rectangle((5, 5), 90, 90, facecolor='#F3F3F3', edgecolor='#107C41', linewidth=2)
ax2.add_patch(dlg2)
dlg2_h = patches.Rectangle((5, 87), 90, 8, facecolor='#107C41')
ax2.add_patch(dlg2_h)
ax2.text(8, 91, "资料验证 (Data Validation) - 规范源头录入", color='white', fontsize=10.5, fontweight='bold', va='center')

ax2.text(10, 81, "[设定 (Settings)]   [输入讯息]   [错误警讯]", fontsize=8.5, fontweight='bold', color='#107C41')

box_set = patches.Rectangle((8, 22), 84, 55, facecolor='white', edgecolor='#A6A6A6')
ax2.add_patch(box_set)

ax2.text(12, 70, "验证准则 (Validation criteria):", fontsize=8.5, fontweight='bold')
ax2.text(12, 62, "允许 (Allow):  [ 清单 (List)                 ▼ ]", fontsize=8.5, fontweight='bold', color='#004E8C')
ax2.text(12, 53, "☑ 忽略空白 (Ignore blank)", fontsize=8)
ax2.text(12, 46, "☑ 储存格内提供下拉式清单 (In-cell dropdown)", fontsize=8, fontweight='bold', color='#107C41')

ax2.text(12, 36, "来源 (Source):", fontsize=8.5, fontweight='bold')
r_src = patches.Rectangle((12, 26), 76, 7, facecolor='#F2F2F2', edgecolor='#7F7F7F')
ax2.add_patch(r_src)
ax2.text(14, 29.5, "=基础代码参照表!$B$2:$B$19  (或 20GP,40GP,40HC)", fontsize=8, fontfamily='monospace', color='#C00000', fontweight='bold')

ax2.text(8, 14, "💡 航运实务要诀: 锁定港口代码与箱型选项，\n从源头杜绝 CNSHA 误打成 cn-sha、sha 等混乱资料。", fontsize=7.5, color='#1E4E79')
ax2.text(65, 8, "[ 确定 (OK) ]   [ 取消 ]", fontsize=8.5, fontweight='bold')

plt.tight_layout()
plt.savefig("images/fig_4_2_remove_duplicates_and_validation.png", dpi=200, bbox_inches='tight')
plt.close()
print("Generated images/fig_4_2_remove_duplicates_and_validation.png")

# Figure 5-1: 300-row Practice Pipeline
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
    rect_st = patches.Rectangle((x_pos, 22), 17.5, 62, facecolor=bg_color, edgecolor=border_color, linewidth=2, linestyle='-')
    ax.add_patch(rect_st)
    rect_sth = patches.Rectangle((x_pos, 72), 17.5, 12, facecolor=border_color)
    ax.add_patch(rect_sth)
    ax.text(x_pos + 8.75, 78, s_title, color='white', fontsize=8, fontweight='bold', ha='center', va='center')
    ax.text(x_pos + 1.2, 66, s_desc, fontsize=7.5, color='#323130', va='top', linespacing=1.6)
    if idx < len(stages) - 1:
        ax.annotate('', xy=(x_pos + 19.5, 52), xytext=(x_pos + 17.7, 52),
                    arrowprops=dict(facecolor='#1F497D', edgecolor='#1F497D', width=2, headwidth=6, headlength=6))
    x_pos += 19.6

rect_sum = patches.Rectangle((2, 5), 96, 12, facecolor='#1F497D', edgecolor='#1F497D')
ax.add_patch(rect_sum)
ax.text(50, 11, "🎯 目标达成: 从 ERP/EDI 导出混杂脏资料 -> 蜕变为零错误、标准统一、即开即用的高品质航运资料库", 
        color='white', fontsize=9.5, fontweight='bold', ha='center', va='center')

plt.tight_layout()
plt.savefig("images/fig_5_1_practice_pipeline.png", dpi=200, bbox_inches='tight')
plt.close()
print("Generated images/fig_5_1_practice_pipeline.png")

print("All datasets and images generated successfully!")
