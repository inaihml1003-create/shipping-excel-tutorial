# 航运行政人员 Excel 基础资料分析：资料前处理核心技能实务教程

本项目为专为航运物流行政人员打造的完整教学专案，涵盖 MS Excel 资料前处理核心技能教程（Word 教材、PPT 简报、互动式网页简报）以及两份配套实训活页簿（30 笔实务示范档、300 笔多场景情境练习档）。

---

## 📂 项目包含档案清单

| 档案名称 | 格式 | 档案大小 | 说明 |
| :--- | :---: | :---: | :--- |
| **`航运行政人员_Excel基础资料分析与资料前处理教学手册.docx`** | Word (DOCX) | ~1.47 MB | 完整出版级教学手册（含专业封面、章节目录、图表目录、页码、公式示范与参考文献） |
| **`航运行政人员_Excel基础资料分析与资料前处理教学手册.pptx`** | PPT (PPTX) | ~1.46 MB | 16:9 宽萤幕 13 页高阶商务投影简报（含架构图、对齐表格与步骤指引） |
| **`shipping_excel_tutorial_presentation.html`** | HTML | ~98 KB | 互动式 PPT 简报网页（支援全萤幕播放、键盘翻页、公式一键复制、实验前后对比） |
| **`航运资料前处理_30笔示范档.xlsx`** | Excel (XLSX) | ~14 KB | 30 笔原始脏数据、清洗后标准数据对照及处理步骤映射 |
| **`航运资料前处理_300笔情境练习档.xlsx`** | Excel (XLSX) | ~51 KB | 300 笔涵盖集装箱流转、运费滞期、船期准点率三大场景实训档与考核题解 |

---

## 🎯 教学核心大纲

1. **第 1 章：Excel 操作界面全景与核心区域解析**
   - 标题列、快速存取工具列、「资料 (Data)」功能区群组、资料编辑列、名称框与工作表网格。
   - 航运行政高频快捷键速查（Ctrl+Shift+方向键、Ctrl+E、Ctrl+H、F4 等）。
2. **第 2 章：资料前处理核心常识与整洁数据 (Tidy Data)**
   - 计算机科学 GIGO 原则（垃圾进，垃圾出）在航运运费结算与报关中的代价。
   - Excel 五大资料类型（文字、数值、日期时间、逻辑值、公式与错误值）判别与避坑。
   - Hadley Wickham 整洁数据三大黄金法则与航运 4 大脏数据根源。
3. **第 3 章：30 笔航运实务示范档全流程演练**
   - 原始脏数据剖析、体检诊断清册。
   - 8 步标准化清洗工序（备份 ➔ 双主键去重 ➔ 文本净化 ➔ 资料剖析 ➔ 日期标准化 ➔ 数值化 ➔ XLOOKUP 码表关联 ➔ 固化数值）。
4. **第 4 章：资料前处理关键功能选项与核心函数**
   - 资料剖析 (Text to Columns)、移除重复项 (Remove Duplicates)、资料验证下拉选单 (Data Validation)、快速填入 (Flash Fill)。
   - 核心函数：`TRIM`, `CLEAN`, `UPPER`, `TEXT`, `VALUE`, `LEFT/MID/RIGHT`, `DATE/DATEVALUE`, `XLOOKUP`, `IFERROR`。
   - 航运三大万能嵌套公式模板。
5. **第 5 章：300 笔多场景综合实战练习与题解**
   - 场景 A：集装箱装箱单流转与动态跟踪（100 笔）
   - 场景 B：海运运杂费核算与超期滞箱费结算（100 笔）
   - 场景 C：船期表准点率与港口泊位作业监控（100 笔）
   - 十大核心考核任务手把手题解与标准答案校验。
6. **第 6 章：品质管控、自动化与日常铁律**
   - 源头表单刚性防错机制设计（下拉选单、格式掩码、Ctrl+T 表格）。
   - Power Query (M 语言) 自动清洗与 Python in Excel 进阶展望。
   - 航运行政日常数据处理十条铁律。

---

## 📚 参考文献与行业标准引证

- 国际海事组织 (IMO). 《国际便利海上运输公约》(FAL Convention)
- 国际标准化组织 (ISO). ISO 6346: Freight containers — Coding, identification and marking
- 联合国欧洲经济委员会 (UNECE). UN/LOCODE: United Nations Code for Trade and Transport Locations
- 威克姆, 哈德利 (Hadley Wickham). Tidy Data (整洁数据理论), Journal of Statistical Software, 2014.
- 微软官方技术支持库 (Microsoft Support). Excel 数据转换与函数技术规范, 2026.

---

## 👤 作者与专案资讯
- **GitHub 账户**：`inaihml1003@gmail.com`
- **制作工具**：Python 3.15, python-docx, python-pptx, openpyxl, Tailwind CSS
