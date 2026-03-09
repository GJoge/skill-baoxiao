---
name: baoxiao
description: |
  处理报销发票识别、归类和数据提取的自动化流程。当用户需要处理发票文件、
  填写报销表格、识别发票类型（机票、火车、住宿、滴滴等）、提取发票金额和日期时，
  必须使用此技能。适用于财务报销、差旅费统计、发票管理等场景。

  触发场景：
  - 用户提到"发票"、"报销"、"baoxiao"、"差旅费"等关键词
  - 需要填写报销表格（如biaoge.xlsx）
  - 需要识别PDF发票并提取金额、日期
  - 需要按类型归类发票文件
  - 需要验证发票数据的合理性
---

# 报销发票处理技能 (Baoxiao)

## 概述

本技能提供完整的发票处理流程，包括：
1. OFD文件自动转换为PDF
2. 发票文件自动识别与归类
3. 关键数据提取（金额、日期、城市名称）
4. 数据合理性校验
5. Excel表格自动填写（含城市信息）
6. Word审批文档自动填写
7. PDF转换和合并

## 前置要求

必须安装以下工具,先检查用户环境是否满足,不满足需要执行安装：
```bash
# PDF处理
pip3 install pdfplumber pdf2image pypytesseract pillow openpyxl pandas python-docx reportlab pypdf pypdf2
# YAML配置支持 (用于读取 config.yaml 配置文件)
apt-get install python3-yaml

# OCR引擎
apt-get install tesseract-ocr tesseract-ocr-chi-sim poppler-utils

# PDF转换（用于Excel/Word转PDF）
apt-get install libreoffice-writer libreoffice-calc
```
字体依赖, 识别中文发票需要安装常见的字体，如宋体、楷体、黑体、仿宋、仿宋_GB2312、方正小标宋简体、Arial等.

## 配置说明

### 城市单位映射配置

工具通过**配置文件**或**命令行参数**获取城市到单位的映射，用于自动填写Word审批文档中的"到达单位"字段。

#### 方式1：配置文件（推荐）

1. **创建配置文件**
   ```bash
   cp config.example.yaml config.yaml
   ```

2. **编辑配置内容**
   ```yaml
   # config.yaml
   city_units:
     北京: 总部
     上海: 分公司
     广州: 办事处
     # 根据需要添加更多...
   ```

3. **配置文件位置**（按以下顺序查找）
   - 当前工作目录 `./config.yaml`
   - 脚本目录 `scripts/config.yaml`
   - 用户配置目录 `~/.config/baoxiao/config.yaml`

4. **安全提示**：`config.yaml` 包含敏感信息，**不要**提交到Git仓库（已加入 `.gitignore`）

#### 方式2：命令行参数

在阶段2命令中添加 `--city-units` 参数：

```bash
--city-units "北京:总部,上海:分公司,广州:办事处"
```

**优先级**：命令行参数 > 配置文件 > 空配置

---

## 支持的发票类型

本技能支持识别和处理以下发票类型：

| 类型 | 说明 | 识别关键字 | 输出文件名格式 |
|------|------|-----------|---------------|
| **机票** | 航空运输电子客票行程单 | 航空、航班、承运人、民航、座位等级、登机 | `jipiao1.pdf`, `jipiao2.pdf`... |
| **火车票** | 铁路电子客票 | 铁路、车次、二等座、高铁、动车、12306、出发站 | `huoche1.pdf`, `huoche2.pdf`... |
| **住宿** | 酒店/宾馆住宿费发票 | 住宿、酒店、宾馆、住宿费、房费、客房 | `zhusu1.pdf`, `zhusu2.pdf`... |
| **滴滴电子发票** | 滴滴出行增值税电子普通发票 | 滴滴、发票号码、价税合计 | `滴滴电子发票A.pdf`... |
| **滴滴行程单** | 滴滴出行行程报销单 | 滴滴、行程单、行程记录 | `滴滴出行行程报销单A.pdf`... |
| **其他交通** | 地铁、公交、出租车等 | 地铁、轨道交通、公交、出租车、一卡通、乘车码 | `交通费发票A.pdf`... |
| **退改签** | 机票/火车票退改签费用 | 退票、改签、变更 | `退改签1.pdf`... |
| **未分类** | 无法识别的发票（用户选择跳过） | - | `未分类_1.pdf`... |

**自动归类规则**：
- 地铁、公交、出租车等发票 → 自动归入**其他交通**类别
- 滴滴发票 → 自动区分为**电子发票**（统计金额）和**行程单**（仅归档）
- 无法识别的发票 → 交互式询问用户，非交互环境默认归入其他交通类别

## 三阶段处理流程

本技能采用**三阶段执行模型**，确保数据在写入前经过用户确认：

---

### 阶段1: 数据提取（--extract-only）

从发票文件中提取所有关键数据，生成数据报告供用户复核。

#### 阶段1包含的步骤：

**1. OFD文件转换**
- 自动扫描输入目录中的 `*.ofd` 文件
- 使用 `ofd2pdf` 工具转换为PDF格式
- 跳过已存在的同名PDF（避免重复转换）

**2. 发票类型识别**
- 使用内容分析（非文件名）识别发票类型
- 根据识别结果重命名文件（见上表）

**3. 数据提取**
- **金额**：提取"价税合计"小写金额
- **日期**：提取航班/车次日期（机票日期使用OCR+pdfplumber双重校验）
- **城市**：提取出发和到达城市（自动排除北京）

**4. 数据校验**
- 金额合理性检查（机票¥500-5000，火车¥50-2000等）
- 日期合理性检查（OCR常见错误：11月→1月，29日→9昌）
- 日期交叉验证（多张机票的日期顺序）
- 滴滴文件区分（电子发票vs行程单）

**阶段1输出**：`data_report.json` 数据报告文件

```bash
python3 /root/.claude/skills/baoxiao/scripts/invoice_processor.py \
  --input-dir 发票 \
  --extract-only
```

---

### 阶段2: 数据写入（--write-data）

根据确认的数据报告，写入Excel报销表和Word审批文档。

#### Excel表格填写

| 项目 | 数量单元格 | 金额单元格 | 日期单元格 | 说明 |
|------|-----------|-----------|-----------|------|
| 城市 | E4 | - | - | 多个城市用顿号分隔，如"成都、绵阳" |
| 火车 | E6 | F6 | - | 数量和金额 |
| 机票 | E7 | F7 | J4(最早), M4(最晚) | 数量和金额，日期写入J4/M4 |
| 退改 | E8 | F8 | - | 退改签费用 |
| 滴滴 | E10 | F10 | - | 仅统计电子发票金额 |
| 住宿 | - | O9 | - | 住宿费用总额 |

同时自动将sheet1的日期复制到sheet2（D3和F3）。

#### Word审批文档填写

- **地点**：填写城市名称（如"地点：成都、绵阳"）
- **到达单位**：根据城市自动填写对应单位
  - 从配置文件 `config.yaml` 读取城市到单位的映射
  - 或通过 `--city-units` 参数指定，如 `"北京:总部,上海:分公司"`
- **起止时间**：填写最早日期至最晚日期

**阶段2命令**：

```bash
python3 /root/.claude/skills/baoxiao/scripts/invoice_processor.py \
  --output-excel biaoge.xlsx \
  --work-dir . \
  --write-data \
  --city-units "北京:总部,上海:分公司,广州:办事处"
```

---

### 阶段3: PDF转换和合并（--merge-pdfs）

将已填写的文档转换为PDF，并与发票PDF按顺序合并。

#### 阶段3包含的步骤：

**1. Excel转PDF**
- 将 `biaoge.xlsx` 转换为 `biaoge.pdf`

**2. Word转PDF**
- 将 `shenpi.docx` 转换为 `shenpi.pdf`

**3. 收集PDF文件**
按以下顺序收集：
1. `biaoge.pdf`（报销表格）
2. `shenpi.pdf`（审批文件）
3. `jipiao*.pdf`（机票发票）
4. `huoche*.pdf`（火车票）
5. `zhusu*.pdf`（住宿发票）
6. `滴滴*.pdf`、`交通费*.pdf`（滴滴和交通发票）

**4. 合并PDF**
- 将所有PDF合并为 `汇总打印.pdf`

**阶段3命令**：

```bash
python3 /root/.claude/skills/baoxiao/scripts/invoice_processor.py \
  --input-dir 发票 \
  --work-dir . \
  --merge-pdfs
```

---

### 一键自动执行（--auto）

如果没有警告，自动完成所有三个阶段；如果有警告，则停止等待复核：

```bash
python3 /root/.claude/skills/baoxiao/scripts/invoice_processor.py \
  --input-dir 发票 \
  --output-excel biaoge.xlsx \
  --work-dir . \
  --merge-pdfs \
  --auto \
  --city-units "北京:总部,上海:分公司"
```

## 参数说明

| 参数 | 说明 | 适用阶段 |
|------|------|---------|
| `--input-dir` | 发票文件夹路径 | 阶段1、阶段3 |
| `--output-excel` | 输出的Excel文件路径 | 阶段2、阶段3（可选） |
| `--sheet` | Excel工作表名称 | 阶段2 |
| `--work-dir` | 工作目录（存放docx等文件） | 阶段2、阶段3 |
| `--report` | 数据报告文件路径 | 阶段1、阶段2（默认：`data_report.json`） |
| `--city-units` | 城市到单位的映射，格式：`"北京:总部,上海:分公司"` | 阶段2 |
| `--no-rename` | 不自动重命名发票文件 | 阶段1 |
| `--extract-only` | 仅执行阶段1（提取数据） | - |
| `--write-data` | 仅执行阶段2（写入数据） | - |
| `--merge-pdfs` | 仅执行阶段3（PDF转换和合并） | - |
| `--auto` | 自动模式（无警告时自动完成全部阶段） | - |

**说明**：
- `--output-excel` 参数为**可选**。如果不提供，脚本会自动查找工作目录中的 `biaoge.xlsx`
- 如果 `biaoge.xlsx` 不存在但 `biaoge.pdf` 已存在，则跳过Excel转换直接使用PDF
- 必须提供 `--input-dir` 才能收集发票PDF进行合并

### 输出文件

| 文件 | 说明 | 生成阶段 |
|------|------|---------|
| `data_report.json` | 数据报告，包含提取的城市、日期、金额等 | 阶段1 |
| `biaoge.pdf` | Excel报销表转换的PDF | 阶段3 |
| `shenpi.pdf` | Word审批文档转换的PDF（如果存在shenpi.docx） | 阶段3 |
| `汇总打印.pdf` | 所有PDF按顺序合并后的最终文件 | 阶段3 |
| `jipiao*.pdf` 等 | 重命名后的发票文件 | 阶段1 |

**data_report.json 示例：**
```json
{
  "cities": ["成都", "绵阳"],
  "cities_str": "成都、绵阳",
  "dates": { "earliest": "2025-11-26", "latest": "2025-11-28" },
  "amounts": { "jipiao": 1250.00, "huoche": 450.00, "didi": 120.50, "zhusu": 800.00 },
  "counts": { "jipiao": 2, "huoche": 1, "didi": 5 },
  "warnings": [],
  "needs_review": false
}
```

## 完整执行示例（分阶段）

```python
# 1. 导入技能脚本
import sys
sys.path.insert(0, '/root/.claude/skills/baoxiao/scripts')
from invoice_processor import process_invoices, generate_data_report, load_data_report

# 2. 阶段1: 提取数据（不写入Excel）
result = process_invoices(
    input_dir='发票',
    output_excel=None,  # 阶段1不写入
    sheet_name='sheet1',
    force_write=True
)

# 3. 生成数据报告
report = generate_data_report(result, 'data_report.json')

# 4. 检查是否需要复核
if report['needs_review']:
    print("⚠️ 发现以下问题需要人工复核:")
    for warning in report['warnings']:
        print(f"  - {warning}")
    print("\n请修改 data_report.json 后，再执行阶段2")
else:
    print("✓ 数据无误，可以执行阶段2写入文档")

# 5. 阶段2: 确认数据后写入（在另一个脚本中执行）
# loaded = load_data_report('data_report.json')
# write_to_excel('biaoge.xlsx', 'sheet1', loaded, loaded['dates'])
```

## 注意事项

### 三阶段执行流程

脚本采用**先提取、后写入**的三阶段设计：
1. **阶段1** 只提取数据，生成 `data_report.json`，**不修改任何文档**
2. **阶段2** 根据用户确认的数据报告，写入Excel和Word文档
3. **阶段3** 转换文档为PDF并合并

这种设计解决了Python非交互式执行无法暂停等待用户输入的问题。用户可以在阶段1和阶段2之间复核并修改 `data_report.json` 文件。

### 各阶段详细说明

**阶段1 - 数据提取**：
- 自动完成OFD转PDF、发票识别、数据提取、数据校验
- 生成 `data_report.json` 供用户复核
- 如果 `needs_review` 为 `true`，建议先检查警告再执行阶段2

**阶段2 - 数据写入**：
- 读取 `data_report.json` 中的确认数据
- 填写Excel报销表和Word审批文档
- 自动复制日期到sheet2
- 需要指定城市到单位的映射（如 `--city-units "北京:总部,上海:分公司"`）

**阶段3 - PDF合并**：
- 需要 `--input-dir` 参数才能收集发票PDF
- 按固定顺序合并：报销表 → 审批文档 → 机票 → 火车票 → 住宿 → 滴滴/交通
- 输出 `汇总打印.pdf` 可直接打印提交

### 其他重要提示

1. **OCR质量**: 机票PDF通常是扫描件，必须使用高DPI(400+)识别
2. **日期陷阱**:
   - "11月"常被误识为"1月"
   - "29日"常被误识为"9昌"
   - 必须交叉验证多张机票的日期顺序
3. **金额单位**: 注意区分"票价"和"价税合计"，应提取后者
4. **城市识别**: 火车票站名会去除"南站"、"北站"等后缀，只保留城市名
5. **单位映射**: 从配置文件 `config.yaml` 读取城市到单位的映射，或通过 `--city-units` 参数指定
6. **文件备份**: 重命名前自动创建备份到 `/tmp/invoice_backup/`
7. **滴滴金额**: 仅统计电子发票金额，行程单金额不计入（避免重复）
8. **发票类型识别**:
   - 地铁、公交、出租车等交通类发票会自动归入滴滴/交通类别
   - 无法识别的发票默认归入滴滴/交通类别
   - 用户可选择跳过不相关的发票文件（标记为"未分类"）
