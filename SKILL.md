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
## 前置：环境检查

检查依赖是否已安装，通过则写入缓存标记 `~/.cache/baoxiao/.env_checked`，后续跳过。

```bash
if [ ! -f ~/.cache/baoxiao/.env_checked ]; then
    读取 references/env-setup.md 开展环境检查工作
fi
```

## 执行流程

采用三阶段模型： 提取 → 写入 → 合并


### 阶段1: 数据提取（--extract-only）

1. OFD转PDF（使用ofd2pdf，失败则停止）
2. 发票类型识别（按内容关键词分类）
3. 数据提取
   - 金额：价税合计小写金额
   - 日期：航班/车次日期（OCR+pdfplumber双重校验）
   - 城市：出发和到达城市（自动排除出发地）
4. 数据校验
   - 金额合理性（机票¥500-5000，火车¥50-2000等）
   - 日期合理性（OCR常见错误：11月→1月，29日→9昌）
   - 日期交叉验证（多张机票日期顺序）
5. 生成 `data_report.json`

```bash
python3 ~/.claude/skills/baoxiao/scripts/invoice_processor.py \
  --input-dir 发票 --extract-only
```

### 阶段2: 数据写入（--write-data）

**前置检查**：若 `data_report.json` 中有 `unknown_cities`（即城市未在config.yaml中配置），**必须暂停询问用户**：
- 选项1：新增城市映射 → 添加到 `config.yaml`
- 选项2：修正城市名称 → 修改 `data_report.json`（如"纺阳"→"沈阳"）
- 选项3：删除该城市 → 从 `data_report.json` 移除

**说明**：若所有城市都已在 `config.yaml` 中配置，则无需 `--city-units` 参数，直接执行即可。

解决后才能继续：

1. 填写Excel报销表（城市→E4，火车→E6/F6，机票→E7/F7/J4/M4等）
2. 填写Word审批文档（地点、到达单位、起止时间）
3. 自动复制sheet1日期到sheet2（D3和F3）

```bash
# 方式1：城市已在config.yaml中配置
python3 ~/.claude/skills/baoxiao/scripts/invoice_processor.py \
  --output-excel biaoge.xlsx --work-dir . --write-data

# 方式2：通过命令行临时指定城市单位映射（优先级高于config.yaml）
python3 ~/.claude/skills/baoxiao/scripts/invoice_processor.py \
  --output-excel biaoge.xlsx --work-dir . --write-data \
  --city-units "城市A:单位A,城市B:单位B"
```

### 阶段3: PDF转换和合并（--merge-pdfs）

按顺序合并为 `汇总打印.pdf`：
1. biaoge.pdf（报销表）
2. shenpi.pdf（审批文档）
3. jipiao*.pdf（机票）
4. huoche*.pdf（火车票）
5. zhusu*.pdf（住宿）
6. 滴滴*.pdf、交通费*.pdf

```bash
python3 ~/.claude/skills/baoxiao/scripts/invoice_processor.py \
  --input-dir 发票 --work-dir . --merge-pdfs
```

### 一键执行（--auto）

无警告时自动完成三阶段：

```bash
python3 ~/.claude/skills/baoxiao/scripts/invoice_processor.py \
  --input-dir 发票 --output-excel biaoge.xlsx --work-dir . \
  --merge-pdfs --auto --city-units "城市A:单位A"
```

## 关键规则

### 发票分类

| 类型 | 识别关键字 | 输出文件名 |
|------|-----------|-----------|
| 机票 | 航空、航班、承运人、民航、座位等级、登机 | jipiao1.pdf |
| 火车票 | 铁路、车次、二等座、高铁、动车、12306、出发站 | huoche1.pdf |
| 住宿 | 住宿、酒店、宾馆、住宿费、房费、客房 | zhusu1.pdf |
| 滴滴电子发票 | 滴滴、发票号码、价税合计 | 滴滴电子发票A.pdf |
| 滴滴行程单 | 滴滴、行程单、行程记录 | 滴滴出行行程报销单A.pdf |
| 其他交通 | 地铁、轨道交通、公交、出租车、一卡通、乘车码 | 交通费发票A.pdf |
| 退改签 | 退票、改签、变更 | 退改签1.pdf |

### 环境检查缓存

检查通过后写入 `~/.cache/baoxiao/.env_checked`，下次跳过。手动删除可重新验证。

### 重要校验规则

- OCR质量：机票PDF使用高DPI(400+)识别
- 日期陷阱："11月"→"1月"、"29日"→"9昌"需交叉验证
- 金额单位：提取"价税合计"而非"票价"
- 城市识别：去除"南站"、"北站"后缀，只保留城市名
- 滴滴金额：仅统计电子发票，行程单不计入（避免重复）
