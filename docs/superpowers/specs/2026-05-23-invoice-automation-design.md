# 发票自动化工具 - 设计文档

## 概述

Python CLI 工具，实现发票PDF文件的自动重命名和信息校验。全程离线运行，使用 PaddleOCR 进行本地 OCR 识别。

## 功能

### 任务1：文件重命名

根据合同号索引Excel和OCR识别结果，自动获取甲方合同号，对PDF文件重命名。

**文件类型识别：**

| 类型 | 文件名模式 | 处理方式 |
|------|-----------|---------|
| Type 1 | `2100002601080A_xxx.pdf` | 乙方合同号在文件名中，直接查索引 |
| Type 2 | `dzfp_26317000001473818238_xxx.pdf` | OCR识别备注栏获取乙方合同号，再查索引 |

**重命名规则：**
- Type 1: `{甲方合同号}-{乙方合同号}_{原文件名中乙方合同号之后的部分}.pdf`
  - 例：`2100002601080A_引望智能技术有限公司_20260519.pdf` → `7871Y00200172-2100002601080A_引望智能技术有限公司_20260519.pdf`
- Type 2: `{甲方合同号}-{乙方合同号}_{发票号码}_{原文件名中发票号码之后的部分}.pdf`
  - 例：`dzfp_26317000001473818238_引望智能技术有限公司_20260519.pdf` → `7871Y00200172-2100002601080A_26317000001473818238_引望智能技术有限公司_20260519.pdf`

### 任务2：发票信息校验

比对Excel中填写的发票信息与PDF OCR识别结果，自动修复错误并生成报告。

**校验字段（Excel列映射）：**

| Excel列 | 字段 | 含义 |
|---------|------|-----|
| AF | party_a_id | 甲方合同号 |
| BZ | billing_qty | 本次开票数量 |
| CA | invoice_no | 发票号 |
| CB | invoice_date | 发票日期 |
| CC | tax_amount | 税金金额 |
| CD | total_amount | 含税金额 |

**校验逻辑：**
- 数据从Excel第4行开始（前3行为表头）
- 通过甲方合同号匹配Excel行与PDF文件
- OCR置信度 > 0.8 时自动修复Excel
- OCR置信度 <= 0.8 时标记为需手动核查
- 生成校验报告，列出所有差异和修复操作

## 技术方案

### OCR策略（方案B）

优先用 PyMuPDF 提取PDF文本层，失败则 fallback 到 PaddleOCR：

1. `fitz.open()` 打开PDF，`page.get_text()` 提取文本
2. 如果提取到足够文本（>50字），认为有文本层，直接返回
3. 否则 fallback：PDF转图片 → PaddleOCR识别

### 甲方合同号提取规则

从合同名称（`YW-xxx` 格式）中提取甲方合同号：

1. 去掉 `YW-` 前缀
2. 如果剩余部分以 `SZ` 开头 → 取到第二个 `-`（含）
3. 否则 → 取到第一个 `-`（不含）

示例：
- `YW-7871Y00200172-0C1911-651400P` → `7871Y00200172`
- `YW-SZYG4502802-4-S292912F-3000P` → `SZYG4502802-4`

### 发票字段提取

从OCR文本中用正则匹配：
- 发票号码：`发票号码[:：]\s*(\d{20})`
- 开票日期：`开票日期[:：]\s*(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日)`
- 税额：`税额\s*[:：]?\s*[\¥￥]?\s*([\d,.]+)`
- 价税合计：`价税合计.*?[\¥￥]\s*([\d,.]+)`
- 乙方合同号（备注栏）：备注区域文本中匹配合同号模式

## 项目结构

```
auto_fapiao/
├── main.py              # CLI入口（argparse）
├── ocr.py               # PDF文本提取 + PaddleOCR封装
├── contract.py          # 合同号索引表解析 + 甲方合同号提取
├── rename.py            # 任务1：文件重命名
├── verify.py            # 任务2：发票信息校验
├── excel_utils.py       # Excel读写工具（openpyxl）
├── tests/
│   ├── conftest.py      # pytest fixtures，构造mock数据
│   ├── fixtures/        # 生成的mock Excel和PDF
│   ├── test_ocr.py
│   ├── test_contract.py
│   ├── test_rename.py
│   └── test_verify.py
└── docs/
    └── superpowers/specs/
```

## CLI接口

```bash
# 任务1：重命名
uv run main.py rename --dir ./invoices --excel ./合同号索引表.xlsx

# 任务2：校验
uv run main.py verify --dir ./invoices --excel ./发票验证表.xlsx

# 生成测试数据
uv run main.py generate-test-data --output ./test_data
```

## 依赖

- `paddleocr` — 离线OCR
- `paddlepaddle` — PaddleOCR推理后端
- `PyMuPDF (fitz)` — PDF文本提取 + PDF转图片
- `openpyxl` — Excel读写
- `reportlab` — 生成mock PDF测试数据
- `pytest` — 测试

## 数据结构

```python
@dataclass
class RenameResult:
    original_name: str
    new_name: str
    party_a_id: str
    party_b_id: str
    status: str  # "success" | "skipped" | "error"
    message: str

@dataclass
class FieldDiff:
    field_name: str        # "发票号" | "发票日期" | "税额" | "含税金额"
    excel_value: str
    ocr_value: str
    confidence: float      # PaddleOCR置信度，0-1
    fixed: bool            # 是否已自动修复

@dataclass
class VerifyResult:
    pdf_file: str
    party_a_id: str
    diffs: list[FieldDiff]
    fixed: bool
    needs_manual: bool
```

## 约束

- 全程离线运行，不联网，不调用在线API
- 使用本地CPU进行OCR推理
- Python 3.12，使用 uv 管理依赖
