# 发票自动化工具 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Python CLI tool that renames invoice PDFs based on contract number lookup and verifies invoice data against Excel records using offline OCR.

**Architecture:** Modular Python project with separate modules for OCR, contract parsing, renaming, verification, and Excel I/O. Uses PyMuPDF for text extraction with PaddleOCR fallback. CLI via argparse.

**Tech Stack:** Python 3.12, PaddleOCR, PyMuPDF (fitz), openpyxl, reportlab, pytest, uv

---

## File Map

| File | Responsibility |
|------|---------------|
| `pyproject.toml` | Project metadata and dependencies |
| `main.py` | CLI entry point (argparse), dispatches to rename/verify/generate-test-data |
| `ocr.py` | PDF text extraction (PyMuPDF) + PaddleOCR fallback + invoice field extraction |
| `contract.py` | Load contract index Excel, extract 甲方合同号 from 合同名称 |
| `rename.py` | Classify PDF type, rename files using contract index |
| `verify.py` | Compare Excel vs OCR data, auto-fix, generate report |
| `excel_utils.py` | Read/write Excel with openpyxl |
| `tests/test_contract.py` | Tests for contract number extraction and index loading |
| `tests/test_ocr.py` | Tests for text extraction and invoice field parsing |
| `tests/test_rename.py` | Tests for file classification and renaming |
| `tests/test_verify.py` | Tests for field comparison and auto-fix |
| `tests/conftest.py` | Shared fixtures: mock Excel files, mock PDFs |

---

### Task 1: Project Setup + Dependencies

**Files:**
- Modify: `pyproject.toml`

- [ ] **Step 1: Update pyproject.toml with dependencies**

```toml
[project]
name = "auto-fapiao"
version = "0.1.0"
description = "发票自动重命名和校验工具"
readme = "README.md"
requires-python = ">=3.12"
dependencies = [
    "paddleocr>=2.7",
    "paddlepaddle>=2.5",
    "pymupdf>=1.24",
    "openpyxl>=3.1",
    "reportlab>=4.0",
]

[project.optional-dependencies]
dev = [
    "pytest>=8.0",
    "pytest-cov",
]
```

- [ ] **Step 2: Install dependencies**

Run: `uv sync`
Expected: Dependencies installed successfully

- [ ] **Step 3: Create tests directory structure**

Run: `mkdir -p tests/fixtures`
Expected: Directory created

- [ ] **Step 4: Commit**

```bash
git add pyproject.toml uv.lock
git commit -m "chore: add project dependencies"
```

---

### Task 2: Contract Number Parsing

**Files:**
- Create: `contract.py`
- Create: `tests/test_contract.py`

- [ ] **Step 1: Write failing test for extract_party_a_contract**

```python
# tests/test_contract.py
from contract import extract_party_a_contract


def test_extract_non_sz():
    """非SZ开头的合同号，取到第一个-"""
    result = extract_party_a_contract("YW-7871Y00200172-0C1911-651400P")
    assert result == "7871Y00200172"


def test_extract_sz():
    """SZ开头的合同号，取到第二个-（含）"""
    result = extract_party_a_contract("YW-SZYG4502802-4-S292912F-3000P")
    assert result == "SZYG4502802-4"


def test_extract_no_prefix():
    """没有YW-前缀的情况，仍然正常处理"""
    result = extract_party_a_contract("7871Y00200172-0C1911-651400P")
    assert result == "7871Y00200172"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_contract.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'contract'`

- [ ] **Step 3: Write minimal implementation**

```python
# contract.py
import re


def extract_party_a_contract(full_name: str) -> str:
    """从合同名称中提取甲方合同号。

    规则：
    1. 去掉 YW- 前缀（如有）
    2. 如果剩余部分以 SZ 开头 → 取到第二个 -（含）
    3. 否则 → 取到第一个 -（不含）
    """
    name = re.sub(r"^YW-", "", full_name)
    if name.startswith("SZ"):
        first_dash = name.index("-")
        second_dash = name.index("-", first_dash + 1)
        return name[:second_dash]
    else:
        return name.split("-")[0]
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_contract.py -v`
Expected: 3 passed

- [ ] **Step 5: Write failing test for load_contract_index**

```python
# tests/test_contract.py (append)
import os
import tempfile
from openpyxl import Workbook
from contract import load_contract_index


def test_load_contract_index():
    """从Excel加载合同号索引映射"""
    wb = Workbook()
    ws = wb.active
    # 表头
    ws["A1"] = "乙方合同号"
    ws["H1"] = "合同名称"
    # 数据行
    ws["A2"] = "2100002601080A"
    ws["H2"] = "YW-7871Y00200172-0C1911-651400P"
    ws["A3"] = "2100002602020J"
    ws["H3"] = "YW-SZYG4502802-4-S292912F-3000P"

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        wb.save(f.name)
        tmp_path = f.name

    try:
        index = load_contract_index(tmp_path)
        assert index == {
            "2100002601080A": "7871Y00200172",
            "2100002602020J": "SZYG4502802-4",
        }
    finally:
        os.unlink(tmp_path)
```

- [ ] **Step 6: Run test to verify it fails**

Run: `uv run pytest tests/test_contract.py::test_load_contract_index -v`
Expected: FAIL with `AttributeError: module 'contract' has no attribute 'load_contract_index'`

- [ ] **Step 7: Write implementation**

```python
# contract.py (append)
from openpyxl import load_workbook


def load_contract_index(excel_path: str) -> dict[str, str]:
    """加载合同号索引表，返回 {乙方合同号: 甲方合同号} 映射。

    假设 Column A = 乙方合同号，Column H = 合同名称，从第2行开始读取。
    """
    wb = load_workbook(excel_path, read_only=True)
    ws = wb.active
    index = {}
    for row in ws.iter_rows(min_row=2, max_col=8):
        party_b_id = row[0].value  # Column A
        contract_name = row[7].value  # Column H
        if party_b_id and contract_name:
            party_a_id = extract_party_a_contract(str(contract_name))
            index[str(party_b_id)] = party_a_id
    wb.close()
    return index
```

- [ ] **Step 8: Run test to verify it passes**

Run: `uv run pytest tests/test_contract.py -v`
Expected: 4 passed

- [ ] **Step 9: Commit**

```bash
git add contract.py tests/test_contract.py
git commit -m "feat: contract number parsing and index loading"
```

---

### Task 3: Excel Utilities

**Files:**
- Create: `excel_utils.py`
- Create: `tests/test_excel_utils.py` (optional, covered by contract/verify tests)

- [ ] **Step 1: Create excel_utils.py**

```python
# excel_utils.py
from openpyxl import load_workbook, Workbook


# Excel列号映射（基于列字母）
COLUMN_MAP = {
    "party_a_id": "AF",    # 甲方合同号
    "billing_qty": "BZ",   # 本次开票数量
    "invoice_no": "CA",    # 发票号
    "invoice_date": "CB",  # 发票日期
    "tax_amount": "CC",    # 税金金额
    "total_amount": "CD",  # 含税金额
}

# 数据起始行
DATA_START_ROW = 4


def load_workbook_rw(path: str):
    """加载Excel工作簿（读写模式）"""
    return load_workbook(path)


def get_column_index(col_letter: str) -> int:
    """将列字母转换为1-based列号，如 'AF' → 32"""
    result = 0
    for c in col_letter:
        result = result * 26 + (ord(c) - ord("A") + 1)
    return result


def read_invoice_rows(ws) -> list[dict]:
    """从工作表中读取发票数据行（从DATA_START_ROW开始）"""
    rows = []
    for row in ws.iter_rows(min_row=DATA_START_ROW):
        party_a_id = row[get_column_index("AF") - 1].value
        if not party_a_id:
            continue
        rows.append({
            "party_a_id": str(party_a_id),
            "billing_qty": row[get_column_index("BZ") - 1].value,
            "invoice_no": row[get_column_index("CA") - 1].value,
            "invoice_date": row[get_column_index("CB") - 1].value,
            "tax_amount": row[get_column_index("CC") - 1].value,
            "total_amount": row[get_column_index("CD") - 1].value,
            "_row_idx": row[0].row,  # 记录行号用于回写
        })
    return rows


def write_cell(ws, row_idx: int, col_letter: str, value):
    """写入单元格"""
    col_idx = get_column_index(col_letter)
    ws.cell(row=row_idx, column=col_idx, value=value)
```

- [ ] **Step 2: Commit**

```bash
git add excel_utils.py
git commit -m "feat: excel utilities for reading/writing invoice data"
```

---

### Task 4: OCR Module

**Files:**
- Create: `ocr.py`
- Create: `tests/test_ocr.py`

- [ ] **Step 1: Write failing test for extract_text_from_pdf**

```python
# tests/test_ocr.py
import os
import tempfile
import fitz  # PyMuPDF
from ocr import extract_text_from_pdf


def test_extract_text_from_pdf_with_text_layer():
    """从有文本层的PDF中提取文本"""
    # 创建一个包含文本的PDF
    doc = fitz.open()
    page = doc.new_page()
    text = "发票号码：26317000001473818243\n开票日期：2026年05月19日\n税额：3237085.21"
    page.insert_text((72, 72), text)

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        doc.save(f.name)
        tmp_path = f.name
    doc.close()

    try:
        result = extract_text_from_pdf(tmp_path)
        assert "发票号码" in result
        assert "26317000001473818243" in result
    finally:
        os.unlink(tmp_path)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_ocr.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'ocr'`

- [ ] **Step 3: Write minimal implementation**

```python
# ocr.py
import re
import fitz


# 文本层阈值：少于这个字数认为无文本层
TEXT_LAYER_THRESHOLD = 50


def extract_text_from_pdf(pdf_path: str) -> str:
    """从PDF提取文本。优先用PyMuPDF文本层，失败则fallback到OCR。"""
    text = _extract_text_layer(pdf_path)
    if len(text) >= TEXT_LAYER_THRESHOLD:
        return text
    return ocr_pdf(pdf_path)


def _extract_text_layer(pdf_path: str) -> str:
    """用PyMuPDF提取PDF文本层"""
    doc = fitz.open(pdf_path)
    text_parts = []
    for page in doc:
        text_parts.append(page.get_text())
    doc.close()
    return "\n".join(text_parts)


def ocr_pdf(pdf_path: str) -> str:
    """将PDF转图片后用PaddleOCR识别"""
    from paddleocr import PaddleOCR

    ocr = PaddleOCR(use_angle_cls=True, lang="ch", show_log=False)
    doc = fitz.open(pdf_path)
    all_text = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        # PDF转图片（2x缩放）
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_bytes = pix.tobytes("png")

        result = ocr.ocr(img_bytes, cls=True)
        if result and result[0]:
            for line in result[0]:
                all_text.append(line[1][0])

    doc.close()
    return "\n".join(all_text)


def extract_invoice_fields(text: str) -> dict:
    """从OCR文本中提取发票关键字段"""
    fields = {}

    # 发票号码
    m = re.search(r"发票号码[:：]\s*(\d{20})", text)
    if m:
        fields["invoice_no"] = m.group(1)

    # 开票日期
    m = re.search(r"开票日期[:：]\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日", text)
    if m:
        fields["invoice_date"] = f"{m.group(1)}/{m.group(2).zfill(2)}/{m.group(3).zfill(2)}"

    # 税额
    m = re.search(r"税额\s*[:：]?\s*[\¥￥]?\s*([\d,.]+)", text)
    if m:
        fields["tax_amount"] = m.group(1).replace(",", "")

    # 价税合计
    m = re.search(r"价税合计.*?[\¥￥]\s*([\d,.]+)", text)
    if m:
        fields["total_amount"] = m.group(1).replace(",", "")

    # 乙方合同号（备注栏）
    m = re.search(r"备注[：:]?\s*.*?(\d{13}[A-Z])", text, re.DOTALL)
    if m:
        fields["party_b_id"] = m.group(1)

    return fields


def get_ocr_confidence(pdf_path: str) -> float:
    """获取OCR识别的平均置信度"""
    from paddleocr import PaddleOCR

    ocr = PaddleOCR(use_angle_cls=True, lang="ch", show_log=False)
    doc = fitz.open(pdf_path)
    confidences = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_bytes = pix.tobytes("png")
        result = ocr.ocr(img_bytes, cls=True)
        if result and result[0]:
            for line in result[0]:
                confidences.append(line[1][1])

    doc.close()
    return sum(confidences) / len(confidences) if confidences else 0.0
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_ocr.py -v`
Expected: PASS

- [ ] **Step 5: Write test for extract_invoice_fields**

```python
# tests/test_ocr.py (append)
from ocr import extract_invoice_fields


def test_extract_invoice_fields():
    """从OCR文本中提取发票字段"""
    text = """电子发票（增值税专用发票）
    发票号码：26317000001473818243
    开票日期：2026年05月19日
    税额：3237085.21
    价税合计（大写）贰仟陆佰壹拾叁万柒仟柒佰肆拾圆陆角陆分（小写）¥28137740.66
    备注：
    1100002609110C
    乙方合同号"""

    fields = extract_invoice_fields(text)
    assert fields["invoice_no"] == "26317000001473818243"
    assert fields["invoice_date"] == "2026/05/19"
    assert fields["tax_amount"] == "3237085.21"
    assert fields["total_amount"] == "28137740.66"
    assert fields["party_b_id"] == "1100002609110C"
```

- [ ] **Step 6: Run test to verify it passes**

Run: `uv run pytest tests/test_ocr.py::test_extract_invoice_fields -v`
Expected: PASS

- [ ] **Step 7: Commit**

```bash
git add ocr.py tests/test_ocr.py
git commit -m "feat: OCR module with PyMuPDF text extraction and PaddleOCR fallback"
```

---

### Task 5: File Rename Logic

**Files:**
- Create: `rename.py`
- Create: `tests/test_rename.py`

- [ ] **Step 1: Write failing test for classify_pdf**

```python
# tests/test_rename.py
from rename import classify_pdf


def test_classify_type1():
    """Type 1: 乙方合同号在文件名中"""
    file_type, party_b_id = classify_pdf("2100002601080A_引望智能技术有限公司_20260519.pdf")
    assert file_type == "type1"
    assert party_b_id == "2100002601080A"


def test_classify_type2():
    """Type 2: dzfp开头的发票文件"""
    file_type, party_b_id = classify_pdf("dzfp_26317000001473818238_引望智能技术有限公司_20260519.pdf")
    assert file_type == "type2"
    assert party_b_id is None


def test_classify_unknown():
    """未知类型"""
    file_type, party_b_id = classify_pdf("unknown_file.pdf")
    assert file_type == "unknown"
    assert party_b_id is None
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_rename.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'rename'`

- [ ] **Step 3: Write minimal implementation**

```python
# rename.py
import os
import re
from dataclasses import dataclass
from ocr import extract_text_from_pdf, extract_invoice_fields, get_ocr_confidence
from contract import load_contract_index


@dataclass
class RenameResult:
    original_name: str
    new_name: str
    party_a_id: str
    party_b_id: str
    status: str  # "success" | "skipped" | "error"
    message: str


def classify_pdf(filename: str) -> tuple[str, str | None]:
    """判断PDF类型。

    Returns:
        (类型, 乙方合同号或None)
        type1: 2100002601080A_xxx.pdf → 乙方合同号在文件名中
        type2: dzfp_26317000001473818238_xxx.pdf → 需要OCR识别
        unknown: 无法识别的文件
    """
    basename = os.path.basename(filename)
    # Type 1: 以数字+大写字母开头（如 2100002601080A_）
    m = re.match(r"^(\d+[A-Z])_", basename)
    if m:
        return "type1", m.group(1)
    # Type 2: 以dzfp_开头
    if basename.startswith("dzfp_"):
        return "type2", None
    return "unknown", None


def extract_invoice_number(filename: str) -> str | None:
    """从Type2文件名中提取发票号码"""
    m = re.match(r"dzfp_(\d{20})_", filename)
    return m.group(1) if m else None


def rename_files(pdf_dir: str, contract_index: dict[str, str]) -> list[RenameResult]:
    """重命名目录中的PDF文件"""
    results = []
    for filename in sorted(os.listdir(pdf_dir)):
        if not filename.lower().endswith(".pdf"):
            continue

        file_type, party_b_id = classify_pdf(filename)
        old_path = os.path.join(pdf_dir, filename)

        if file_type == "unknown":
            results.append(RenameResult(
                original_name=filename, new_name="", party_a_id="",
                party_b_id="", status="skipped", message="无法识别的文件类型"
            ))
            continue

        try:
            if file_type == "type1":
                party_a_id = contract_index.get(party_b_id, "")
                if not party_a_id:
                    results.append(RenameResult(
                        original_name=filename, new_name="", party_a_id="",
                        party_b_id=party_b_id, status="error",
                        message=f"索引中未找到乙方合同号 {party_b_id}"
                    ))
                    continue

                # 保留乙方合同号之后的部分
                remainder = filename[len(party_b_id):]
                new_name = f"{party_a_id}-{party_b_id}{remainder}"

            else:  # type2
                text = extract_text_from_pdf(old_path)
                fields = extract_invoice_fields(text)
                party_b_id = fields.get("party_b_id", "")
                if not party_b_id:
                    results.append(RenameResult(
                        original_name=filename, new_name="", party_a_id="",
                        party_b_id="", status="error",
                        message="OCR未能识别乙方合同号"
                    ))
                    continue

                party_a_id = contract_index.get(party_b_id, "")
                if not party_a_id:
                    results.append(RenameResult(
                        original_name=filename, new_name="", party_a_id="",
                        party_b_id=party_b_id, status="error",
                        message=f"索引中未找到乙方合同号 {party_b_id}"
                    ))
                    continue

                invoice_no = extract_invoice_number(filename)
                # dzfp_ 后是发票号码，之后的部分
                remainder_start = filename.index("_", 5) + 1
                remainder = filename[remainder_start:]
                new_name = f"{party_a_id}-{party_b_id}_{invoice_no}_{remainder}"

            new_path = os.path.join(pdf_dir, new_name)
            os.rename(old_path, new_path)
            results.append(RenameResult(
                original_name=filename, new_name=new_name,
                party_a_id=party_a_id, party_b_id=party_b_id,
                status="success", message="重命名成功"
            ))

        except Exception as e:
            results.append(RenameResult(
                original_name=filename, new_name="", party_a_id="",
                party_b_id=party_b_id or "", status="error", message=str(e)
            ))

    return results
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_rename.py -v`
Expected: 3 passed

- [ ] **Step 5: Commit**

```bash
git add rename.py tests/test_rename.py
git commit -m "feat: PDF file classification and rename logic"
```

---

### Task 6: Verify Logic

**Files:**
- Create: `verify.py`
- Create: `tests/test_verify.py`

- [ ] **Step 1: Write failing test for compare_fields**

```python
# tests/test_verify.py
from verify import compare_fields


def test_compare_fields_all_match():
    """所有字段匹配"""
    excel_row = {
        "invoice_no": "26317000001473818243",
        "invoice_date": "2026/05/19",
        "tax_amount": "3237085.21",
        "total_amount": "28137740.66",
    }
    ocr_fields = {
        "invoice_no": "26317000001473818243",
        "invoice_date": "2026/05/19",
        "tax_amount": "3237085.21",
        "total_amount": "28137740.66",
    }
    diffs = compare_fields(excel_row, ocr_fields, confidence=0.95)
    assert len(diffs) == 0


def test_compare_fields_has_diff():
    """存在差异"""
    excel_row = {
        "invoice_no": "26317000001473818243",
        "invoice_date": "2026/05/19",
        "tax_amount": "3237085.21",
        "total_amount": "28137740.66",
    }
    ocr_fields = {
        "invoice_no": "26317000001473818243",
        "invoice_date": "2026/05/20",  # 不一致
        "tax_amount": "3237085.00",    # 不一致
        "total_amount": "28137740.66",
    }
    diffs = compare_fields(excel_row, ocr_fields, confidence=0.95)
    assert len(diffs) == 2
    assert diffs[0].field_name == "invoice_date"
    assert diffs[1].field_name == "tax_amount"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_verify.py -v`
Expected: FAIL with `ModuleNotFoundError: No module named 'verify'`

- [ ] **Step 3: Write minimal implementation**

```python
# verify.py
import os
from dataclasses import dataclass, field
from openpyxl import load_workbook
from ocr import extract_text_from_pdf, extract_invoice_fields, get_ocr_confidence
from excel_utils import COLUMN_MAP, DATA_START_ROW, get_column_index, read_invoice_rows, write_cell


CONFIDENCE_THRESHOLD = 0.8


@dataclass
class FieldDiff:
    field_name: str
    excel_value: str
    ocr_value: str
    confidence: float
    fixed: bool


@dataclass
class VerifyResult:
    pdf_file: str
    party_a_id: str
    diffs: list[FieldDiff] = field(default_factory=list)
    fixed: bool = False
    needs_manual: bool = False


def compare_fields(excel_row: dict, ocr_fields: dict, confidence: float) -> list[FieldDiff]:
    """比对Excel行与OCR识别结果，返回差异列表"""
    field_map = {
        "invoice_no": "发票号",
        "invoice_date": "发票日期",
        "tax_amount": "税金金额",
        "total_amount": "含税金额",
    }
    diffs = []
    for key, label in field_map.items():
        excel_val = str(excel_row.get(key, "")).strip()
        ocr_val = str(ocr_fields.get(key, "")).strip()
        if excel_val and ocr_val and excel_val != ocr_val:
            diffs.append(FieldDiff(
                field_name=label,
                excel_value=excel_val,
                ocr_value=ocr_val,
                confidence=confidence,
                fixed=False,
            ))
    return diffs


def resolve_party_a_id_from_filename(filename: str) -> str | None:
    """从已重命名的文件名中提取甲方合同号"""
    import re
    # 匹配 {甲方合同号}-{乙方合同号}_... 格式
    m = re.match(r"^(.+?)-(\d+[A-Z])_", filename)
    if m:
        return m.group(1)
    return None


def verify_invoices(pdf_dir: str, excel_path: str) -> list[VerifyResult]:
    """校验目录中所有PDF发票与Excel数据"""
    wb = load_workbook(excel_path)
    ws = wb.active
    excel_rows = read_invoice_rows(ws)

    # 按甲方合同号索引Excel行
    excel_by_party_a = {r["party_a_id"]: r for r in excel_rows}

    results = []
    for filename in sorted(os.listdir(pdf_dir)):
        if not filename.lower().endswith(".pdf"):
            continue

        party_a_id = resolve_party_a_id_from_filename(filename)
        if not party_a_id:
            results.append(VerifyResult(
                pdf_file=filename, party_a_id="",
                needs_manual=True,
            ))
            continue

        excel_row = excel_by_party_a.get(party_a_id)
        if not excel_row:
            results.append(VerifyResult(
                pdf_file=filename, party_a_id=party_a_id,
                needs_manual=True,
            ))
            continue

        pdf_path = os.path.join(pdf_dir, filename)
        text = extract_text_from_pdf(pdf_path)
        ocr_fields = extract_invoice_fields(text)
        confidence = get_ocr_confidence(pdf_path)

        diffs = compare_fields(excel_row, ocr_fields, confidence)

        # 自动修复高置信度的差异
        fixed = False
        needs_manual = False
        for diff in diffs:
            if diff.confidence > CONFIDENCE_THRESHOLD:
                # 找到对应的Excel列并修复
                field_key = {
                    "发票号": "invoice_no",
                    "发票日期": "invoice_date",
                    "税金金额": "tax_amount",
                    "含税金额": "total_amount",
                }.get(diff.field_name)
                if field_key:
                    col_letter = COLUMN_MAP[field_key]
                    write_cell(ws, excel_row["_row_idx"], col_letter, diff.ocr_value)
                    diff.fixed = True
                    fixed = True
            else:
                needs_manual = True

        results.append(VerifyResult(
            pdf_file=filename,
            party_a_id=party_a_id,
            diffs=diffs,
            fixed=fixed,
            needs_manual=needs_manual,
        ))

    # 保存修复后的Excel
    wb.save(excel_path)
    wb.close()
    return results


def format_report(results: list[VerifyResult]) -> str:
    """格式化校验报告"""
    lines = ["=== 发票校验报告 ==="]
    for r in results:
        if not r.diffs:
            lines.append(f"✓ {r.party_a_id}: 所有字段正确")
            continue
        for diff in r.diffs:
            status = "已自动修复" if diff.fixed else "需手动核查"
            lines.append(
                f"✗ {r.party_a_id}: {diff.field_name}不一致 "
                f"(Excel: {diff.excel_value}, OCR: {diff.ocr_value}) [{status}]"
            )
    return "\n".join(lines)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_verify.py -v`
Expected: 2 passed

- [ ] **Step 5: Commit**

```bash
git add verify.py tests/test_verify.py
git commit -m "feat: invoice verification with auto-fix and reporting"
```

---

### Task 7: CLI Entry Point

**Files:**
- Modify: `main.py`

- [ ] **Step 1: Write CLI with argparse**

```python
# main.py
import argparse
import sys


def cmd_rename(args):
    from contract import load_contract_index
    from rename import rename_files

    print(f"加载合同号索引: {args.excel}")
    contract_index = load_contract_index(args.excel)
    print(f"找到 {len(contract_index)} 条合同记录")

    print(f"扫描目录: {args.dir}")
    results = rename_files(args.dir, contract_index)

    success = sum(1 for r in results if r.status == "success")
    skipped = sum(1 for r in results if r.status == "skipped")
    errors = sum(1 for r in results if r.status == "error")

    for r in results:
        if r.status == "success":
            print(f"  ✓ {r.original_name} → {r.new_name}")
        elif r.status == "error":
            print(f"  ✗ {r.original_name}: {r.message}")
        else:
            print(f"  - {r.original_name}: {r.message}")

    print(f"\n完成: {success} 重命名, {skipped} 跳过, {errors} 错误")


def cmd_verify(args):
    from verify import verify_invoices, format_report

    print(f"校验发票: {args.dir}")
    print(f"对照Excel: {args.excel}")
    results = verify_invoices(args.dir, args.excel)
    print(format_report(results))


def cmd_generate_test_data(args):
    from tests.conftest import generate_all_test_data

    print(f"生成测试数据到: {args.output}")
    generate_all_test_data(args.output)
    print("完成")


def main():
    parser = argparse.ArgumentParser(description="发票自动重命名和校验工具")
    subparsers = parser.add_subparsers(dest="command", required=True)

    # rename
    p_rename = subparsers.add_parser("rename", help="根据合同号索引重命名PDF文件")
    p_rename.add_argument("--dir", required=True, help="PDF文件所在目录")
    p_rename.add_argument("--excel", required=True, help="合同号索引Excel文件路径")

    # verify
    p_verify = subparsers.add_parser("verify", help="校验发票信息与Excel是否一致")
    p_verify.add_argument("--dir", required=True, help="PDF文件所在目录")
    p_verify.add_argument("--excel", required=True, help="发票验证Excel文件路径")

    # generate-test-data
    p_gen = subparsers.add_parser("generate-test-data", help="生成测试用的mock数据")
    p_gen.add_argument("--output", required=True, help="输出目录")

    args = parser.parse_args()
    if args.command == "rename":
        cmd_rename(args)
    elif args.command == "verify":
        cmd_verify(args)
    elif args.command == "generate-test-data":
        cmd_generate_test_data(args)


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Verify CLI runs without error**

Run: `uv run main.py --help`
Expected: Shows help with rename/verify/generate-test-data subcommands

- [ ] **Step 3: Commit**

```bash
git add main.py
git commit -m "feat: CLI entry point with argparse"
```

---

### Task 8: Test Data Generation

**Files:**
- Modify: `tests/conftest.py`

- [ ] **Step 1: Write test data generator**

```python
# tests/conftest.py
import os
import fitz
from openpyxl import Workbook


# 基于参考图中的真实数据
MOCK_CONTRACTS = [
    ("2100002601080A", "YW-7871Y00200172-0C1911-651400P"),
    ("2100002602020J", "YW-7871Y00200315L-092176-200P-0109"),
    ("2100002602040D", "YW-7871Y00200344H-1105-8040P-0106"),
    ("2100002602040F", "YW-7871Y00200343H-1105-246P-0106"),
    ("2100002602100D", "YW-SZYG4502802-4-S292912F-3000P"),
]

MOCK_INVOICES = [
    {
        "party_a_id": "7871Y00200172",
        "party_b_id": "2100002601080A",
        "billing_qty": 1265,
        "invoice_no": "26317000001473818247",
        "invoice_date": "2026/05/19",
        "tax_amount": 256308.48,
        "total_amount": 2227912.18,
    },
    {
        "party_a_id": "7871Y00200315L",
        "party_b_id": "2100002602020J",
        "billing_qty": 925,
        "invoice_no": "26317000001473818224",
        "invoice_date": "2026/05/19",
        "tax_amount": 187419.25,
        "total_amount": 1629105.75,
    },
    {
        "party_a_id": "7871Y00200344H",
        "party_b_id": "2100002602040D",
        "billing_qty": 1127,
        "invoice_no": "26317000001473818221",
        "invoice_date": "2026/05/19",
        "tax_amount": 228347.56,
        "total_amount": 1984867.22,
    },
]


def _create_contract_excel(path: str):
    """创建合同号索引Excel"""
    wb = Workbook()
    ws = wb.active
    ws.title = "合同汇总"
    ws["A1"] = "乙方合同号"
    ws["H1"] = "合同名称"
    ws["A2"] = "summary information"

    for i, (party_b, full_name) in enumerate(MOCK_CONTRACTS, start=3):
        ws[f"A{i}"] = party_b
        ws[f"H{i}"] = full_name

    wb.save(path)
    wb.close()


def _create_invoice_excel(path: str):
    """创建发票验证Excel"""
    wb = Workbook()
    ws = wb.active
    # 表头 (Row 1-3)
    ws.cell(row=1, column=32, value="Fields in this area ca")
    ws.cell(row=2, column=32, value="PO No.")
    ws.cell(row=2, column=70, value="Billing Qty*")
    ws.cell(row=2, column=79, value="Invoice No*")
    ws.cell(row=2, column=80, value="Invoice Date*")
    ws.cell(row=2, column=81, value="Tax Amount*")
    ws.cell(row=2, column=82, value="Invoice Amount (Incl. Tax)*")
    ws.cell(row=3, column=32, value="PO号")
    ws.cell(row=3, column=70, value="本次开票数量")
    ws.cell(row=3, column=79, value="发票号")
    ws.cell(row=3, column=80, value="发票日期")
    ws.cell(row=3, column=81, value="税金金额")
    ws.cell(row=3, column=82, value="含税金额")

    # 数据行 (Row 4+)
    for i, inv in enumerate(MOCK_INVOICES, start=4):
        ws.cell(row=i, column=32, value=inv["party_a_id"])  # AF
        ws.cell(row=i, column=70, value=inv["billing_qty"])  # BZ
        ws.cell(row=i, column=79, value=inv["invoice_no"])   # CA
        ws.cell(row=i, column=80, value=inv["invoice_date"]) # CB
        ws.cell(row=i, column=81, value=inv["tax_amount"])   # CC
        ws.cell(row=i, column=82, value=inv["total_amount"]) # CD

    wb.save(path)
    wb.close()


def _create_mock_pdf(path: str, party_b_id: str, invoice_no: str,
                     invoice_date: str, tax_amount: str, total_amount: str,
                     include_text_layer: bool = True):
    """创建包含发票内容的mock PDF"""
    doc = fitz.open()
    page = doc.new_page()

    # 构造发票文本
    text = f"""电子发票（增值税专用发票）
    发票号码：{invoice_no}
    开票日期：{invoice_date.replace('/', '年', 1).replace('/', '月')}日
    数量：10000
    金额：{total_amount}
    税率/征收率：13%
    税额：{tax_amount}
    价税合计（大写）贰仟陆佰壹拾叁万柒仟柒佰肆拾圆陆角陆分（小写）¥{total_amount}
    备注：
    {party_b_id}
    乙方合同号"""

    if include_text_layer:
        # 有文本层的PDF
        page.insert_text((50, 50), text, fontsize=10)
    else:
        # 无文本层的PDF（只插入图片）
        # 创建一个简单的图片
        from PIL import Image
        import io
        img = Image.new("RGB", (800, 600), "white")
        # 简单画一些文字（用默认字体）
        img_bytes = io.BytesIO()
        img.save(img_bytes, format="PNG")
        img_bytes.seek(0)
        rect = fitz.Rect(0, 0, 800, 600)
        page.insert_image(rect, stream=img_bytes.getvalue())

    doc.save(path)
    doc.close()


def generate_all_test_data(output_dir: str):
    """生成所有测试数据"""
    os.makedirs(output_dir, exist_ok=True)
    pdf_dir = os.path.join(output_dir, "pdfs")
    os.makedirs(pdf_dir, exist_ok=True)

    # 生成合同号索引Excel
    contract_excel = os.path.join(output_dir, "合同号索引表.xlsx")
    _create_contract_excel(contract_excel)

    # 生成发票验证Excel
    invoice_excel = os.path.join(output_dir, "发票验证表.xlsx")
    _create_invoice_excel(invoice_excel)

    # 生成Type 1 PDF（有文本层）
    for inv in MOCK_INVOICES:
        filename = f"{inv['party_b_id']}_引望智能技术有限公司_20260519.pdf"
        _create_mock_pdf(
            os.path.join(pdf_dir, filename),
            party_b_id=inv["party_b_id"],
            invoice_no=inv["invoice_no"],
            invoice_date=inv["invoice_date"],
            tax_amount=str(inv["tax_amount"]),
            total_amount=str(inv["total_amount"]),
            include_text_layer=True,
        )

    # 生成Type 2 PDF（有文本层）
    for inv in MOCK_INVOICES:
        filename = f"dzfp_{inv['invoice_no']}_引望智能技术有限公司_20260519.pdf"
        _create_mock_pdf(
            os.path.join(pdf_dir, filename),
            party_b_id=inv["party_b_id"],
            invoice_no=inv["invoice_no"],
            invoice_date=inv["invoice_date"],
            tax_amount=str(inv["tax_amount"]),
            total_amount=str(inv["total_amount"]),
            include_text_layer=True,
        )

    return {
        "contract_excel": contract_excel,
        "invoice_excel": invoice_excel,
        "pdf_dir": pdf_dir,
    }
```

- [ ] **Step 2: Verify test data generation works**

Run: `uv run main.py generate-test-data --output ./test_data`
Expected: Creates test_data/ with Excel files and PDFs

- [ ] **Step 3: Verify generated files exist**

Run: `ls -la test_data/ && ls -la test_data/pdfs/`
Expected: Excel files and PDF files listed

- [ ] **Step 4: Commit**

```bash
git add tests/conftest.py
git commit -m "feat: test data generation with mock invoices"
```

---

### Task 9: Integration Test

**Files:**
- Create: `tests/test_integration.py`

- [ ] **Step 1: Write integration test**

```python
# tests/test_integration.py
import os
import shutil
import tempfile
from tests.conftest import generate_all_test_data
from contract import load_contract_index
from rename import rename_files
from verify import verify_invoices, format_report


def test_full_workflow():
    """端到端测试：生成数据 → 重命名 → 校验"""
    tmp_dir = tempfile.mkdtemp()
    try:
        # 生成测试数据
        data = generate_all_test_data(tmp_dir)

        # 任务1：重命名
        contract_index = load_contract_index(data["contract_excel"])
        rename_results = rename_files(data["pdf_dir"], contract_index)

        success_count = sum(1 for r in rename_results if r.status == "success")
        assert success_count > 0, f"没有成功重命名的文件: {rename_results}"

        # 验证重命名后的文件名格式
        for r in rename_results:
            if r.status == "success":
                assert "-" in r.new_name, f"文件名缺少甲方合同号: {r.new_name}"

        # 任务2：校验
        verify_results = verify_invoices(data["pdf_dir"], data["invoice_excel"])
        report = format_report(verify_results)
        assert "发票校验报告" in report

    finally:
        shutil.rmtree(tmp_dir)
```

- [ ] **Step 2: Run integration test**

Run: `uv run pytest tests/test_integration.py -v`
Expected: PASS

- [ ] **Step 3: Run all tests**

Run: `uv run pytest tests/ -v`
Expected: All tests pass

- [ ] **Step 4: Commit**

```bash
git add tests/test_integration.py
git commit -m "feat: end-to-end integration test"
```

---

### Task 10: Final Polish

**Files:**
- Modify: `pyproject.toml` (add pytest config)

- [ ] **Step 1: Add pytest config to pyproject.toml**

```toml
# 在 pyproject.toml 中添加
[tool.pytest.ini_options]
testpaths = ["tests"]
asyncio_mode = "auto"
```

- [ ] **Step 2: Run full test suite**

Run: `uv run pytest tests/ -v --tb=short`
Expected: All tests pass

- [ ] **Step 3: Commit**

```bash
git add pyproject.toml
git commit -m "chore: add pytest configuration"
```
