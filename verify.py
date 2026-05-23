import logging
import os
import re
from dataclasses import dataclass, field

from openpyxl import load_workbook

from excel_utils import COLUMN_MAP, get_column_index, read_invoice_rows, write_cell
from ocr import extract_invoice_fields, extract_text_from_pdf, get_ocr_confidence

logger = logging.getLogger(__name__)

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
            logger.debug("字段 '%s' 不一致: Excel='%s', OCR='%s'", label, excel_val, ocr_val)
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
    # 匹配 {甲方合同号}-{乙方合同号}_... 格式
    m = re.match(r"^(.+?)-(\d+[A-Z])_", filename)
    if m:
        return m.group(1)
    return None


def verify_invoices(pdf_dir: str, excel_path: str) -> list[VerifyResult]:
    """校验目录中所有PDF发票与Excel数据"""
    logger.info("打开验证Excel: %s", excel_path)
    wb = load_workbook(excel_path)
    ws = wb.active
    excel_rows = read_invoice_rows(ws)
    logger.info("Excel中读取到 %d 行发票数据", len(excel_rows))

    # 按甲方合同号索引Excel行
    excel_by_party_a = {r["party_a_id"]: r for r in excel_rows}

    pdf_files = [f for f in sorted(os.listdir(pdf_dir)) if f.lower().endswith(".pdf")]
    logger.info("发现 %d 个PDF文件", len(pdf_files))

    results = []
    for filename in pdf_files:
        party_a_id = resolve_party_a_id_from_filename(filename)
        if not party_a_id:
            logger.warning("文件名格式无法识别甲方合同号: %s", filename)
            results.append(VerifyResult(
                pdf_file=filename, party_a_id="",
                needs_manual=True,
            ))
            continue

        excel_row = excel_by_party_a.get(party_a_id)
        if not excel_row:
            logger.warning("Excel中未找到甲方合同号: %s (文件: %s)", party_a_id, filename)
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
                    logger.info("自动修复: %s %s, '%s' → '%s' (置信度=%.3f)",
                                party_a_id, diff.field_name, diff.excel_value, diff.ocr_value, confidence)
            else:
                needs_manual = True
                logger.warning("需手动核查: %s %s, Excel='%s' OCR='%s' (置信度=%.3f, 低于阈值%.1f)",
                               party_a_id, diff.field_name, diff.excel_value, diff.ocr_value,
                               confidence, CONFIDENCE_THRESHOLD)

        if not diffs:
            logger.debug("✓ %s: 所有字段正确", party_a_id)

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
    logger.info("验证完成，Excel已保存")
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
