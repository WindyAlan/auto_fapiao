import logging
import os
import re
import shutil
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation

from openpyxl import load_workbook

from excel_utils import COLUMN_MAP, get_column_index, read_invoice_rows
from ocr import extract_invoice_fields, extract_pdf_content

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
    filled: list[str] = field(default_factory=list)  # OCR填充的字段描述
    invoice_no: str = ""  # OCR识别的发票号（用于重命名并复制到_filled文件夹）
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
    m = re.match(r"^(.+?)-(\d+[A-Z]+)_", filename)
    if m:
        return m.group(1)
    return None


def _as_decimal(value) -> Decimal:
    """将Excel/OCR数量转换为十进制数，避免浮点数舍入。"""
    return Decimal(str(value).replace(",", "").strip())


def reconcile_billing_quantities(
    excel_rows: list[dict], ocr_quantities: list[Decimal],
) -> tuple[dict[int, Decimal], str | None]:
    """计算同一 PO 的本次开票数量下调方案，绝不增加任何Excel行。"""
    if not ocr_quantities:
        return {}, None

    try:
        original = [_as_decimal(row.get("billing_qty", "")) for row in excel_rows]
    except (InvalidOperation, ValueError):
        return {}, "Excel中的本次开票数量包含无法识别的数值"

    if any(value < 0 for value in original) or any(value < 0 for value in ocr_quantities):
        return {}, "本次开票数量不能为负数"

    excel_total = sum(original)
    ocr_total = sum(ocr_quantities)
    if ocr_total > excel_total:
        return {}, f"OCR开票数量合计 {ocr_total} 大于Excel原始合计 {excel_total}，不能增加数量"

    # 数量条目数一致时，按明细顺序逐项匹配，避免把数量写到错误行。
    if len(ocr_quantities) == len(excel_rows):
        if any(new > old for new, old in zip(ocr_quantities, original)):
            return {}, "OCR明细数量会增加Excel中的某一行，本次开票数量只能减少，不能增加"
        return {
            row["_row_idx"]: new
            for row, old, new in zip(excel_rows, original, ocr_quantities)
            if new != old
        }, None

    # 只有汇总数量时，从最大的行开始下调，保留较小的明细行。
    adjusted = original[:]
    remaining_reduction = excel_total - ocr_total
    for index in sorted(range(len(adjusted)), key=lambda i: adjusted[i], reverse=True):
        reduction = min(adjusted[index], remaining_reduction)
        adjusted[index] -= reduction
        remaining_reduction -= reduction
        if remaining_reduction == 0:
            break
    return {
        row["_row_idx"]: new
        for row, old, new in zip(excel_rows, original, adjusted)
        if new != old
    }, None


def verify_invoices(pdf_dir: str, excel_path: str) -> tuple[list[VerifyResult], str]:
    """校验目录中所有PDF发票与Excel数据，并将OCR结果填入新的Excel文件。

    Returns:
        (结果列表, 输出Excel路径)
    """
    # 生成输出文件名
    base, ext = os.path.splitext(excel_path)
    output_excel = f"{base}_Verified{ext}"

    logger.info("打开验证Excel: %s", excel_path)
    wb = load_workbook(excel_path)
    ws = wb.active
    excel_rows = read_invoice_rows(ws)
    logger.info("Excel中读取到 %d 行发票数据", len(excel_rows))

    # 同一 PO 可以有多行，必须保留全部行（旧逻辑的字典推导会只留下最后一行）。
    excel_by_party_a: dict[str, list[dict]] = {}
    for row in excel_rows:
        excel_by_party_a.setdefault(row["party_a_id"], []).append(row)

    pdf_files = [f for f in sorted(os.listdir(pdf_dir)) if f.lower().endswith(".pdf")]
    logger.info("发现 %d 个PDF文件", len(pdf_files))

    results = []
    filled_count = 0
    # 收集每个 PO 的所有OCR数量；一个PDF可包含多条数量明细，或同一PO可有多个PDF。
    ocr_quantities_by_party: dict[str, list[Decimal]] = {}
    results_by_party: dict[str, list[VerifyResult]] = {}
    for filename in pdf_files:
        # 先识别发票号，确保未完成重命名或Excel匹配的PDF也能按发票号导出。
        pdf_path = os.path.join(pdf_dir, filename)
        text, confidence = extract_pdf_content(pdf_path)
        ocr_fields = extract_invoice_fields(text)

        party_a_id = resolve_party_a_id_from_filename(filename)
        if not party_a_id:
            logger.warning("文件名格式无法识别甲方合同号: %s", filename)
            results.append(VerifyResult(
                pdf_file=filename, party_a_id="",
                invoice_no=ocr_fields.get("invoice_no", ""),
                needs_manual=True,
            ))
            continue

        po_rows = excel_by_party_a.get(party_a_id)
        if not po_rows:
            logger.warning("Excel中未找到甲方合同号: %s (文件: %s)", party_a_id, filename)
            results.append(VerifyResult(
                pdf_file=filename, party_a_id=party_a_id,
                invoice_no=ocr_fields.get("invoice_no", ""),
                needs_manual=True,
            ))
            continue

        # 除数量外的既有字段仍以该 PO 的首行作为校验基准；回填则作用于该 PO 的所有行。
        diffs = compare_fields(po_rows[0], ocr_fields, confidence)

        # 将OCR识别的发票号和发票日期填入Excel（如果原表为空）
        filled = []
        for excel_row in po_rows:
            row_idx = excel_row.get("_row_idx")
            if not row_idx:
                continue
            for field_key, col_letter, label in [
                ("invoice_no", "CA", "发票号"),
                ("invoice_date", "CB", "发票日期"),
            ]:
                ocr_val = ocr_fields.get(field_key, "")
                excel_val = str(excel_row.get(field_key, "")).strip()
                if ocr_val and not excel_val:
                    col_idx = get_column_index(col_letter)
                    ws.cell(row=row_idx, column=col_idx, value=ocr_val)
                    # 更新内存数据，避免同一 PO 有多个PDF时重复视为“空”。
                    excel_row[field_key] = str(ocr_val)
                    filled_count += 1
                    filled.append(f"{label}={ocr_val}")
                    logger.info("填充: %s %s = %s", party_a_id, field_key, ocr_val)

        # 标记需要手动核查的差异
        needs_manual = bool(diffs)
        for diff in diffs:
            if diff.confidence <= CONFIDENCE_THRESHOLD:
                logger.warning("需手动核查: %s %s, Excel='%s' OCR='%s' (置信度=%.3f)",
                               party_a_id, diff.field_name, diff.excel_value, diff.ocr_value,
                               confidence)

        if not diffs:
            logger.debug("✓ %s: 所有字段正确", party_a_id)

        result = VerifyResult(
            pdf_file=filename,
            party_a_id=party_a_id,
            diffs=diffs,
            filled=filled,
            invoice_no=ocr_fields.get("invoice_no", ""),
            needs_manual=needs_manual,
        )
        results.append(result)
        results_by_party.setdefault(party_a_id, []).append(result)
        ocr_quantities_by_party.setdefault(party_a_id, []).extend(
            ocr_fields.get("billing_qty_values", [])
        )

    # PO 级别校验并修正数量。所有写入均来自 reconcile_billing_quantities，
    # 它在任何单行会增加时拒绝整组更新。
    billing_qty_col = get_column_index("BZ")
    for party_a_id, ocr_quantities in ocr_quantities_by_party.items():
        po_rows = excel_by_party_a[party_a_id]
        updates, error = reconcile_billing_quantities(po_rows, ocr_quantities)
        try:
            excel_total = sum(_as_decimal(row["billing_qty"]) for row in po_rows)
        except (InvalidOperation, ValueError):
            # reconcile_billing_quantities 已拒绝此组；这里仅为生成可读的核查信息。
            excel_total = "无法解析"
        ocr_total = sum(ocr_quantities)
        po_results = results_by_party[party_a_id]

        if error:
            logger.warning("PO %s 数量未修改: %s", party_a_id, error)
            po_results[0].diffs.append(FieldDiff(
                field_name="本次开票数量",
                excel_value=str(excel_total),
                ocr_value=str(ocr_total),
                confidence=1.0,
                fixed=False,
            ))
            po_results[0].needs_manual = True
            continue

        if updates:
            for row_idx, new_value in updates.items():
                # 写整数时保持Excel中为数值而不是字符串；小数同样保持精度。
                value = int(new_value) if new_value == new_value.to_integral_value() else float(new_value)
                ws.cell(row=row_idx, column=billing_qty_col, value=value)
                for row in po_rows:
                    if row["_row_idx"] == row_idx:
                        row["billing_qty"] = str(new_value)
                        break
            po_results[0].diffs.append(FieldDiff(
                field_name="本次开票数量",
                excel_value=str(excel_total),
                ocr_value=str(ocr_total),
                confidence=1.0,
                fixed=True,
            ))
            po_results[0].filled.append(f"本次开票数量已下调至合计={ocr_total}")
            logger.info("PO %s 本次开票数量已由 %s 下调至 %s", party_a_id, excel_total, ocr_total)

    # 保存为新Excel文件
    # 如果原文件是 .xlsm，同时保存 .xlsm 和 .xlsx 两份
    base, ext = os.path.splitext(excel_path)
    if ext.lower() == ".xlsm":
        xlsm_path = f"{base}_Verified.xlsm"
        xlsx_path = f"{base}_Verified.xlsx"
        wb.save(xlsm_path)
        wb.close()
        # 重新加载 .xlsm 再另存为干净的 .xlsx
        wb2 = load_workbook(xlsm_path)
        wb2.save(xlsx_path)
        wb2.close()
        output_xlsx = xlsx_path
        logger.info("验证完成，填充了 %d 个字段", filled_count)
        logger.info("  .xlsm 版本: %s", xlsm_path)
        logger.info("  .xlsx 版本: %s", xlsx_path)
    else:
        wb.save(output_excel)
        wb.close()
        output_xlsx = output_excel
        logger.info("验证完成，填充了 %d 个字段，已保存: %s", filled_count, output_excel)

    # 将所有识别出发票号的PDF复制到_filled文件夹，以发票号重命名。
    # 是否在本次验证中回填Excel字段不影响PDF导出。
    invoice_pdfs = [r for r in results if r.invoice_no]
    if invoice_pdfs:
        parent_dir = os.path.dirname(pdf_dir)
        dir_basename = os.path.basename(pdf_dir)
        # 将 _Renamed 替换为 _filled，如果没有 _Renamed 后缀则加 _filled
        if dir_basename.endswith("_Renamed"):
            filled_dir_name = dir_basename[:-len("_Renamed")] + "_filled"
        else:
            filled_dir_name = dir_basename + "_filled"
        filled_dir = os.path.join(parent_dir, filled_dir_name)
        os.makedirs(filled_dir, exist_ok=True)

        for r in invoice_pdfs:
            src = os.path.join(pdf_dir, r.pdf_file)
            # 清理发票号中的非法文件名字符
            safe_name = re.sub(r'[\\/:*?"<>|]', '_', r.invoice_no)
            dst = os.path.join(filled_dir, f"{safe_name}.pdf")
            shutil.copy2(src, dst)
            logger.info("重命名发票文件: %s → %s", r.pdf_file, dst)

        logger.info("已将 %d 个发票文件复制并重命名到: %s", len(invoice_pdfs), filled_dir)

    return results, output_xlsx
