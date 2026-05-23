import os
import re
from dataclasses import dataclass

from contract import load_contract_index
from ocr import extract_invoice_fields, extract_text_from_pdf


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
