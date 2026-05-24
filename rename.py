import logging
import os
import re
import shutil
from dataclasses import dataclass

from contract import load_contract_index
from ocr import extract_invoice_fields, extract_text_from_pdf

logger = logging.getLogger(__name__)


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
        logger.debug("文件 '%s' → Type1, 乙方合同号=%s", basename, m.group(1))
        return "type1", m.group(1)
    # Type 2: 以dzfp_开头
    if basename.startswith("dzfp_"):
        logger.debug("文件 '%s' → Type2, 需要OCR", basename)
        return "type2", None
    logger.debug("文件 '%s' → unknown, 跳过", basename)
    return "unknown", None


def extract_invoice_number(filename: str) -> str | None:
    """从Type2文件名中提取发票号码"""
    m = re.match(r"dzfp_(\d{20})_", filename)
    return m.group(1) if m else None


def rename_files(pdf_dir: str, contract_index: dict[str, str]) -> tuple[list[RenameResult], str]:
    """重命名目录中的PDF文件，输出到 {pdf_dir}_Renamed 文件夹。

    Returns:
        (结果列表, 输出文件夹路径)
    """
    output_dir = pdf_dir.rstrip(os.sep) + "_Renamed"
    os.makedirs(output_dir, exist_ok=True)
    logger.info("输出目录: %s", output_dir)

    pdf_files = [f for f in sorted(os.listdir(pdf_dir)) if f.lower().endswith(".pdf")]
    logger.info("发现 %d 个PDF文件", len(pdf_files))

    results = []
    for filename in pdf_files:
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
                    logger.warning("索引中未找到乙方合同号: %s", party_b_id)
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
                logger.info("Type2文件，开始OCR识别: %s", filename)
                text = extract_text_from_pdf(old_path)
                fields = extract_invoice_fields(text)
                party_b_id = fields.get("party_b_id", "")
                if not party_b_id:
                    logger.warning("OCR未能识别乙方合同号: %s", filename)
                    results.append(RenameResult(
                        original_name=filename, new_name="", party_a_id="",
                        party_b_id="", status="error",
                        message="OCR未能识别乙方合同号"
                    ))
                    continue

                party_a_id = contract_index.get(party_b_id, "")
                if not party_a_id:
                    logger.warning("索引中未找到乙方合同号: %s", party_b_id)
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

            new_path = os.path.join(output_dir, new_name)
            shutil.copy2(old_path, new_path)
            logger.info("重命名: %s → %s", filename, new_name)
            results.append(RenameResult(
                original_name=filename, new_name=new_name,
                party_a_id=party_a_id, party_b_id=party_b_id,
                status="success", message="重命名成功"
            ))

        except Exception as e:
            logger.error("处理文件 '%s' 时出错: %s", filename, e)
            results.append(RenameResult(
                original_name=filename, new_name="", party_a_id="",
                party_b_id=party_b_id or "", status="error", message=str(e)
            ))

    return results, output_dir
