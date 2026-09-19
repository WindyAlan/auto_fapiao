import logging
import re
from decimal import Decimal, InvalidOperation
from functools import lru_cache

import fitz
import numpy as np

logger = logging.getLogger(__name__)

# 文本层阈值：少于这个字数认为无文本层
TEXT_LAYER_THRESHOLD = 50


def extract_text_from_pdf(pdf_path: str) -> str:
    """从PDF提取文本。优先用PyMuPDF文本层，失败则fallback到OCR。"""
    text, _ = extract_pdf_content(pdf_path)
    return text


def extract_pdf_content(pdf_path: str) -> tuple[str, float]:
    """提取PDF文本及其置信度，避免校验流程对同一文件重复OCR。"""
    logger.debug("提取PDF文本: %s", pdf_path)
    text = _extract_text_layer(pdf_path)
    if len(text) >= TEXT_LAYER_THRESHOLD:
        logger.debug("文本层提取成功，长度=%d", len(text))
        # 文本层无需OCR；字段差异仍会被标记为需要人工核查。
        return text, 1.0
    logger.info("文本层不足(len=%d)，fallback到OCR: %s", len(text), pdf_path)
    return _run_ocr(pdf_path)


def _extract_text_layer(pdf_path: str) -> str:
    """用PyMuPDF提取PDF文本层"""
    with fitz.open(pdf_path) as doc:
        text_parts = [page.get_text() for page in doc]
    return "\n".join(text_parts)


@lru_cache(maxsize=1)
def _get_ocr_engine():
    """延迟初始化并复用OCR模型，避免每个PDF重复加载模型。"""
    try:
        from paddleocr import PaddleOCR
    except ImportError as e:
        logger.error("PaddleOCR导入失败: %s", e)
        logger.error("请确认已安装 paddlepaddle 和 paddleocr: uv pip install -r requirements.txt")
        raise

    try:
        logger.info("初始化PaddleOCR...")
        return PaddleOCR(lang="ch")
    except Exception as e:
        logger.error("PaddleOCR初始化失败: %s", e)
        logger.error("Windows用户请确认已安装 Visual C++ Redistributable (vc_redist.x64.exe)")
        raise


def _run_ocr(pdf_path: str) -> tuple[str, float]:
    """将PDF转图片后OCR一次，同时返回文本和平均置信度。"""
    ocr = _get_ocr_engine()
    all_text = []
    confidences = []
    with fitz.open(pdf_path) as doc:
        page_count = len(doc)
        logger.info("开始OCR识别，共%d页", page_count)
        for page_num, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            image = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                pix.height, pix.width, pix.n,
            )

            logger.debug("OCR第%d/%d页...", page_num, page_count)
            result = ocr.ocr(image)
            if result and result[0]:
                for line in result[0]:
                    all_text.append(line[1][0])
                    confidences.append(line[1][1])

    logger.info("OCR完成，识别到%d段文本", len(all_text))
    confidence = sum(confidences) / len(confidences) if confidences else 0.0
    logger.debug("OCR置信度: avg=%.3f, samples=%d", confidence, len(confidences))
    return "\n".join(all_text), confidence


def ocr_pdf(pdf_path: str) -> str:
    """将PDF转图片后用PaddleOCR识别。"""
    text, _ = _run_ocr(pdf_path)
    return text


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

    # 一张发票可能包含多个明细行，保留每个数量以便与同一 PO 的多行 Excel 匹配。
    quantity_values = []
    for match in re.finditer(r"数量\s*(?:[：:]\s*)?([+-]?\d[\d,]*(?:\.\d+)?)", text):
        try:
            quantity_values.append(Decimal(match.group(1).replace(",", "")))
        except InvalidOperation:
            logger.warning("无法解析发票数量: %s", match.group(1))
    if quantity_values:
        fields["billing_qty_values"] = quantity_values
        fields["billing_qty"] = str(sum(quantity_values))

    # 乙方合同号（备注栏）— 数字+大写字母，可能有空格
    # 先尝试找"备注"后面的内容
    m = re.search(r"备注[：:]?\s*.*?(\d+\s*[A-Z])", text, re.DOTALL)
    if not m:
        # 兜底：直接在整个文本里找 13 位数字+字母的合同号模式
        m = re.search(r"(\d{13}\s*[A-Z])", text)
    if m:
        fields["party_b_id"] = m.group(1).replace(" ", "")

    logger.debug("提取到发票字段: %s", fields)
    return fields


def get_ocr_confidence(pdf_path: str) -> float:
    """获取OCR识别的平均置信度。"""
    try:
        _, confidence = _run_ocr(pdf_path)
        return confidence
    except Exception as e:
        logger.error("OCR置信度获取失败: %s", e)
        return 0.0
