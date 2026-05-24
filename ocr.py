import logging
import re

import fitz

logger = logging.getLogger(__name__)

# 文本层阈值：少于这个字数认为无文本层
TEXT_LAYER_THRESHOLD = 50


def extract_text_from_pdf(pdf_path: str) -> str:
    """从PDF提取文本。优先用PyMuPDF文本层，失败则fallback到OCR。"""
    logger.debug("提取PDF文本: %s", pdf_path)
    text = _extract_text_layer(pdf_path)
    if len(text) >= TEXT_LAYER_THRESHOLD:
        logger.debug("文本层提取成功，长度=%d", len(text))
        return text
    logger.info("文本层不足(len=%d)，fallback到OCR: %s", len(text), pdf_path)
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
    try:
        from paddleocr import PaddleOCR
    except ImportError as e:
        logger.error("PaddleOCR导入失败: %s", e)
        logger.error("请确认已安装 paddlepaddle 和 paddleocr: uv pip install -r requirements.txt")
        raise

    try:
        logger.info("初始化PaddleOCR...")
        ocr = PaddleOCR(lang="ch")
    except Exception as e:
        logger.error("PaddleOCR初始化失败: %s", e)
        logger.error("Windows用户请确认已安装 Visual C++ Redistributable (vc_redist.x64.exe)")
        raise

    doc = fitz.open(pdf_path)
    all_text = []
    page_count = len(doc)
    logger.info("开始OCR识别，共%d页", page_count)

    for page_num in range(page_count):
        page = doc[page_num]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_bytes = pix.tobytes("png")

        logger.debug("OCR第%d/%d页...", page_num + 1, page_count)
        result = ocr.ocr(img_bytes)
        if result and result[0]:
            for line in result[0]:
                all_text.append(line[1][0])

    doc.close()
    logger.info("OCR完成，识别到%d段文本", len(all_text))
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

    logger.debug("提取到发票字段: %s", fields)
    return fields


def get_ocr_confidence(pdf_path: str) -> float:
    """获取OCR识别的平均置信度"""
    try:
        from paddleocr import PaddleOCR
        ocr = PaddleOCR(lang="ch")
    except Exception as e:
        logger.error("OCR置信度获取失败: %s", e)
        return 0.0
    doc = fitz.open(pdf_path)
    confidences = []

    for page_num in range(len(doc)):
        page = doc[page_num]
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img_bytes = pix.tobytes("png")
        result = ocr.ocr(img_bytes)
        if result and result[0]:
            for line in result[0]:
                confidences.append(line[1][1])

    doc.close()
    avg = sum(confidences) / len(confidences) if confidences else 0.0
    logger.debug("OCR置信度: avg=%.3f, samples=%d", avg, len(confidences))
    return avg
