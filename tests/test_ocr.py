import os
import tempfile

import fitz

from ocr import extract_text_from_pdf, extract_invoice_fields


def test_extract_text_from_pdf_with_text_layer():
    """从有文本层的PDF中提取文本"""
    doc = fitz.open()
    page = doc.new_page()
    # Use ASCII text since default font doesn't support Chinese
    text = "Invoice No: 26317000001473818243\nDate: 2026/05/19\nTax: 3237085.21"
    page.insert_text((72, 72), text)

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        doc.save(f.name)
        tmp_path = f.name
    doc.close()

    try:
        result = extract_text_from_pdf(tmp_path)
        assert "26317000001473818243" in result
        assert "3237085.21" in result
    finally:
        os.unlink(tmp_path)


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


def test_extract_invoice_fields_collects_all_billing_quantities():
    """数量字段可重复出现，供同一PO的多行数据逐项校验。"""
    fields = extract_invoice_fields("数量：3\n数量: 4.5\n数量 3")

    assert fields["billing_qty"] == "10.5"
    assert [str(value) for value in fields["billing_qty_values"]] == ["3", "4.5", "3"]
