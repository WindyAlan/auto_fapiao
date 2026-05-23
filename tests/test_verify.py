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
        "invoice_date": "2026/05/20",
        "tax_amount": "3237085.00",
        "total_amount": "28137740.66",
    }
    diffs = compare_fields(excel_row, ocr_fields, confidence=0.95)
    assert len(diffs) == 2
    assert diffs[0].field_name == "发票日期"
    assert diffs[1].field_name == "税金金额"
