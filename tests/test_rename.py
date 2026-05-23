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
