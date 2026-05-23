import shutil
import tempfile

from contract import load_contract_index
from rename import rename_files
from tests.conftest import generate_all_test_data
from verify import format_report, verify_invoices


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
