import os
import shutil
import tempfile

from contract import load_contract_index
from rename import rename_files
from report import generate_rename_report, generate_verify_report
from tests.conftest import generate_all_test_data
from verify import verify_invoices


def test_full_workflow():
    """端到端测试：生成数据 → 重命名 → 校验"""
    tmp_dir = tempfile.mkdtemp()
    try:
        # 生成测试数据
        data = generate_all_test_data(tmp_dir)

        # 任务1：重命名
        contract_index = load_contract_index(data["contract_excel"])
        rename_results, output_dir = rename_files(data["pdf_dir"], contract_index)

        success_count = sum(1 for r in rename_results if r.status == "success")
        assert success_count > 0, f"没有成功重命名的文件: {rename_results}"

        # 验证重命名后的文件名格式
        for r in rename_results:
            if r.status == "success":
                assert "-" in r.new_name, f"文件名缺少甲方合同号: {r.new_name}"

        # 生成重命名报告
        rename_report = generate_rename_report(rename_results, output_dir)
        assert "发票重命名报告" in rename_report
        assert os.path.exists(os.path.join(output_dir, "rename_report.txt"))

        # 任务2：校验（校验输出目录中的重命名文件）
        verify_results = verify_invoices(output_dir, data["invoice_excel"])
        verify_report = generate_verify_report(verify_results, data["invoice_excel"])
        assert "发票校验报告" in verify_report
        assert os.path.exists(os.path.join(os.path.dirname(data["invoice_excel"]), "verify_report.txt"))

    finally:
        shutil.rmtree(tmp_dir)
