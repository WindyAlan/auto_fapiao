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
        verify_results, output_excel = verify_invoices(output_dir, data["invoice_excel"])
        verify_report = generate_verify_report(verify_results, output_excel, pdf_dir=output_dir)
        assert "发票校验报告" in verify_report
        assert output_excel.endswith("_Verified.xlsx")
        assert os.path.exists(output_excel)

        # 验证_filled文件夹
        filled_dir = os.path.join(tmp_dir, "pdfs_filled")
        assert os.path.isdir(filled_dir), f"_filled文件夹未创建: {filled_dir}"

        # 检查_filled中的文件：有2个发票被填充（invoice_no为空的）
        filled_files = [f for f in os.listdir(filled_dir) if f.endswith(".pdf")]
        assert len(filled_files) == 2, f"期望2个填充文件，实际: {filled_files}"

        # 文件名应该是发票号.pdf
        for f in filled_files:
            name_without_ext = f[:-4]  # 去掉 .pdf
            assert name_without_ext.isdigit(), f"文件名应为纯数字发票号: {f}"

    finally:
        shutil.rmtree(tmp_dir)
