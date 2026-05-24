import logging
import os
from datetime import datetime

from rename import RenameResult
from verify import VerifyResult

logger = logging.getLogger(__name__)


def generate_rename_report(results: list[RenameResult], output_dir: str) -> str:
    """生成重命名报告并保存到文件，返回报告内容"""
    lines = []
    lines.append("=" * 50)
    lines.append("发票重命名报告")
    lines.append(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 50)

    success = [r for r in results if r.status == "success"]
    skipped = [r for r in results if r.status == "skipped"]
    errors = [r for r in results if r.status == "error"]

    lines.append(f"\n总计: {len(results)} 个文件")
    lines.append(f"  成功: {len(success)}")
    lines.append(f"  跳过: {len(skipped)}")
    lines.append(f"  失败: {len(errors)}")

    if success:
        lines.append(f"\n{'─' * 50}")
        lines.append("【成功重命名】")
        for r in success:
            lines.append(f"  {r.original_name}")
            lines.append(f"    → {r.new_name}")
            lines.append(f"    甲方合同号: {r.party_a_id}, 乙方合同号: {r.party_b_id}")

    if skipped:
        lines.append(f"\n{'─' * 50}")
        lines.append("【跳过 - 无法识别的文件】")
        for r in skipped:
            lines.append(f"  {r.original_name}: {r.message}")

    if errors:
        lines.append(f"\n{'─' * 50}")
        lines.append("【失败 - 需要人工处理】")
        for r in errors:
            lines.append(f"  {r.original_name}: {r.message}")
            if r.party_b_id:
                lines.append(f"    乙方合同号: {r.party_b_id}")

    lines.append(f"\n{'=' * 50}")
    lines.append(f"输出目录: {output_dir}")

    report = "\n".join(lines)

    # 保存报告文件
    report_path = os.path.join(output_dir, "rename_report.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report)
    logger.info("重命名报告已保存: %s", report_path)

    return report


def generate_verify_report(results: list[VerifyResult], output_excel: str,
                           pdf_dir: str = "") -> str:
    """生成校验报告并保存到文件，返回报告内容"""
    lines = []
    lines.append("=" * 50)
    lines.append("发票校验报告")
    lines.append(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("=" * 50)

    correct = [r for r in results if not r.diffs]
    filled = [r for r in results if r.filled]
    diff_found = [r for r in results if r.diffs]
    manual = [r for r in results if r.needs_manual]

    lines.append(f"\n总计: {len(results)} 个文件")
    lines.append(f"  全部正确: {len(correct)}")
    lines.append(f"  OCR填充了信息: {len(filled)}")
    lines.append(f"  有差异: {len(diff_found)}")
    lines.append(f"  需手动核查: {len(manual)}")

    if filled:
        lines.append(f"\n{'─' * 50}")
        lines.append("【OCR自动填充】")
        for r in filled:
            lines.append(f"  {r.party_a_id} ({r.pdf_file})")
            for f_desc in r.filled:
                lines.append(f"    {f_desc}")

    filled_pdfs = [r for r in results if r.filled and r.invoice_no]
    if filled_pdfs and pdf_dir:
        import re as _re
        parent_dir = os.path.dirname(pdf_dir)
        dir_basename = os.path.basename(pdf_dir)
        if dir_basename.endswith("_Renamed"):
            filled_dir_name = dir_basename[:-len("_Renamed")] + "_filled"
        else:
            filled_dir_name = dir_basename + "_filled"
        filled_dir = os.path.join(parent_dir, filled_dir_name)
        lines.append(f"\n{'─' * 50}")
        lines.append(f"【已复制到_filled文件夹】({len(filled_pdfs)} 个)")
        lines.append(f"  目录: {filled_dir}")
        for r in filled_pdfs:
            safe_name = _re.sub(r'[\\/:*?"<>|]', '_', r.invoice_no)
            lines.append(f"  {r.pdf_file}")
            lines.append(f"    → {safe_name}.pdf")

    if diff_found:
        lines.append(f"\n{'─' * 50}")
        lines.append("【有差异 - 请人工确认】")
        for r in diff_found:
            lines.append(f"  {r.party_a_id} ({r.pdf_file})")
            for diff in r.diffs:
                lines.append(f"    {diff.field_name}: Excel='{diff.excel_value}', OCR='{diff.ocr_value}' (置信度: {diff.confidence:.1%})")

    if manual:
        lines.append(f"\n{'─' * 50}")
        lines.append("【需手动核查】")
        for r in manual:
            lines.append(f"  {r.party_a_id} ({r.pdf_file})")
            if not r.diffs and not r.party_a_id:
                lines.append(f"    原因: 文件名格式无法识别甲方合同号")
            elif not r.diffs:
                lines.append(f"    原因: Excel中未找到甲方合同号")
            else:
                for diff in r.diffs:
                    if not diff.fixed:
                        lines.append(f"    {diff.field_name}: Excel='{diff.excel_value}', OCR='{diff.ocr_value}' (置信度: {diff.confidence:.1%}, 低于阈值)")

    if correct:
        lines.append(f"\n{'─' * 50}")
        lines.append("【全部正确】")
        for r in correct:
            lines.append(f"  {r.party_a_id} ({r.pdf_file})")
            if r.filled:
                for f_desc in r.filled:
                    lines.append(f"    {f_desc}")

    lines.append(f"\n{'=' * 50}")
    lines.append(f"输出Excel: {output_excel}")

    report = "\n".join(lines)

    # 保存报告文件到Excel同目录
    report_path = os.path.join(os.path.dirname(output_excel), "verify_report.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write(report)
    logger.info("校验报告已保存: %s", report_path)

    return report
