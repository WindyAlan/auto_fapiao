import argparse
import logging
import sys


def setup_logging(verbose: bool):
    level = logging.DEBUG if verbose else logging.INFO
    fmt = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    logging.basicConfig(level=level, format=fmt, stream=sys.stderr)


def cmd_rename(args):
    from contract import load_contract_index
    from rename import rename_files
    from report import generate_rename_report

    logger = logging.getLogger("rename")

    logger.info("加载合同号索引: %s", args.excel)
    contract_index = load_contract_index(args.excel)
    logger.info("找到 %d 条合同记录", len(contract_index))

    logger.info("扫描目录: %s", args.dir)
    results, output_dir = rename_files(args.dir, contract_index)

    report = generate_rename_report(results, output_dir)
    print(report)


def cmd_verify(args):
    from report import generate_verify_report
    from verify import verify_invoices

    logger = logging.getLogger("verify")

    logger.info("校验发票: %s", args.dir)
    logger.info("对照Excel: %s", args.excel)
    results = verify_invoices(args.dir, args.excel)

    report = generate_verify_report(results, args.excel)
    print(report)


def cmd_generate_test_data(args):
    from tests.conftest import generate_all_test_data

    logger = logging.getLogger("generate")

    logger.info("生成测试数据到: %s", args.output)
    generate_all_test_data(args.output)
    print("完成")


def main():
    parser = argparse.ArgumentParser(description="发票自动重命名和校验工具")
    parser.add_argument("-v", "--verbose", action="store_true", help="输出详细调试信息")
    subparsers = parser.add_subparsers(dest="command", required=True)

    # rename
    p_rename = subparsers.add_parser("rename", help="根据合同号索引重命名PDF文件")
    p_rename.add_argument("--dir", required=True, help="PDF文件所在目录")
    p_rename.add_argument("--excel", required=True, help="合同号索引Excel文件路径")

    # verify
    p_verify = subparsers.add_parser("verify", help="校验发票信息与Excel是否一致")
    p_verify.add_argument("--dir", required=True, help="PDF文件所在目录")
    p_verify.add_argument("--excel", required=True, help="发票验证Excel文件路径")

    # generate-test-data
    p_gen = subparsers.add_parser("generate-test-data", help="生成测试用的mock数据")
    p_gen.add_argument("--output", required=True, help="输出目录")

    args = parser.parse_args()
    setup_logging(args.verbose)

    if args.command == "rename":
        cmd_rename(args)
    elif args.command == "verify":
        cmd_verify(args)
    elif args.command == "generate-test-data":
        cmd_generate_test_data(args)


if __name__ == "__main__":
    main()
