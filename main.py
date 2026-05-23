import argparse


def cmd_rename(args):
    from contract import load_contract_index
    from rename import rename_files

    print(f"加载合同号索引: {args.excel}")
    contract_index = load_contract_index(args.excel)
    print(f"找到 {len(contract_index)} 条合同记录")

    print(f"扫描目录: {args.dir}")
    results = rename_files(args.dir, contract_index)

    success = sum(1 for r in results if r.status == "success")
    skipped = sum(1 for r in results if r.status == "skipped")
    errors = sum(1 for r in results if r.status == "error")

    for r in results:
        if r.status == "success":
            print(f"  ✓ {r.original_name} → {r.new_name}")
        elif r.status == "error":
            print(f"  ✗ {r.original_name}: {r.message}")
        else:
            print(f"  - {r.original_name}: {r.message}")

    print(f"\n完成: {success} 重命名, {skipped} 跳过, {errors} 错误")


def cmd_verify(args):
    from verify import format_report, verify_invoices

    print(f"校验发票: {args.dir}")
    print(f"对照Excel: {args.excel}")
    results = verify_invoices(args.dir, args.excel)
    print(format_report(results))


def cmd_generate_test_data(args):
    from tests.conftest import generate_all_test_data

    print(f"生成测试数据到: {args.output}")
    generate_all_test_data(args.output)
    print("完成")


def main():
    parser = argparse.ArgumentParser(description="发票自动重命名和校验工具")
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
    if args.command == "rename":
        cmd_rename(args)
    elif args.command == "verify":
        cmd_verify(args)
    elif args.command == "generate-test-data":
        cmd_generate_test_data(args)


if __name__ == "__main__":
    main()
