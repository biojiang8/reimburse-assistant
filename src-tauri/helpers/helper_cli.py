#!/usr/bin/env python3
"""PyInstaller 打包入口：一个二进制分发 sign / reagent / summary / evidence / verify 子命令。

用法: reimburse-helper <sign|reagent|summary|evidence|verify> [args...]
"""

import sys


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: reimburse-helper <sign|reagent|summary|evidence|verify> [args...]", file=sys.stderr)
        return 2
    subcommand = sys.argv[1]
    # 注意：脚本 main(argv) 的约定是 argv 不含程序名
    # （parse_args(None) 时使用 sys.argv[1:]），所以这里传 sys.argv[2:]。
    argv = list(sys.argv[2:])
    if subcommand == "sign":
        import add_invoice_signature

        return add_invoice_signature.main(argv)
    if subcommand == "reagent":
        import reagent_report

        return reagent_report.main(argv)
    if subcommand == "summary":
        import invoice_summary

        return invoice_summary.main(argv)
    if subcommand == "evidence":
        import invoice_evidence

        return invoice_evidence.main(argv)
    if subcommand == "verify":
        import verify_package

        return verify_package.main(argv)
    print(f"unknown subcommand: {subcommand}", file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
