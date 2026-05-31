from __future__ import annotations

import argparse
from pathlib import Path

from src.formatter import load_formatter


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="中文学位论文 Word 排版工具（规则识别 + 样式套用）"
    )
    parser.add_argument("input", help="输入 docx 文件路径")
    parser.add_argument(
        "-o",
        "--output",
        help="输出 docx 文件路径（默认：output/原文件名_formatted.docx）",
    )
    parser.add_argument(
        "-c",
        "--config",
        default="config/default_cn_thesis.yaml",
        help="排版规则配置文件",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        raise SystemExit(f"找不到输入文件: {input_path}")

    output_path = (
        Path(args.output)
        if args.output
        else Path("output") / f"{input_path.stem}_formatted.docx"
    )

    formatter = load_formatter(args.config)
    stats = formatter.format_document(input_path, output_path)

    print(f"排版完成: {output_path}")
    print(f"处理段落数: {stats['total']}")
    print("段落分类统计:")
    for key, value in sorted(stats["by_type"].items()):
        print(f"  - {key}: {value}")


if __name__ == "__main__":
    main()
