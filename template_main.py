from __future__ import annotations

import argparse
import json
from pathlib import Path

from src.template_mode import apply_template_format, list_paragraph_styles, merge_user_doc_into_template


def main() -> None:
    parser = argparse.ArgumentParser(
        description="选择性套模板格式：正文有什么改什么；模板没有的格式保持原样"
    )
    parser.add_argument("article", nargs="?", help="你的论文/文章 docx")
    parser.add_argument("template", nargs="?", help="格式模板 docx")
    parser.add_argument(
        "-o",
        "--output",
        default="output/from_template.docx",
        help="输出文件路径",
    )
    parser.add_argument(
        "--mode",
        choices=["apply", "fill"],
        default="apply",
        help="apply=按正文选择性套格式(推荐); fill=填入模板空行",
    )
    parser.add_argument(
        "-c",
        "--config",
        default="config/default_cn_thesis.yaml",
        help="标题识别规则配置",
    )
    parser.add_argument(
        "--list-styles",
        metavar="TEMPLATE.docx",
        help="只查看模板里有哪些段落样式",
    )
    args = parser.parse_args()

    if args.list_styles:
        styles = list_paragraph_styles(args.list_styles)
        print(f"模板「{args.list_styles}」中的段落样式 ({len(styles)} 个):")
        for name in styles:
            print(f"  - {name}")
        return

    if not args.article or not args.template:
        parser.error(
            "请提供：文章.docx 模板.docx\n"
            "示例: python template_main.py examples\\123.docx templates\\111.docx"
        )

    article = Path(args.article)
    template = Path(args.template)
    if not article.exists():
        raise SystemExit(f"找不到文章: {article}")
    if not template.exists():
        raise SystemExit(f"找不到模板: {template}")

    if args.mode == "fill":
        count = merge_user_doc_into_template(article, template, args.output, mode="fill")
        print(f"已填入 {count} 个段落到模板空位")
        print(f"输出: {args.output}")
        return

    print("paper-formatter v0.7（保留模板封面 + 正文填入模板格式）")
    print("规则: 保留模板封面/前置页；正文来自你的文章；有模板格式才改")
    stats = apply_template_format(article, template, args.output, args.config)

    print(f"\n排版完成: {stats['output']}")
    if stats.get("notice"):
        print(f"注意: {stats['notice']}")

    if stats.get("preserve_template_front"):
        print(f"已保留模板前置页段落数: {stats.get('template_front_preserved_paragraphs', 0)}")

    print(f"文档总段落数: {stats.get('total_paragraphs_in_document', '?')}（含空行）")
    print(f"有内容的段落: {stats.get('content_paragraphs', stats.get('total', '?'))}")
    print(f"已套用模板格式: {stats.get('format_applied_count', 0)} 段")
    print(f"保持原格式未改: {stats.get('format_unchanged_count', 0)} 段")

    print("\n模板里找到的格式范例:")
    for k, v in sorted(stats.get("template_format_library", {}).items()):
        print(f"  - {k}: {v}")

    print("\n文章段落类型统计:")
    for k, v in sorted(stats["by_type"].items()):
        print(f"  - {k}: {v} 段")

    if stats.get("format_applied"):
        print("\n已套用模板格式:")
        for k, v in sorted(stats["format_applied"].items()):
            print(f"  - {k}: {v} 段")

    if stats.get("format_unchanged_count", 0) > 0:
        print("\n保持原格式（模板无对应范例）示例:")
        shown = 0
        for k in sorted(stats.get("format_unchanged", {})):
            print(f"  - {k}")
            shown += 1
            if shown >= 8:
                print("  - ...")
                break

    report_path = Path(args.output).parent / "template_apply_report.json"
    with report_path.open("w", encoding="utf-8") as f:
        json.dump(
            {
                "version": "0.7",
                "preserve_template_front": stats.get("preserve_template_front"),
                "template_front_preserved_paragraphs": stats.get(
                    "template_front_preserved_paragraphs"
                ),
                "article": str(article.resolve()),
                "template": str(template.resolve()),
                "output": stats["output"],
                "total_paragraphs_in_document": stats.get("total_paragraphs_in_document"),
                "content_paragraphs": stats.get("content_paragraphs"),
                "format_applied_count": stats.get("format_applied_count"),
                "format_unchanged_count": stats.get("format_unchanged_count"),
                "by_type": stats["by_type"],
                "template_format_library": stats.get("template_format_library"),
                "template_style_profile": stats.get("template_style_profile"),
                "abstract_page_preface": stats.get("abstract_page_preface"),
                "format_applied": stats.get("format_applied"),
            },
            f,
            ensure_ascii=False,
            indent=2,
        )
    print(f"\n报告: {report_path}")


if __name__ == "__main__":
    main()
