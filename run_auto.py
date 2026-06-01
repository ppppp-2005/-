"""一键运行排版并写入 output/run_result.txt"""
from __future__ import annotations

import json
import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

OUT = ROOT / "output" / "run_result.txt"
ARTICLE = ROOT / "examples" / "123.docx"
TEMPLATE = ROOT / "templates" / "111.docx"
OUTPUT = ROOT / "output" / "from_template.docx"


def main() -> int:
    lines: list[str] = []

    def log(msg: str) -> None:
        lines.append(msg)
        print(msg)

    try:
        log("=== paper-formatter 自动运行 ===")
        log(f"文章: {ARTICLE} 存在={ARTICLE.exists()}")
        log(f"模板: {TEMPLATE} 存在={TEMPLATE.exists()}")

        if not ARTICLE.exists():
            log("错误: 找不到 examples/123.docx")
            return 1
        if not TEMPLATE.exists():
            log("错误: 找不到 templates/111.docx")
            return 1

        from src.template_mode import apply_template_format

        log("\npaper-formatter v0.7（保留模板封面 + 正文填入模板格式）")
        stats = apply_template_format(ARTICLE, TEMPLATE, OUTPUT)

        if stats.get("notice"):
            log(f"\n注意: {stats['notice']}")

        out_file = Path(stats["output"])
        log(f"\n排版完成: {out_file}")
        log(f"保留模板前置页: {stats.get('template_front_preserved_paragraphs', 0)} 段")
        log(f"文档总段落: {stats.get('total_paragraphs_in_document')}")
        log(f"有内容段落: {stats.get('content_paragraphs')}")
        log(f"已套模板格式: {stats.get('format_applied_count')} 段")
        log(f"保持原格式: {stats.get('format_unchanged_count')} 段")

        log("\n模板格式范例:")
        for k, v in sorted(stats.get("template_format_library", {}).items()):
            log(f"  - {k}: {v}")

        log("\n段落类型:")
        for k, v in sorted(stats.get("by_type", {}).items()):
            log(f"  - {k}: {v} 段")

        report = {
            "version": "0.7",
            "preserve_template_front": stats.get("preserve_template_front"),
            "template_front_preserved_paragraphs": stats.get(
                "template_front_preserved_paragraphs"
            ),
            "article": str(ARTICLE),
            "template": str(TEMPLATE),
            "output": stats["output"],
            "total_paragraphs_in_document": stats.get("total_paragraphs_in_document"),
            "content_paragraphs": stats.get("content_paragraphs"),
            "format_applied_count": stats.get("format_applied_count"),
            "format_unchanged_count": stats.get("format_unchanged_count"),
            "by_type": stats.get("by_type"),
            "template_format_library": stats.get("template_format_library"),
            "template_style_profile": stats.get("template_style_profile"),
            "format_applied": stats.get("format_applied"),
        }
        report_path = ROOT / "output" / "template_apply_report.json"
        report_path.write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        log(f"\n报告: {report_path}")

        if out_file.exists():
            log(f"输出文件大小: {out_file.stat().st_size} 字节")

        return 0
    except Exception:
        log("\n运行失败:")
        log(traceback.format_exc())
        return 1
    finally:
        OUT.parent.mkdir(parents=True, exist_ok=True)
        OUT.write_text("\n".join(lines), encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
