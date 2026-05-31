from __future__ import annotations

from pathlib import Path

import yaml
from docx import Document

from .classifier import compile_rules, classify_paragraph
from .styles import apply_page_setup, apply_paragraph_style


class ThesisFormatter:
    """基于「规则识别 + 样式套用 + 页面设置」的论文排版器。"""

    def __init__(self, config_path: str | Path):
        self.config_path = Path(config_path)
        with self.config_path.open("r", encoding="utf-8") as f:
            self.config = yaml.safe_load(f)
        self.rules = compile_rules(self.config.get("rules", []))
        self.styles = self.config["styles"]

    def format_document(self, input_path: str | Path, output_path: str | Path) -> dict:
        input_path = Path(input_path)
        output_path = Path(output_path)

        doc = Document(str(input_path))
        stats = {"total": 0, "by_type": {}}

        for section in doc.sections:
            apply_page_setup(section, self.config["page"])

        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if not text:
                continue

            para_type = classify_paragraph(text, self.rules)
            stats["total"] += 1
            stats["by_type"][para_type] = stats["by_type"].get(para_type, 0) + 1

            style_cfg = self._style_for_type(para_type)
            apply_paragraph_style(paragraph, style_cfg)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_path))
        return stats

    def _style_for_type(self, para_type: str) -> dict:
        if para_type in self.styles:
            return self.styles[para_type]

        body = dict(self.styles["body"])
        if para_type != "body":
            body["clear_first_line_indent"] = True
        return body


def load_formatter(config_path: str | Path) -> ThesisFormatter:
    return ThesisFormatter(config_path)
