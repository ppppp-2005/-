from __future__ import annotations

import re
from dataclasses import dataclass

from docx.text.paragraph import Paragraph

from .classifier import Rule, classify_paragraph

# 文章/模板里 Word 样式名 -> 段落类型
STYLE_NAME_TO_TYPE: dict[str, list[str]] = {
    "heading1": [
        "heading 1",
        "heading1",
        "标题 1",
        "标题1",
        "一级标题",
        "1级标题",
        "chapter",
        "章标题",
    ],
    "heading2": [
        "heading 2",
        "heading2",
        "标题 2",
        "标题2",
        "二级标题",
        "2级标题",
    ],
    "heading3": [
        "heading 3",
        "heading3",
        "标题 3",
        "标题3",
        "三级标题",
        "3级标题",
    ],
    "heading4": [
        "heading 4",
        "heading4",
        "标题 4",
        "标题4",
        "四级标题",
        "4级标题",
    ],
    "abstract_title": ["摘要", "abstract", "abstracttitle", "摘要标题"],
    "abstract_body": ["摘要正文", "abstract body"],
    "keywords": ["关键词", "keywords", "关键字"],
    "references_title": ["参考文献", "references", "参考书目"],
    "acknowledgment": ["致谢", "acknowledgment", "acknowledgement"],
    "appendix_title": ["附录", "appendix"],
    "figure_caption": ["图标题", "figure caption", "caption"],
    "table_caption": ["表标题", "table caption"],
    "body": ["正文", "normal", "body text", "body", "正文文本", "文本"],
}


def _normalize_style(name: str) -> str:
    return re.sub(r"\s+", "", name.lower())


def detect_type_from_style(style_name: str | None) -> str | None:
    if not style_name:
        return None
    norm = _normalize_style(style_name)
    for para_type, keywords in STYLE_NAME_TO_TYPE.items():
        for kw in keywords:
            kn = _normalize_style(kw)
            if kn == norm or kn in norm or norm in kn:
                return para_type
    return None


def detect_type_from_text(text: str, rules: list[Rule]) -> str:
    normalized = text.strip()
    if not normalized:
        return "empty"

    extra_patterns: list[tuple[str, list[str]]] = [
        (
            "heading1",
            [
                r"^第[\d一二三四五六七八九十百]+章",
                r"^第[\d一二三四五六七八九十百]+节",
                r"^[一二三四五六七八九十]+、\s*\S",
                r"^绪论$",
                r"^引言$",
                r"^结论$",
                r"^总结$",
            ],
        ),
        (
            "heading2",
            [
                r"^\d+\.\d+\s+\S",
                r"^（[一二三四五六七八九十]+）",
                r"^\([一二三四五六七八九十]+\)",
            ],
        ),
        ("heading3", [r"^\d+\.\d+\.\d+\s+\S", r"^\d+\.\d+\.\d+\.\d+\s+\S"]),
        ("heading4", [r"^\d+\.\d+\.\d+\.\d+\s+\S"]),
        ("abstract_title", [r"^摘\s*要$", r"^Abstract$"]),
        ("keywords", [r"^关键词[：:]", r"^Keywords[：:]", r"^关键字[：:]"]),
        ("references_title", [r"^参考文献$", r"^References$"]),
        ("acknowledgment", [r"^致\s*谢$"]),
        ("appendix_title", [r"^附\s*录"]),
        ("figure_caption", [r"^图[\d\-\.]+", r"^Figure\s+\d+"]),
        ("table_caption", [r"^表[\d\-\.]+", r"^Table\s+\d+"]),
    ]

    for para_type, patterns in extra_patterns:
        for pat in patterns:
            if re.search(pat, normalized, re.I):
                return para_type

    typed = classify_paragraph(normalized, rules)
    return typed if typed != "empty" else "body"


def classify_article_paragraph(para: Paragraph, rules: list[Rule]) -> str:
    """综合文章段落的样式名 + 文字内容，判断属于哪一类。"""
    text = para.text.strip()
    if not text:
        return "empty"

    style_name = para.style.name if para.style else None
    from_style = detect_type_from_style(style_name)
    from_text = detect_type_from_text(text, rules)

    # 样式名是标题时优先信任（用户已在 Word 里标好过级别）
    if from_style and from_style.startswith("heading"):
        return from_style
    if from_style in ("abstract_title", "keywords", "references_title", "acknowledgment"):
        return from_style

    # 文字能识别为标题时也用标题
    if from_text.startswith("heading") or from_text in (
        "abstract_title",
        "keywords",
        "references_title",
        "acknowledgment",
        "appendix_title",
        "figure_caption",
        "table_caption",
    ):
        return from_text

    if from_style:
        return from_style

    return from_text


@dataclass
class ArticleParagraph:
    text: str
    para_type: str
    source_style: str | None
