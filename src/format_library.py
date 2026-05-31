from __future__ import annotations

import re
from collections import defaultdict

from docx.text.paragraph import Paragraph

from .format_clone import exemplar_style_label, find_exemplar_paragraphs
from .paragraph_detect import detect_type_from_style, detect_type_from_text


# 应用格式时的兜底顺序：找不到三级标题范例就用二级，以此类推
FORMAT_FALLBACK: dict[str, list[str]] = {
    "heading4": ["heading4", "heading3", "heading2", "heading1", "body"],
    "heading3": ["heading3", "heading2", "heading1", "body"],
    "heading2": ["heading2", "heading1", "body"],
    "heading1": ["heading1", "body"],
    "abstract_title": ["abstract_title", "heading1", "body"],
    "abstract_body": ["abstract_body", "body"],
    "keywords": ["keywords", "body"],
    "references_title": ["references_title", "heading1", "body"],
    "references_body": ["references_body", "body"],
    "acknowledgment": ["acknowledgment", "heading1", "body"],
    "appendix_title": ["appendix_title", "heading1", "body"],
    "figure_caption": ["figure_caption", "body"],
    "table_caption": ["table_caption", "body"],
    "body": ["body"],
}


def build_template_format_library(
    template_doc, rules: list
) -> dict[str, Paragraph]:
    """从模板收集各级标题/正文的「格式范例段落」。"""
    library: dict[str, Paragraph] = {}

    # 1) 按模板段落的 Word 样式名归类
    for para in template_doc.paragraphs:
        text = para.text.strip()
        if not text or not para.style or not para.style.name:
            continue
        para_type = detect_type_from_style(para.style.name)
        if para_type and para_type not in library:
            library[para_type] = para

    # 2) 按模板文字内容归类（第X章、1.1、摘要…）
    for para in template_doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        para_type = detect_type_from_text(text, rules)
        if para_type not in ("empty", "body") and para_type not in library:
            library[para_type] = para

    # 3) 内容启发式范例
    for key, para in find_exemplar_paragraphs(template_doc).items():
        if key not in library:
            library[key] = para

    # 4) 同一样式名下：选文字最长的作为正文范例，最短的标题样作为标题
    by_style: dict[str, list[Paragraph]] = defaultdict(list)
    for para in template_doc.paragraphs:
        if para.text.strip() and para.style and para.style.name:
            by_style[para.style.name].append(para)

    for style_name, paras in by_style.items():
        para_type = detect_type_from_style(style_name)
        if para_type and para_type not in library:
            library[para_type] = paras[0]
        if "body" not in library:
            longest = max(paras, key=lambda p: len(p.text))
            if len(longest.text.strip()) >= 10:
                library["body"] = longest

    # 5) 摘要正文：摘要标题后面第一段长文本
    for i, para in enumerate(template_doc.paragraphs):
        text = para.text.strip()
        if re.match(r"^摘\s*要$|^Abstract$", text, re.I):
            for nxt in template_doc.paragraphs[i + 1 : i + 6]:
                if len(nxt.text.strip()) >= 20:
                    library.setdefault("abstract_body", nxt)
                    break

    # 6) 参考文献条目
    for para in template_doc.paragraphs:
        t = para.text.strip()
        if re.match(r"^\[\d+\]|^\d+\.", t) and "references_body" not in library:
            library["references_body"] = para

    return library


def pick_format_exemplar_strict(
    library: dict[str, Paragraph], para_type: str
) -> Paragraph | None:
    """仅当模板里明确有该类型格式时才返回，不做降级兜底。"""
    return library.get(para_type)


def pick_format_exemplar(
    library: dict[str, Paragraph], para_type: str
) -> Paragraph | None:
    """旧逻辑：带降级兜底（fill 模式等仍可使用）。"""
    for key in FORMAT_FALLBACK.get(para_type, [para_type, "body"]):
        if key in library:
            return library[key]
    return library.get("body")


def library_summary(library: dict[str, Paragraph]) -> dict[str, str]:
    return {k: exemplar_style_label(v) for k, v in library.items()}
