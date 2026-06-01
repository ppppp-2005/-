from __future__ import annotations

import re
from collections import defaultdict

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph

from .format_clone import exemplar_style_label, find_exemplar_paragraphs
from .paragraph_detect import detect_type_from_style, detect_type_from_text


def is_centered(para: Paragraph) -> bool:
    align = para.alignment
    if align is None:
        align = para.paragraph_format.alignment
    return align == WD_ALIGN_PARAGRAPH.CENTER


def _heading_score(text: str, level: int) -> int:
    """参考 word_chat：在多个候选标题段落里选格式最可信的一个。"""
    score = 0
    if level == 1 and re.match(r"^第\s*[\d一二三四五六七八九十百]+\s*章\s+\S", text):
        score += 10
    elif level == 2 and re.match(r"^\d+\.\d+\s+\S", text):
        score += 10
    elif level == 3 and re.match(r"^\d+\.\d+\.\d+\s+\S", text):
        score += 10
    if len(text) < 80:
        score += 2
    if "格式" in text or "要求" in text:
        score -= 8
    return score


def _pick_best_heading_paragraph(
    paragraphs: list[Paragraph], level: int, pattern: str
) -> Paragraph | None:
    candidates = [
        p
        for p in paragraphs
        if re.match(pattern, p.text.strip()) and len(p.text.strip()) < 120
    ]
    if not candidates:
        return None
    return max(candidates, key=lambda p: _heading_score(p.text.strip(), level))


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
    template_doc, rules: list, start_index: int = 0
) -> dict[str, Paragraph]:
    """从模板【正文区】收集各级标题/正文的「格式范例段落」。

    start_index 之前为封面/前置页，其格式（如居中大标题）不应作为正文范例。
    """
    library: dict[str, Paragraph] = {}
    paras = template_doc.paragraphs[start_index:]

    # 0) 优先挑一个真正的正文段落作为 body 范例：非居中、较长、内容判定为正文
    for para in paras:
        text = para.text.strip()
        if len(text) < 30 or is_centered(para):
            continue
        style_type = detect_type_from_style(para.style.name if para.style else None)
        if style_type in (None, "body") and detect_type_from_text(text, rules) == "body":
            library["body"] = para
            break

    # 1) 按模板段落的 Word 样式名归类（body 排除居中段落）
    for para in paras:
        text = para.text.strip()
        if not text or not para.style or not para.style.name:
            continue
        para_type = detect_type_from_style(para.style.name)
        if para_type and para_type not in library:
            if para_type == "body" and is_centered(para):
                continue
            library[para_type] = para

    # 2) 按模板文字内容归类（第X章、1.1、摘要…）
    for para in paras:
        text = para.text.strip()
        if not text:
            continue
        para_type = detect_type_from_text(text, rules)
        if para_type not in ("empty", "body") and para_type not in library:
            library[para_type] = para

    # 3) 内容启发式范例
    for key, para in find_exemplar_paragraphs(template_doc, start_index).items():
        if key not in library:
            library[key] = para

    # 4) 同一样式名下兜底；body 选最长的非居中段落
    by_style: dict[str, list[Paragraph]] = defaultdict(list)
    for para in paras:
        if para.text.strip() and para.style and para.style.name:
            by_style[para.style.name].append(para)

    for style_name, style_paras in by_style.items():
        para_type = detect_type_from_style(style_name)
        if para_type and para_type not in library:
            library[para_type] = style_paras[0]
    if "body" not in library:
        non_center = [p for p in paras if p.text.strip() and not is_centered(p)]
        if non_center:
            longest = max(non_center, key=lambda p: len(p.text))
            if len(longest.text.strip()) >= 10:
                library["body"] = longest

    # 5) 摘要正文：摘要标题后面第一段长文本
    for i, para in enumerate(paras):
        text = para.text.strip()
        if re.match(r"^摘\s*要$|^Abstract$", text, re.I):
            for nxt in paras[i + 1 : i + 6]:
                if len(nxt.text.strip()) >= 20:
                    library.setdefault("abstract_body", nxt)
                    break

    # 6) 参考文献条目：必须在「参考文献」标题之后，且形如 [1] 开头
    in_references = False
    for para in paras:
        t = para.text.strip()
        if re.match(r"^参考文献$|^References$", t, re.I):
            in_references = True
            continue
        if in_references and re.match(r"^\[\d+\]", t):
            library.setdefault("references_body", para)
            break

    # 7) 标题范例：在同类段落里选格式最可信的。
    #    带章节编号（第X章 / X.Y）的标题优先覆盖「按样式名」选出的范例，
    #    避免把封面/正文里的论文大标题误当成章节一级标题。
    for para_type, level, pattern in [
        ("heading1", 1, r"^第\s*[\d一二三四五六七八九十百]+\s*章\s+\S"),
        ("heading2", 2, r"^\d+\.\d+\s+\S"),
        ("heading3", 3, r"^\d+\.\d+\.\d+\s+\S"),
    ]:
        best = _pick_best_heading_paragraph(paras, level, pattern)
        if best and (para_type == "heading1" or para_type not in library):
            library[para_type] = best

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
