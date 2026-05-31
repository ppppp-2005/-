from __future__ import annotations

import re
from collections import Counter

from docx import Document

from .classifier import classify_paragraph


def styles_used_in_template(doc: Document) -> set[str]:
    return {para.style.name for para in doc.paragraphs if para.text.strip() and para.style and para.style.name}


def build_template_style_profile(doc: Document) -> dict[str, str | None]:
    """从模板里实际出现的段落，推断各级标题/正文对应的样式名。"""
    used = styles_used_in_template(doc)
    if not used:
        return {}

    body_counter: Counter[str] = Counter()
    h1_counter: Counter[str] = Counter()
    h2_counter: Counter[str] = Counter()
    h3_counter: Counter[str] = Counter()
    abs_counter: Counter[str] = Counter()
    kw_counter: Counter[str] = Counter()

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text or not para.style or not para.style.name:
            continue
        name = para.style.name
        kind = classify_paragraph(text, [])  # empty rules -> body, but we also use heuristics

        if re.search(r"^第[\d一二三四五六七八九十百]+章", text):
            h1_counter[name] += 1
        elif re.match(r"^\d+\.\d+\.\d+\s+\S", text):
            h3_counter[name] += 1
        elif re.match(r"^\d+\.\d+\s+\S", text):
            h2_counter[name] += 1
        elif re.match(r"^摘\s*要$|^Abstract$", text, re.I):
            abs_counter[name] += 1
        elif re.match(r"^关键词[：:]|^Keywords[：:]", text, re.I):
            kw_counter[name] += 1
        elif len(text) >= 20:
            body_counter[name] += 1

    def pick(counter: Counter[str], fallback_keywords: tuple[str, ...]) -> str | None:
        if counter:
            return counter.most_common(1)[0][0]
        for style_name in used:
            lower = style_name.lower()
            if any(k in lower or k in style_name for k in fallback_keywords):
                return style_name
        return None

    profile = {
        "body": body_counter.most_common(1)[0][0] if body_counter else pick(Counter(), ("正文", "normal", "body")),
        "heading1": pick(h1_counter, ("标题1", "标题 1", "heading 1", "一级标题", "1级标题")),
        "heading2": pick(h2_counter, ("标题2", "标题 2", "heading 2", "二级标题", "2级标题")),
        "heading3": pick(h3_counter, ("标题3", "标题 3", "heading 3", "三级标题", "3级标题")),
        "abstract_title": pick(abs_counter, ("摘要", "abstract")),
        "keywords": pick(kw_counter, ("关键词", "keyword")),
    }

    # 正文兜底：模板里用得最多的样式
    if not profile["body"]:
        all_used = Counter(para.style.name for para in doc.paragraphs if para.text.strip() and para.style)
        if all_used:
            profile["body"] = all_used.most_common(1)[0][0]

    return profile


from .style_constants import STYLE_CANDIDATES


def resolve_style_with_profile(
    para_type: str,
    profile: dict[str, str | None],
    available: set[str],
    used_in_template: set[str],
) -> str | None:
    """优先用模板里真实用过的样式，避免误选 Word 内置的 Empty Heading 1。"""
    if para_type in profile and profile[para_type] and profile[para_type] in available:
        return profile[para_type]

    # 仅在模板中出现过的样式里再匹配
    for candidate in STYLE_CANDIDATES.get(para_type, STYLE_CANDIDATES["body"]):
        if candidate in used_in_template and candidate in available:
            return candidate

    for candidate in STYLE_CANDIDATES.get(para_type, STYLE_CANDIDATES["body"]):
        if candidate in available:
            return candidate

    if profile.get("body") in available:
        return profile["body"]
    if "Normal" in available:
        return "Normal"
    return next(iter(available), None)
