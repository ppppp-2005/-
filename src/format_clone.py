from __future__ import annotations

import re
from copy import deepcopy

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


def find_exemplar_paragraphs(template_doc) -> dict[str, Paragraph]:
    """在模板里找「范例段落」，用于复制真实格式（含手工设置的字体）。"""
    exemplars: dict[str, Paragraph] = {}

    for para in template_doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        if "heading1" not in exemplars and re.search(
            r"^第[\d一二三四五六七八九十百]+章", text
        ):
            exemplars["heading1"] = para
        elif "heading2" not in exemplars and re.match(r"^\d+\.\d+\s+\S", text):
            exemplars["heading2"] = para
        elif "heading3" not in exemplars and re.match(r"^\d+\.\d+\.\d+\s+\S", text):
            exemplars["heading3"] = para
        elif "abstract_title" not in exemplars and re.match(
            r"^摘\s*要$|^Abstract$", text, re.I
        ):
            exemplars["abstract_title"] = para
        elif "keywords" not in exemplars and re.match(
            r"^关键词[：:]|^Keywords[：:]", text, re.I
        ):
            exemplars["keywords"] = para
        elif "body" not in exemplars and len(text) >= 15:
            exemplars["body"] = para

    return exemplars


def _copy_east_asia_font(source_run, target_run) -> None:
    src_rpr = source_run._element.rPr
    if src_rpr is None:
        return
    rfonts = src_rpr.rFonts
    if rfonts is None:
        return
    east = rfonts.get(qn("w:eastAsia"))
    if east:
        target_run._element.get_or_add_rPr()
        target_run._element.rPr.rFonts.set(qn("w:eastAsia"), east)


def clone_run_format(source_run, target_run) -> None:
    target_run.font.name = source_run.font.name
    target_run.font.size = source_run.font.size
    target_run.font.bold = source_run.font.bold
    target_run.font.italic = source_run.font.italic
    target_run.font.underline = source_run.font.underline
    try:
        if source_run.font.color and source_run.font.color.rgb:
            target_run.font.color.rgb = source_run.font.color.rgb
    except (AttributeError, TypeError):
        pass
    _copy_east_asia_font(source_run, target_run)

    src_rpr = source_run._element.rPr
    if src_rpr is not None:
        tgt_rpr = target_run._element.get_or_add_rPr()
        for child in src_rpr:
            tag = child.tag.split("}")[-1]
            if tag in ("rFonts", "sz", "szCs", "b", "i", "u", "color", "highlight"):
                existing = tgt_rpr.find(qn(f"w:{tag}"))
                if existing is not None:
                    tgt_rpr.remove(existing)
                tgt_rpr.append(deepcopy(child))


def clone_paragraph_format(source: Paragraph, target: Paragraph) -> None:
    """完整复制段落格式：样式 + 段落属性 + 字体（含中文字体）。"""
    target.style = source.style

    sp = source.paragraph_format
    tp = target.paragraph_format
    tp.alignment = sp.alignment
    tp.left_indent = sp.left_indent
    tp.right_indent = sp.right_indent
    tp.first_line_indent = sp.first_line_indent
    tp.line_spacing = sp.line_spacing
    tp.line_spacing_rule = sp.line_spacing_rule
    tp.space_before = sp.space_before
    tp.space_after = sp.space_after
    tp.keep_together = sp.keep_together
    tp.keep_with_next = sp.keep_with_next
    tp.widow_control = sp.widow_control
    tp.page_break_before = sp.page_break_before

    # 复制段落级 XML（缩进、行距等更完整）
    if source._element.pPr is not None:
        new_ppr = deepcopy(source._element.pPr)
        old = target._element.pPr
        if old is not None:
            target._element.remove(old)
        target._element.insert(0, new_ppr)

    if not source.runs or not target.runs:
        return

    src_run = source.runs[0]
    for run in target.runs:
        clone_run_format(src_run, run)


def exemplar_style_label(para: Paragraph) -> str:
    text = para.text.strip()
    preview = text[:20] + ("..." if len(text) > 20 else "")
    style = para.style.name if para.style and para.style.name else "?"
    return f"{style}「{preview}」"
