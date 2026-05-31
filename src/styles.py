from __future__ import annotations

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


ALIGNMENT_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def cm_to_twips(cm: float) -> int:
    return int(cm * 567)


def apply_page_setup(section, page_cfg: dict) -> None:
    section.page_width = Cm(page_cfg["width_cm"])
    section.page_height = Cm(page_cfg["height_cm"])
    section.top_margin = Cm(page_cfg["margin_top_cm"])
    section.bottom_margin = Cm(page_cfg["margin_bottom_cm"])
    section.left_margin = Cm(page_cfg["margin_left_cm"])
    section.right_margin = Cm(page_cfg["margin_right_cm"])


def set_run_font(run, font_cn: str, font_en: str, size_pt: float, bold: bool = False) -> None:
    run.font.name = font_en
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font_cn)


def apply_paragraph_style(paragraph, style_cfg: dict) -> None:
    alignment = style_cfg.get("alignment")
    if alignment:
        paragraph.alignment = ALIGNMENT_MAP.get(alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)

    pf = paragraph.paragraph_format
    pf.space_before = Pt(style_cfg.get("space_before_pt", 0))
    pf.space_after = Pt(style_cfg.get("space_after_pt", 0))

    if style_cfg.get("page_break_before"):
        pf.page_break_before = True

    line_spacing = style_cfg.get("line_spacing")
    if line_spacing:
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = line_spacing

    first_line_indent_cm = style_cfg.get("first_line_indent_cm")
    if first_line_indent_cm is not None:
        pf.first_line_indent = Cm(first_line_indent_cm)
    elif style_cfg.get("clear_first_line_indent"):
        pf.first_line_indent = Cm(0)

    font_cn = style_cfg["font_cn"]
    font_en = style_cfg.get("font_en", "Times New Roman")
    size_pt = style_cfg["size_pt"]
    bold = style_cfg.get("bold", False)

    if not paragraph.runs:
        paragraph.add_run(paragraph.text)

    for run in paragraph.runs:
        set_run_font(run, font_cn, font_en, size_pt, bold)
