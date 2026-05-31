from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path

import yaml
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from .classifier import compile_rules
from .format_clone import clone_paragraph_format, exemplar_style_label
from .format_library import (
    build_template_format_library,
    library_summary,
    pick_format_exemplar,
    pick_format_exemplar_strict,
)
from .paragraph_detect import ArticleParagraph, classify_article_paragraph
from .style_constants import STYLE_CANDIDATES


def iter_block_items(parent):
    parent_elm = parent.element.body
    for child in parent_elm.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, parent)
        elif child.tag == qn("w:tbl"):
            yield Table(child, parent)


def extract_classified_paragraphs(source: Document, rules) -> list[ArticleParagraph]:
    """读取文章每一段，识别是一级标题/二级标题/正文等。"""
    items: list[ArticleParagraph] = []
    for block in iter_block_items(source):
        if not isinstance(block, Paragraph):
            continue
        text = block.text.strip()
        if not text:
            continue
        para_type = classify_article_paragraph(block, rules)
        style_name = block.style.name if block.style else None
        items.append(ArticleParagraph(text, para_type, style_name))
    return items


def extract_plain_paragraphs(source: Document) -> list[str]:
    texts: list[str] = []
    for block in iter_block_items(source):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            if text:
                texts.append(text)
    return texts


def list_paragraph_styles(docx_path: str | Path) -> list[str]:
    doc = Document(str(docx_path))
    names: list[str] = []
    for style in doc.styles:
        if style.type == WD_STYLE_TYPE.PARAGRAPH and style.name:
            names.append(style.name)
    return sorted(set(names))


def resolve_style_name(doc: Document, para_type: str) -> str | None:
    available = {
        style.name
        for style in doc.styles
        if style.type == WD_STYLE_TYPE.PARAGRAPH and style.name
    }
    for candidate in STYLE_CANDIDATES.get(para_type, STYLE_CANDIDATES["body"]):
        if candidate in available:
            return candidate
    if "正文" in available:
        return "正文"
    if "Normal" in available:
        return "Normal"
    return None


def clear_document_body(doc: Document) -> None:
    body = doc.element.body
    for child in list(body):
        if child.tag in (qn("w:p"), qn("w:tbl")):
            body.remove(child)


def add_paragraph_with_style(doc: Document, text: str, style_name: str | None) -> Paragraph:
    if style_name:
        try:
            return doc.add_paragraph(text, style=style_name)
        except KeyError:
            pass
    return doc.add_paragraph(text)


def load_classification_rules(config_path: str | Path | None) -> list:
    if config_path is None:
        default = Path(__file__).resolve().parent.parent / "config" / "default_cn_thesis.yaml"
        config_path = default if default.exists() else None
    if config_path and Path(config_path).exists():
        with Path(config_path).open("r", encoding="utf-8") as f:
            config = yaml.safe_load(f)
        return compile_rules(config.get("rules", []))
    return compile_rules([])


def prepare_output_from_article(
    article_path: Path, output_path: Path
) -> tuple[Path, str | None]:
    """以正文为基底复制输出文件（不引入模板里正文没有的内容）。"""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        shutil.copy2(article_path, output_path)
        return output_path, None
    except PermissionError:
        alt = output_path.with_name(
            f"{output_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{output_path.suffix}"
        )
        shutil.copy2(article_path, alt)
        notice = (
            f"无法写入 {output_path.name}：文件可能正被 Word 打开。"
            f"已改存为 {alt.name}。"
        )
        return alt, notice


def prepare_output_from_template(
    template_path: Path, output_path: Path
) -> tuple[Path, str | None]:
    """把模板复制为输出文件。若目标被 Word 占用，自动改用带时间戳的新文件名。"""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        shutil.copy2(template_path, output_path)
        return output_path, None
    except PermissionError:
        alt = output_path.with_name(
            f"{output_path.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{output_path.suffix}"
        )
        shutil.copy2(template_path, alt)
        notice = (
            f"无法写入 {output_path.name}：文件可能正被 Word 或其他程序打开。"
            f"已改存为 {alt.name}。请关闭 Word 后重新运行即可覆盖原文件。"
        )
        return alt, notice


def copy_page_setup_from_template(output_doc: Document, template_doc: Document) -> None:
    """只复制页边距/页面尺寸，不复制模板文字内容。"""
    if not template_doc.sections or not output_doc.sections:
        return
    tpl = template_doc.sections[0]
    for section in output_doc.sections:
        section.page_width = tpl.page_width
        section.page_height = tpl.page_height
        section.top_margin = tpl.top_margin
        section.bottom_margin = tpl.bottom_margin
        section.left_margin = tpl.left_margin
        section.right_margin = tpl.right_margin


def iter_content_paragraphs(doc: Document):
    for para in doc.paragraphs:
        if para.text.strip():
            yield para


def load_front_matter_config(config_path: str | Path | None) -> dict:
    if config_path is None:
        default = Path(__file__).resolve().parent.parent / "config" / "default_cn_thesis.yaml"
        config_path = default if default.exists() else None
    if config_path and Path(config_path).exists():
        with Path(config_path).open("r", encoding="utf-8") as f:
            config = yaml.safe_load(f) or {}
        return config.get("front_matter") or {}
    return {}


def find_body_start_paragraph_index(doc: Document, rules, body_start_at: str = "abstract_or_chapter") -> int:
    """在模板中定位「正文起点」：此前段落视为封面/声明等前置页并保留。"""
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        para_type = classify_article_paragraph(para, rules)
        if body_start_at in ("abstract_or_chapter", "abstract") and para_type == "abstract_title":
            return i
        if body_start_at in ("abstract_or_chapter", "chapter") and para_type == "heading1":
            return i
    return len(doc.paragraphs)


def remove_paragraphs_from_index(doc: Document, start_index: int) -> None:
    """删除从 start_index 起的所有段落（保留封面等前置内容）。"""
    for para in reversed(doc.paragraphs[start_index:]):
        element = para._element
        parent = element.getparent()
        if parent is not None:
            parent.remove(element)


def apply_selective_format_to_paragraph(
    source_para: Paragraph,
    target_para: Paragraph,
    para_type: str,
    format_library: dict,
    stats: dict,
) -> None:
    exemplar = pick_format_exemplar_strict(format_library, para_type)
    if exemplar:
        clone_paragraph_format(exemplar, target_para)
        stats["format_applied_count"] += 1
        apply_key = f"{para_type} -> {exemplar_style_label(exemplar)}"
        stats["format_applied"][apply_key] = stats["format_applied"].get(apply_key, 0) + 1
    else:
        clone_paragraph_format(source_para, target_para)
        stats["format_unchanged_count"] += 1
        preview = source_para.text.strip()[:30]
        unchanged_key = f"{para_type}: {preview}"
        stats["format_unchanged"][unchanged_key] = stats["format_unchanged"].get(unchanged_key, 0) + 1


def apply_template_format(
    article_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
    config_path: str | Path | None = None,
    keep_front_matter: bool = True,
    preserve_template_front: bool | None = None,
) -> dict:
    """v0.6：保留模板封面/前置页 + 正文选择性套格式。

    1. 默认复制**模板**为输出，保留「摘要/第1章」之前的封面、声明等
    2. 删除模板中正文部分，换入你的文章内容
    3. 对文章每段：模板有对应格式则套用，否则保持文章原格式
    4. preserve_template_front=False 时退化为 v0.5（以文章为底，无封面）
    """
    article_path = Path(article_path)
    template_path = Path(template_path)
    output_path = Path(output_path)

    if not article_path.exists():
        raise FileNotFoundError(f"找不到文章: {article_path}")
    if not template_path.exists():
        raise FileNotFoundError(f"找不到模板: {template_path}")

    if preserve_template_front is None:
        preserve_template_front = keep_front_matter

    rules = load_classification_rules(config_path)
    front_cfg = load_front_matter_config(config_path)
    if "preserve_from_template" in front_cfg:
        preserve_template_front = bool(front_cfg["preserve_from_template"])
    body_start_at = front_cfg.get("body_start_at", "abstract_or_chapter")

    article_doc = Document(str(article_path))
    template_doc = Document(str(template_path))
    format_library = build_template_format_library(template_doc, rules)
    library_info = library_summary(format_library)

    stats: dict = {
        "version": "0.6",
        "preserve_template_front": preserve_template_front,
        "total_paragraphs_in_document": 0,
        "content_paragraphs": 0,
        "format_applied_count": 0,
        "format_unchanged_count": 0,
        "by_type": {},
        "format_applied": {},
        "format_unchanged": {},
        "template_format_library": library_info,
        "notice": None,
        "template_front_preserved_paragraphs": 0,
    }

    article_paras = list(iter_content_paragraphs(article_doc))
    stats["content_paragraphs"] = len(article_paras)

    if preserve_template_front:
        output_path, output_notice = prepare_output_from_template(template_path, output_path)
        output_doc = Document(str(output_path))
        body_start = find_body_start_paragraph_index(template_doc, rules, body_start_at)
        stats["template_front_preserved_paragraphs"] = body_start
        stats["mode"] = "template_front_plus_article_body"
        remove_paragraphs_from_index(output_doc, body_start)

        for article_para in article_paras:
            para_type = classify_article_paragraph(article_para, rules)
            stats["by_type"][para_type] = stats["by_type"].get(para_type, 0) + 1
            new_para = output_doc.add_paragraph(article_para.text)
            apply_selective_format_to_paragraph(
                article_para, new_para, para_type, format_library, stats
            )

        stats["total_paragraphs_in_document"] = len(output_doc.paragraphs)
    else:
        output_path, output_notice = prepare_output_from_article(article_path, output_path)
        output_doc = Document(str(output_path))
        copy_page_setup_from_template(output_doc, template_doc)
        stats["mode"] = "selective_from_article"
        stats["total_paragraphs_in_document"] = len(article_doc.paragraphs)

        output_paras = list(iter_content_paragraphs(output_doc))
        if len(article_paras) != len(output_paras):
            raise RuntimeError(
                f"正文与输出段落数量不一致（{len(article_paras)} vs {len(output_paras)}）"
            )

        for article_para, output_para in zip(article_paras, output_paras):
            para_type = classify_article_paragraph(article_para, rules)
            stats["by_type"][para_type] = stats["by_type"].get(para_type, 0) + 1
            apply_selective_format_to_paragraph(
                article_para, output_para, para_type, format_library, stats
            )

    stats["notice"] = output_notice
    output_doc.save(str(output_path))
    stats["output"] = str(output_path.resolve())
    stats["total"] = stats["content_paragraphs"]
    stats["template_styles_available"] = list_paragraph_styles(template_path)
    return stats


def fill_template_empty_slots(
    user_docx: str | Path,
    template_docx: str | Path,
    output_docx: str | Path,
) -> int:
    user_doc = Document(str(user_docx))
    rules = load_classification_rules(None)
    items = extract_classified_paragraphs(user_doc, rules)

    template_docx = Path(template_docx)
    output_docx = Path(output_docx)
    doc = Document(str(template_docx))
    empty_paragraphs = [p for p in doc.paragraphs if not p.text.strip()]
    format_library = build_template_format_library(doc, rules)

    idx = 0
    for item in items:
        exemplar = pick_format_exemplar_strict(format_library, item.para_type)
        if idx < len(empty_paragraphs):
            empty_paragraphs[idx].text = item.text
            if exemplar:
                clone_paragraph_format(exemplar, empty_paragraphs[idx])
            idx += 1
        else:
            if exemplar:
                style_name = exemplar.style.name if exemplar.style else None
                paragraph = add_paragraph_with_style(doc, item.text, style_name)
                clone_paragraph_format(exemplar, paragraph)
            else:
                doc.add_paragraph(item.text)

    output_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_docx))
    return len(items)


def merge_user_doc_into_template(
    user_docx: str | Path,
    template_docx: str | Path,
    output_docx: str | Path,
    mode: str = "apply",
    config_path: str | Path | None = None,
) -> dict | int:
    if mode == "fill":
        return fill_template_empty_slots(user_docx, template_docx, output_docx)
    return apply_template_format(user_docx, template_docx, output_docx, config_path)
