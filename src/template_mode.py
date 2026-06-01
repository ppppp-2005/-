from __future__ import annotations

import shutil
import re
from copy import deepcopy
from datetime import datetime
from pathlib import Path

import yaml
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from .classifier import compile_rules
from .cover_fill import extract_article_cover_fields, fill_template_cover
from .format_clone import clone_paragraph_format, exemplar_style_label
from .format_library import (
    build_template_format_library,
    library_summary,
    pick_format_exemplar,
    pick_format_exemplar_strict,
)
from .paragraph_detect import ArticleParagraph, classify_article_paragraph
from .style_constants import STYLE_CANDIDATES
from .template_profile import (
    build_template_style_profile,
    merge_profile_with_exemplars,
    resolve_style_with_profile,
    styles_used_in_template,
)


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


def remove_body_after_paragraph(doc: Document, start_para_index: int) -> None:
    """删除从第 start_para_index 个段落起的所有段落和表格（保留封面、sectPr）。"""
    paras = doc.paragraphs
    if start_para_index >= len(paras):
        return
    start_el = paras[start_para_index]._element
    body = doc.element.body
    removing = False
    for child in list(body):
        if child is start_el:
            removing = True
        if removing and child.tag in (qn("w:p"), qn("w:tbl")):
            body.remove(child)


def iter_blocks_from(doc: Document, start_para_index: int):
    """从第 start_para_index 个段落对应的块开始，按顺序产出段落和表格。"""
    paras = doc.paragraphs
    start_el = paras[start_para_index]._element if start_para_index < len(paras) else None
    body = doc.element.body
    started = start_el is None
    for child in body.iterchildren():
        if child is start_el:
            started = True
        if not started:
            continue
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


def append_block_element(doc: Document, element) -> None:
    """把一个块元素（如表格）追加到正文末尾、sectPr 之前。"""
    body = doc.element.body
    sect_pr = body.find(qn("w:sectPr"))
    if sect_pr is not None:
        sect_pr.addprevious(element)
    else:
        body.append(element)


def set_paragraph_text_keep_format(para: Paragraph, text: str) -> None:
    """替换段落文字，同时尽量保留原段落和首个 run 的格式。"""
    if para.runs:
        para.runs[0].text = text
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.add_run(text)


def append_template_paragraph_with_text(
    doc: Document, template_para: Paragraph, text: str
) -> Paragraph:
    new_element = deepcopy(template_para._element)
    append_block_element(doc, new_element)
    paragraph = Paragraph(new_element, doc)
    set_paragraph_text_keep_format(paragraph, text)
    return paragraph


def is_chinese_abstract_text(text: str) -> bool:
    return bool(re.match(r"^摘\s*要\s*(?:$|[：:])", text.strip(), re.I))


def is_english_abstract_text(text: str) -> bool:
    return bool(re.match(r"^Abstract\s*(?:$|[：:])", text.strip(), re.I))


def is_chinese_abstract_title_text(text: str) -> bool:
    return bool(re.match(r"^摘\s*要$", text.strip(), re.I))


def is_english_abstract_title_text(text: str) -> bool:
    return bool(re.match(r"^Abstract$", text.strip(), re.I))


def is_keywords_text(text: str) -> bool:
    return bool(re.match(r"^(关键词|关键字|Keywords|Key\s+words)\s*[：:]", text.strip(), re.I))


def abstract_inline_prefix(template_para: Paragraph | None) -> str | None:
    if template_para is None:
        return None
    text = template_para.text.strip()
    match = re.match(r"^(摘\s*要\s*[：:]|Abstract\s*[：:])\s*", text, re.I)
    return match.group(1) if match else None


def find_first_index(paras: list[Paragraph], start: int, predicate) -> int | None:
    for index in range(start, len(paras)):
        text = paras[index].text.strip()
        if text and predicate(text):
            return index
    return None


def build_abstract_page_preface(
    template_doc: Document, body_start: int
) -> dict[str, object]:
    """提取模板摘要页里的题名/单位题头，用于换成用户论文信息后再插入。"""
    paras = template_doc.paragraphs
    cn_abs = find_first_index(paras, body_start, is_chinese_abstract_text)
    if cn_abs is None:
        return {
            "cn_header": [],
            "en_header": [],
            "cn_abstract": None,
            "en_abstract": None,
        }

    en_abs = find_first_index(paras, cn_abs + 1, is_english_abstract_text)
    cn_keyword = find_first_index(paras, cn_abs + 1, is_keywords_text)
    if en_abs is not None and cn_keyword is not None and cn_keyword < en_abs:
        en_header = paras[cn_keyword + 1 : en_abs]
    else:
        en_header = []

    return {
        "cn_header": paras[body_start:cn_abs],
        "en_header": en_header,
        "cn_abstract": paras[cn_abs],
        "en_abstract": paras[en_abs] if en_abs is not None else None,
    }


def append_replaced_header_block(
    doc: Document,
    template_paras: list[Paragraph],
    fields: dict[str, str],
    language: str,
) -> int:
    title = fields.get("题目", "")
    affiliation = fields.get("学院", "")
    if language == "en":
        title = fields.get("英文题目") or title
        affiliation = fields.get("英文学院") or affiliation

    text_seen = 0
    inserted = 0
    for template_para in template_paras:
        original = template_para.text.strip()
        replacement = ""
        if original:
            if text_seen == 0:
                replacement = title
            elif text_seen == 1:
                replacement = affiliation
            text_seen += 1
        append_template_paragraph_with_text(doc, template_para, replacement)
        inserted += 1
    return inserted


def try_set_paragraph_style(para: Paragraph, style_name: str) -> bool:
    try:
        para.style = style_name
        return True
    except KeyError:
        return False


def available_paragraph_styles(doc: Document) -> set[str]:
    return {
        style.name
        for style in doc.styles
        if style.type == WD_STYLE_TYPE.PARAGRAPH and style.name
    }


def apply_selective_format_to_paragraph(
    source_para: Paragraph,
    target_para: Paragraph,
    para_type: str,
    format_library: dict,
    style_profile: dict[str, str | None],
    available_styles: set[str],
    used_in_template: set[str],
    stats: dict,
) -> None:
    """v0.7：优先克隆模板范例段落；若无范例则套用模板 Word 样式名。"""
    exemplar = pick_format_exemplar_strict(format_library, para_type)
    style_name = resolve_style_with_profile(
        para_type, style_profile, available_styles, used_in_template
    )

    if exemplar:
        clone_paragraph_format(exemplar, target_para)
        stats["format_applied_count"] += 1
        apply_key = f"{para_type} -> {exemplar_style_label(exemplar)}"
        stats["format_applied"][apply_key] = stats["format_applied"].get(apply_key, 0) + 1
    elif style_name and try_set_paragraph_style(target_para, style_name):
        stats["format_applied_count"] += 1
        apply_key = f"{para_type} -> style:{style_name}"
        stats["format_applied"][apply_key] = stats["format_applied"].get(apply_key, 0) + 1
    else:
        clone_paragraph_format(source_para, target_para)
        stats["format_unchanged_count"] += 1
        preview = source_para.text.strip()[:30]
        unchanged_key = f"{para_type}: {preview}"
        stats["format_unchanged"][unchanged_key] = stats["format_unchanged"].get(unchanged_key, 0) + 1


def apply_exemplar_format(
    exemplar: Paragraph, target_para: Paragraph, para_type: str, stats: dict
) -> None:
    clone_paragraph_format(exemplar, target_para)
    stats["format_applied_count"] += 1
    apply_key = f"{para_type} -> {exemplar_style_label(exemplar)}"
    stats["format_applied"][apply_key] = stats["format_applied"].get(apply_key, 0) + 1


def apply_template_format(
    article_path: str | Path,
    template_path: str | Path,
    output_path: str | Path,
    config_path: str | Path | None = None,
    keep_front_matter: bool = True,
    preserve_template_front: bool | None = None,
) -> dict:
    """v0.7：保留模板封面/前置页 + 正文填入模板格式。

    1. 默认复制**模板**为输出，保留「摘要/第1章」之前的封面、声明等
    2. 删除模板中正文部分，换入你的文章内容
    3. 对文章每段：先克隆模板范例格式，若无范例则套用模板 Word 样式名
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

    # 封面来源：article=保留你自己的封面（默认/推荐）；template=用模板封面；none=不要封面
    cover_source = front_cfg.get("cover_source")
    if cover_source is None:
        cover_source = "template" if preserve_template_front else "article"

    article_doc = Document(str(article_path))
    template_doc = Document(str(template_path))
    # 只从模板正文区取格式范例，避免封面（居中大标题等）污染正文格式
    template_body_start = find_body_start_paragraph_index(
        template_doc, rules, body_start_at
    )
    format_library = build_template_format_library(
        template_doc, rules, start_index=template_body_start
    )
    library_info = library_summary(format_library)
    style_profile = merge_profile_with_exemplars(
        build_template_style_profile(template_doc, start_index=template_body_start),
        format_library,
    )
    used_in_template = styles_used_in_template(template_doc)

    stats: dict = {
        "version": "0.7",
        "preserve_template_front": preserve_template_front,
        "total_paragraphs_in_document": 0,
        "content_paragraphs": 0,
        "format_applied_count": 0,
        "format_unchanged_count": 0,
        "by_type": {},
        "format_applied": {},
        "format_unchanged": {},
        "template_format_library": library_info,
        "template_style_profile": {k: v for k, v in style_profile.items() if v},
        "notice": None,
        "template_front_preserved_paragraphs": 0,
    }

    article_paras = list(iter_content_paragraphs(article_doc))
    stats["content_paragraphs"] = len(article_paras)

    if cover_source == "article":
        # 推荐：以你的论文为底，保留你自己的封面/声明页，只对正文套模板格式
        output_path, output_notice = prepare_output_from_article(article_path, output_path)
        output_doc = Document(str(output_path))
        copy_page_setup_from_template(output_doc, template_doc)
        stats["mode"] = "article_cover_plus_template_body_format"
        stats["cover_source"] = "article"

        body_start = find_body_start_paragraph_index(output_doc, rules, body_start_at)
        available_styles = available_paragraph_styles(output_doc)

        preserved = 0
        formatted = 0
        all_paras = output_doc.paragraphs
        for i, para in enumerate(all_paras):
            if i < body_start:
                if para.text.strip():
                    preserved += 1
                continue
            if not para.text.strip():
                continue
            para_type = classify_article_paragraph(para, rules)
            stats["by_type"][para_type] = stats["by_type"].get(para_type, 0) + 1
            apply_selective_format_to_paragraph(
                para,
                para,
                para_type,
                format_library,
                style_profile,
                available_styles,
                used_in_template,
                stats,
            )
            formatted += 1

        stats["template_front_preserved_paragraphs"] = preserved
        stats["content_paragraphs"] = formatted
        stats["total_paragraphs_in_document"] = len(all_paras)
    elif preserve_template_front or cover_source == "template":
        output_path, output_notice = prepare_output_from_template(template_path, output_path)
        output_doc = Document(str(output_path))
        body_start = template_body_start
        stats["template_front_preserved_paragraphs"] = body_start
        stats["mode"] = "template_front_plus_article_body"
        stats["cover_source"] = "template"
        # 把你论文封面里的字段（姓名/学号/学院/题目…）填进模板封面占位
        article_body_start = find_body_start_paragraph_index(
            article_doc, rules, body_start_at
        )
        article_cover_paras = article_doc.paragraphs[:article_body_start]
        cover_fields = extract_article_cover_fields(article_cover_paras)
        template_cover_paras = output_doc.paragraphs[:body_start]
        cover_filled = fill_template_cover(template_cover_paras, cover_fields)
        stats["cover_fields_extracted"] = cover_fields
        stats["cover_fields_filled"] = cover_filled

        # 删除模板正文区的段落和表格（避免残留模板里的示例表格）
        remove_body_after_paragraph(output_doc, body_start)
        available_styles = available_paragraph_styles(output_doc)
        abstract_preface = build_abstract_page_preface(template_doc, body_start)
        cn_header = abstract_preface.get("cn_header", [])
        en_header = abstract_preface.get("en_header", [])
        cn_abstract_exemplar = abstract_preface.get("cn_abstract")
        en_abstract_exemplar = abstract_preface.get("en_abstract")
        cn_inline_prefix = abstract_inline_prefix(cn_abstract_exemplar)
        en_inline_prefix = abstract_inline_prefix(en_abstract_exemplar)
        stats["abstract_page_preface"] = {
            "cn_header_paragraphs": len(cn_header),
            "en_header_paragraphs": len(en_header),
            "cn_inline_abstract": bool(cn_inline_prefix),
            "en_inline_abstract": bool(en_inline_prefix),
        }

        # 按原顺序接入你论文的正文（段落 + 表格），跳过你自己的封面前置页
        table_count = 0
        para_count = 0
        cn_header_inserted = False
        en_header_inserted = False
        waiting_for_cn_abstract_body = False
        waiting_for_en_abstract_body = False
        for block in iter_blocks_from(article_doc, article_body_start):
            if isinstance(block, Table):
                append_block_element(output_doc, deepcopy(block._element))
                table_count += 1
                continue
            if not block.text.strip():
                continue
            para_type = classify_article_paragraph(block, rules)
            stats["by_type"][para_type] = stats["by_type"].get(para_type, 0) + 1

            if not cn_header_inserted and cn_header:
                inserted = append_replaced_header_block(
                    output_doc, cn_header, cover_fields, "cn"
                )
                stats["abstract_page_preface"]["cn_header_inserted"] = inserted
                cn_header_inserted = True

            if para_type == "abstract_title" and is_chinese_abstract_title_text(block.text):
                if cn_inline_prefix:
                    waiting_for_cn_abstract_body = True
                    continue

            if para_type == "abstract_title" and is_english_abstract_title_text(block.text):
                if not en_header_inserted and en_header:
                    inserted = append_replaced_header_block(
                        output_doc, en_header, cover_fields, "en"
                    )
                    stats["abstract_page_preface"]["en_header_inserted"] = inserted
                    en_header_inserted = True
                if en_inline_prefix:
                    waiting_for_en_abstract_body = True
                    continue

            if waiting_for_cn_abstract_body and para_type == "body" and cn_inline_prefix:
                new_para = output_doc.add_paragraph(cn_inline_prefix + block.text.lstrip())
                if isinstance(cn_abstract_exemplar, Paragraph):
                    apply_exemplar_format(
                        cn_abstract_exemplar, new_para, "abstract_body", stats
                    )
                else:
                    apply_selective_format_to_paragraph(
                        block,
                        new_para,
                        para_type,
                        format_library,
                        style_profile,
                        available_styles,
                        used_in_template,
                        stats,
                    )
                waiting_for_cn_abstract_body = False
                para_count += 1
                continue

            if waiting_for_en_abstract_body and para_type == "body" and en_inline_prefix:
                new_para = output_doc.add_paragraph(en_inline_prefix + block.text.lstrip())
                if isinstance(en_abstract_exemplar, Paragraph):
                    apply_exemplar_format(
                        en_abstract_exemplar, new_para, "abstract_body", stats
                    )
                else:
                    apply_selective_format_to_paragraph(
                        block,
                        new_para,
                        para_type,
                        format_library,
                        style_profile,
                        available_styles,
                        used_in_template,
                        stats,
                    )
                waiting_for_en_abstract_body = False
                para_count += 1
                continue

            new_para = output_doc.add_paragraph(block.text)
            apply_selective_format_to_paragraph(
                block,
                new_para,
                para_type,
                format_library,
                style_profile,
                available_styles,
                used_in_template,
                stats,
            )
            para_count += 1

        stats["article_front_skipped_paragraphs"] = article_body_start
        stats["content_paragraphs"] = para_count
        stats["tables_copied_from_article"] = table_count
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

        available_styles = available_paragraph_styles(output_doc)

        for article_para, output_para in zip(article_paras, output_paras):
            para_type = classify_article_paragraph(article_para, rules)
            stats["by_type"][para_type] = stats["by_type"].get(para_type, 0) + 1
            apply_selective_format_to_paragraph(
                article_para,
                output_para,
                para_type,
                format_library,
                style_profile,
                available_styles,
                used_in_template,
                stats,
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
    style_profile = merge_profile_with_exemplars(
        build_template_style_profile(doc), format_library
    )
    used_in_template = styles_used_in_template(doc)
    available_styles = available_paragraph_styles(doc)

    idx = 0
    for item in items:
        exemplar = pick_format_exemplar_strict(format_library, item.para_type)
        style_name = resolve_style_with_profile(
            item.para_type, style_profile, available_styles, used_in_template
        )
        if idx < len(empty_paragraphs):
            empty_paragraphs[idx].text = item.text
            if exemplar:
                clone_paragraph_format(exemplar, empty_paragraphs[idx])
            elif style_name:
                try_set_paragraph_style(empty_paragraphs[idx], style_name)
            idx += 1
        else:
            if exemplar:
                paragraph = add_paragraph_with_style(
                    doc, item.text, exemplar.style.name if exemplar.style else None
                )
                clone_paragraph_format(exemplar, paragraph)
            elif style_name:
                add_paragraph_with_style(doc, item.text, style_name)
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
