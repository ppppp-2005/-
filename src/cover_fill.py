from __future__ import annotations

import re

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

# 字段同义词：用于在「你的论文封面」里识别字段标签
FIELD_SYNONYMS: dict[str, list[str]] = {
    "学院": ["学院"],
    "专业": ["专业班级", "专业方向", "专业"],
    "姓名": ["学生姓名", "作者姓名", "姓名"],
    "学号": ["学号"],
    "指导教师": ["指导教师", "指导老师", "导师"],
    "日期": ["完成日期", "答辩日期", "日期"],
}


def _compact(text: str | None) -> str:
    return re.sub(r"\s+", "", text or "")


def has_image(para: Paragraph) -> bool:
    el = para._element
    return bool(
        el.findall(".//" + qn("w:drawing")) or el.findall(".//" + qn("w:pict"))
    )


def detect_field(text: str) -> str | None:
    """根据段落文字判断它属于哪个封面字段。"""
    t = _compact(text)
    if "指导教师" in t or "指导老师" in t or t.startswith("导师"):
        return "指导教师"
    if "学院" in t:
        return "学院"
    if "专业" in t:
        return "专业"
    if "姓名" in t:
        return "姓名"
    if "学号" in t:
        return "学号"
    if "答辩日期" in t or "完成日期" in t:
        return "日期"
    return None


def _extract_title(text: str) -> str:
    """从「题目（中文）：xxx（英文）：yyy」里取中文题目。"""
    m = re.search(r"题目[^：:]*[：:]\s*(.+)", text)
    rest = m.group(1) if m else text
    rest = re.split(r"（英文）|\(英文\)|（英文|\(英文", rest)[0]
    return rest.strip()


def _extract_english_title_start(text: str) -> str:
    """从「（英文）：xxx」里取英文题名的起始部分。"""
    m = re.search(r"[（(]\s*英文\s*[）)]?\s*[：:]\s*(.+)", text)
    return m.group(1).strip() if m else ""


def _looks_like_english_title_continuation(text: str) -> bool:
    if not text or detect_field(text):
        return False
    if "声明" in text or "日期" in text or "学院" in text:
        return False
    return bool(re.search(r"[A-Za-z]", text)) and not bool(re.search(r"[\u4e00-\u9fff]", text))


def _extract_value(text: str, field: str) -> str:
    compact = _compact(text)
    for syn in FIELD_SYNONYMS.get(field, [field]):
        idx = compact.find(syn)
        if idx != -1:
            return compact[idx + len(syn) :].strip("：: ").strip()
    return ""


def extract_article_cover_fields(cover_paragraphs: list[Paragraph]) -> dict[str, str]:
    """从你论文封面段落里提取字段值（学院/专业/姓名/学号/指导教师/日期/题目）。"""
    fields: dict[str, str] = {}
    english_title_parts: list[str] = []
    collecting_english_title = False

    for para in cover_paragraphs:
        text = para.text.strip()
        if not text:
            continue
        if "题目" in _compact(text) and "题目" not in fields:
            title = _extract_title(text)
            if title:
                fields["题目"] = title
            english_title = _extract_english_title_start(text)
            if english_title:
                english_title_parts.append(english_title)
                collecting_english_title = True
            continue
        if collecting_english_title:
            if _looks_like_english_title_continuation(text):
                english_title_parts.append(text)
                continue
            collecting_english_title = False
        field = detect_field(text)
        if field and field not in fields:
            value = _extract_value(text, field)
            if value:
                fields[field] = value

    if english_title_parts:
        fields["英文题目"] = re.sub(r"\s+", " ", " ".join(english_title_parts)).strip()
    return fields


def _split_label(text: str) -> tuple[str, str]:
    """把「学    院：xxx」拆成 (标签含冒号, 冒号后内容)。"""
    for sep in ("：", ":"):
        if sep in text:
            i = text.index(sep)
            return text[: i + 1], text[i + 1 :]
    return text, ""


def _set_text_keep_format(para: Paragraph, new_text: str) -> None:
    """改段落文字但保留段落/字体格式（用第一个 run 承载文字）。"""
    if para.runs:
        para.runs[0].text = new_text
        for run in para.runs[1:]:
            run.text = ""
    else:
        para.add_run(new_text)


def fill_template_cover(
    template_cover_paragraphs: list[Paragraph], fields: dict[str, str]
) -> dict[str, str]:
    """把字段值填入模板封面占位；含图片的段落跳过不动。"""
    filled: dict[str, str] = {}
    min_field_pos: int | None = None

    for pos, para in enumerate(template_cover_paragraphs):
        if has_image(para):
            continue
        text = para.text.strip()
        if not text:
            continue
        field = detect_field(text)
        if not field:
            continue
        # 占位字段应是短段落（标签+少量空格）；声明等长正文不是字段，跳过
        if len(_compact(text)) > 16:
            continue
        if min_field_pos is None:
            min_field_pos = pos
        if field in fields and ("：" in para.text or ":" in para.text):
            label, _ = _split_label(para.text)
            _set_text_keep_format(para, label + fields[field])
            filled[field] = fields[field]

    # 题目：在第一个字段占位之前、最后一个非图片的长文本段
    if "题目" in fields and min_field_pos is not None:
        for pos in range(min_field_pos - 1, -1, -1):
            para = template_cover_paragraphs[pos]
            if has_image(para):
                continue
            text = para.text.strip()
            if len(text) >= 8 and detect_field(text) is None and "声明" not in text:
                _set_text_keep_format(para, fields["题目"])
                filled["题目"] = fields["题目"]
                break

    return filled
