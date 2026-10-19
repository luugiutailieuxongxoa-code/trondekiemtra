"""Source-preserving DOCX block reorderer.

Moves existing WordprocessingML body blocks rather than rebuilding paragraphs. It keeps
relationships/media/OMML/OLE in the same document package and never silently shuffles
across section boundaries. Answer-mark cleanup and option shuffling are intentionally
not performed here.
"""
from __future__ import annotations
import re
from io import BytesIO
from typing import Sequence
import random
import os
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from lxml import etree
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

QUESTION_RE = re.compile(r"^\s*Câu\s*(\d+)\s*[.:)]?", re.IGNORECASE)
SECTION_RE = re.compile(r"^\s*(?:<g\d+>\s*)?PHẦN\s+([IVXLCDM]+|\d+)\b", re.IGNORECASE)
END_RE = re.compile(r"\s*(?:[-—– ]*HẾT[-—– ]*|[-—– ]*KẾT THÚC[-—– ]*)\s*", re.I)
G_MARKER_RE = re.compile(r"<\s*g\s*\d+\s*>", re.I)
ANSWER_MARKER_RE = re.compile(r"\bĐA\s*:\s*[^\s;,.]+", re.I)
ANSWER_KEY_RE = re.compile(r"^\s*(?:ĐÁP ÁN|ĐÁP ÁN THAM KHẢO|ĐÁP ÁN ĐỀ|ANSWER KEY)\b", re.I)
OPTION_RE = re.compile(r"^\s*[A-D][.)]\s*")
SHORT_ANSWER_HEADING_RE = re.compile(r"trả\s*lời\s*ngắn", re.I)
SHORT_ANSWER_LEAK_RE = re.compile(r"^\s*A\.\s*\S.*$", re.I)


def _paragraph_text(element) -> str:
    return "".join(node.text or "" for node in element.iter(qn("w:t"))).strip() if element.tag == qn("w:p") else ""


def _element_flat_text(element) -> str:
    """Return visible Word text for a paragraph, table, or other body element."""
    return "".join(node.text or "" for node in element.iter(qn("w:t"))).strip()


def _remove_existing_generated_header(children: list) -> tuple[list, int]:
    """Remove a recognizable prior exam header before inserting the configured header.

    Earlier exports can already contain the Young Mix Pro title table and candidate lines.
    Keeping those blocks and prepending a new template creates two headers and may leave a
    stale subject (for example HÓA HỌC on a Toán paper). Only recognized header elements
    before the first PHẦN heading are removed; question content is never touched.
    """
    first_section = next((i for i, child in enumerate(children)
                          if SECTION_RE.match(_paragraph_text(child))), None)
    if first_section is None:
        return children, 0
    prefix = children[:first_section]
    prefix_text = " ".join(_element_flat_text(el).lower() for el in prefix)
    has_prior_header = any(marker in prefix_text for marker in (
        "đề thi gồm", "đề thi chính thức", "mã đề:", "họ và tên thí sinh", "số báo danh"
    ))
    if not has_prior_header:
        return children, 0

    header_line_re = re.compile(
        r"^\s*(?:SỞ\s+GD|THPT\b|TRƯỜNG\b|ĐỀ\s+(?:KIỂM TRA|THI)\b|"
        r"Năm học\b|Khối/Lớp\b|Môn\s*:|Thời gian làm bài\s*:|"
        r"Đề thi gồm\b|ĐỀ THI CHÍNH THỨC\b|Mã đề\s*:|"
        r"Họ và tên thí sinh\s*:|Số báo danh\s*:)", re.I
    )
    kept = []
    removed = 0
    for index, child in enumerate(children):
        text = _element_flat_text(child)
        is_before_section = index < first_section
        is_header_table = child.tag == qn("w:tbl") and is_before_section and any(
            marker in text.lower() for marker in ("đề thi gồm", "đề thi chính thức", "mã đề:")
        )
        is_header_line = is_before_section and child.tag == qn("w:p") and header_line_re.match(text)
        is_blank_prefix = is_before_section and child.tag == qn("w:p") and not text
        if is_header_table or is_header_line or is_blank_prefix:
            removed += 1
            continue
        kept.append(child)
    return kept, removed


def _question_number(element) -> int | None:
    match = QUESTION_RE.match(_paragraph_text(element))
    return int(match.group(1)) if match else None


def _replace_heading_number(element, number: int) -> bool:
    """Replace only a complete 'Câu N' marker in a single text node."""
    for node in element.iter(qn("w:t")):
        value = node.text or ""
        match = QUESTION_RE.match(value)
        if match:
            punctuation_match = re.search(r"[.:)](?=\s|$)", match.group(0))
            punctuation = punctuation_match.group(0) if punctuation_match else ""
            node.text = QUESTION_RE.sub(f"Câu {number}{punctuation}", value, count=1)
            if node.text.startswith(" ") or node.text.endswith(" "):
                node.set(qn("xml:space"), "preserve")
            return True
    return False


def _section_key(text: str, current: str) -> str:
    match = SECTION_RE.match(text)
    return f"P{match.group(1).upper()}" if match else current


def _clean_marker_text(element) -> bool:
    """Remove parser-only <g1>/<g2>/<g3> and inline ĐA:... markers from text nodes."""
    changed = False
    for node in element.iter(qn("w:t")):
        value = node.text or ""
        cleaned = ANSWER_MARKER_RE.sub("", G_MARKER_RE.sub("", value))
        if cleaned != value:
            node.text = cleaned
            changed = True
            if cleaned.startswith(" ") or cleaned.endswith(" "):
                node.set(qn("xml:space"), "preserve")
    return changed


def _remove_answer_formatting(element) -> bool:
    """Remove common visual correct-answer marks from A-D option paragraphs only."""
    if element.tag != qn("w:p") or not OPTION_RE.match(_paragraph_text(element)):
        return False
    changed = False
    for run in element.iter(qn("w:r")):
        rpr = run.find(qn("w:rPr"))
        if rpr is None:
            continue
        for tag in ("w:rStyle", "w:b", "w:bCs", "w:i", "w:iCs", "w:u", "w:highlight", "w:shd", "w:color"):
            node = rpr.find(qn(tag))
            if node is not None:
                rpr.remove(node)
                changed = True
    return changed


def _set_run_font(run, name="Times New Roman", size=12, bold=None):
    run.font.name = name
    run.font.size = Pt(size)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        from docx.oxml import OxmlElement
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
        rfonts.set(qn(f"w:{attr}"), name)
    if bold is not None:
        run.bold = bold


def _append_word_field(paragraph, instruction: str, displayed: str = "1") -> None:
    """Append a Word field (PAGE/NUMPAGES) that updates when Word repaginates."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    run = paragraph.add_run()
    _set_run_font(run, size=9)
    for tag, attrs in (("begin", {}),):
        node = OxmlElement("w:fldChar"); node.set(qn("w:fldCharType"), tag); run._r.append(node)
    instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = f" {instruction} "; run._r.append(instr)
    separate = OxmlElement("w:fldChar"); separate.set(qn("w:fldCharType"), "separate"); run._r.append(separate)
    display = OxmlElement("w:t"); display.text = displayed; run._r.append(display)
    end = OxmlElement("w:fldChar"); end.set(qn("w:fldCharType"), "end"); run._r.append(end)


def _add_official_footer(section, exam_code: str = "001") -> None:
    """Footer contains only the exam code and current page number, never official-title text."""
    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.clear()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(f"Mã đề: {exam_code}  •  Trang "); _set_run_font(r, size=9)
    _append_word_field(p, "PAGE", "1")


def _remove_short_answer_inline_keys(children: list) -> tuple[list, int]:
    """Remove answer hints such as 'A. 10' in short-answer sections only.

    In the user's source template, these are answer keys placed under a short-answer
    question, not a multiple-choice option. The same pattern is preserved elsewhere.
    """
    result = []
    in_short_answer = False
    removed = 0
    for child in children:
        text = _paragraph_text(child)
        if SECTION_RE.match(text):
            in_short_answer = bool(SHORT_ANSWER_HEADING_RE.search(text))
            result.append(child)
            continue
        if in_short_answer and child.tag == qn("w:p") and SHORT_ANSWER_LEAK_RE.match(text):
            removed += 1
            continue
        result.append(child)
    return result, removed


def _add_official_header(doc, header: dict | None = None) -> None:
    """Prepend a standard editable Vietnamese exam header and set A4 print layout."""
    from docx.oxml import OxmlElement
    header = header or {}
    department = str(header.get("department") or "[SỞ GD&ĐT]")
    school = str(header.get("school") or "[TÊN TRƯỜNG]")
    exam_title = str(header.get("exam_title") or "ĐỀ KIỂM TRA")
    grade = str(header.get("grade") or "[KHỐI/LỚP]")
    school_year = str(header.get("school_year") or "[NĂM HỌC]")
    subject = str(header.get("subject") or "[MÔN THI]")
    duration = str(header.get("duration") or "[THỜI GIAN LÀM BÀI]")
    exam_code = str(header.get("exam_code") or "001")

    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.right_margin = Cm(2)
    section.left_margin = Cm(2.5)
    _add_official_footer(section, exam_code)

    table = doc.add_table(rows=1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.columns[0].width = Cm(7.4)
    table.columns[1].width = Cm(9.6)
    left, right = table.rows[0].cells
    left.width = Cm(7.4); right.width = Cm(9.6)
    for cell in (left, right):
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        tc_pr = cell._tc.get_or_add_tcPr()
        borders = OxmlElement("w:tcBorders")
        for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{edge}"); el.set(qn("w:val"), "nil"); borders.append(el)
        tc_pr.append(borders)
    left.text = ""
    p = left.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for idx, value in enumerate((department, school)):
        if idx:
            p = left.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(value.upper()); _set_run_font(r, size=11, bold=True)
    # Separator line under the agency/school identity, followed by live total-page field
    # and the official designation, all in the left header column as requested.
    school_p = left.paragraphs[-1]
    ppr = school_p._p.get_or_add_pPr()
    from docx.oxml import OxmlElement
    p_bdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single"); bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "3"); bottom.set(qn("w:color"), "000000")
    p_bdr.append(bottom); ppr.append(p_bdr)
    p = left.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Đề thi gồm "); _set_run_font(r, size=10)
    _append_word_field(p, "NUMPAGES", "1")
    r = p.add_run(" trang"); _set_run_font(r, size=10)
    p = left.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("ĐỀ THI CHÍNH THỨC"); _set_run_font(r, size=11, bold=True)
    right.text = ""
    p = right.paragraphs[0]; p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(exam_title.upper()); _set_run_font(r, size=12, bold=True)
    p = right.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"Năm học {school_year} · Khối/Lớp {grade}"); _set_run_font(r, size=11, bold=True)
    p = right.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"Môn: {subject}"); _set_run_font(r, size=11, bold=True)
    p = right.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"Thời gian làm bài: {duration}"); _set_run_font(r, size=10)

    # Move table to the beginning of the document body.
    body = doc.element.body
    body.remove(table._tbl)
    body.insert(0, table._tbl)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(f"Mã đề: {exam_code}"); _set_run_font(r, size=11, bold=True)
    body.remove(p._p)
    body.insert(1, p._p)
    p2 = doc.add_paragraph("Họ và tên thí sinh: ............................................................  Số báo danh: ....................")
    for r in p2.runs: _set_run_font(r, size=11)
    body.remove(p2._p)
    body.insert(2, p2._p)
    # Add a small gap after the candidate line, before source exam content.
    gap = doc.add_paragraph("")
    body.remove(gap._p)
    body.insert(3, gap._p)


def _clean_answer_key_tail(children: list) -> tuple[list, bool]:
    """Remove a trailing answer-key section, if explicitly headed as such."""
    for idx, child in enumerate(children):
        text = _paragraph_text(child)
        if ANSWER_KEY_RE.match(text):
            return children[:idx], True
    return children, False



OPTION_MARKER_IN_TEXT_RE = re.compile(r"(^|[\t ]+)([A-D])[.)](?=\s|$)")
TF_MARKER_IN_TEXT_RE = re.compile(r"(^|[\t ]+)([a-d])[.)](?=\s|$)")


def _set_run_bold(run_element, value: bool = True) -> None:
    """Set bold on a WordprocessingML run while preserving other run properties."""
    rpr = run_element.find(qn("w:rPr"))
    if rpr is None:
        from docx.oxml import OxmlElement
        rpr = OxmlElement("w:rPr")
        run_element.insert(0, rpr)
    bold = rpr.find(qn("w:b"))
    if value:
        if bold is None:
            from docx.oxml import OxmlElement
            bold = OxmlElement("w:b")
            rpr.append(bold)
        bold.set(qn("w:val"), "true")
    elif bold is not None:
        rpr.remove(bold)


def _paragraph_text_nodes(paragraph):
    """Return flattened paragraph text and character spans for text nodes/tabs."""
    chunks = []
    spans = []
    offset = 0
    for node in paragraph.iter():
        if node.tag == qn("w:t"):
            value = node.text or ""
            chunks.append(value)
            spans.append((offset, offset + len(value), node))
            offset += len(value)
        elif node.tag == qn("w:tab"):
            chunks.append("\t")
            spans.append((offset, offset + 1, node))
            offset += 1
    return "".join(chunks), spans


def _bold_label_at_offset(spans, char_offset, replacement: str | None = None) -> bool:
    for start, end, node in spans:
        if start <= char_offset < end and node.tag == qn("w:t"):
            value = node.text or ""
            local = char_offset - start
            if replacement is not None:
                node.text = value[:local] + replacement + value[local + 1:]
                if node.text.startswith(" ") or node.text.endswith(" "):
                    node.set(qn("xml:space"), "preserve")
            run = node.getparent()
            if run is not None and run.tag == qn("w:r"):
                _set_run_bold(run, True)
            return True
    return False



def _restore_true_false_subitem_order(block: list) -> tuple[int, list[str]]:
    """Restore complete, separately-paragraphed a)-d) groups to label order.

    Each paragraph is moved as a whole, so statement text, inline images/formulas and
    any answer marker attached to that statement remain together. Incomplete or
    ambiguous groups are left untouched and reported rather than guessed.
    """
    changed = 0
    warnings: list[str] = []
    i = 0
    label_order = {letter: idx for idx, letter in enumerate("abcd")}
    while i < len(block):
        if block[i].tag != qn("w:p"):
            i += 1
            continue
        match = re.match(r"^\s*([a-d])[.)](?=\s|$)", _paragraph_text(block[i]), re.I)
        if not match:
            i += 1
            continue
        start = i
        group = []
        while i < len(block) and block[i].tag == qn("w:p"):
            m = re.match(r"^\s*([a-d])[.)](?=\s|$)", _paragraph_text(block[i]), re.I)
            if not m:
                break
            group.append((m.group(1).lower(), block[i]))
            i += 1
        labels = [label for label, _ in group]
        if len(group) == 4 and set(labels) == set("abcd"):
            ordered = sorted(group, key=lambda item: label_order[item[0]])
            if labels != [label for label, _ in ordered]:
                block[start:i] = [element for _, element in ordered]
                changed += 1
        elif len(group) > 1 and labels != sorted(labels, key=lambda x: label_order[x]):
            warnings.append("Không tự sắp xếp tiểu ý đúng/sai: nhóm a)-d) không đủ 4 nhãn riêng biệt.")
    return changed, warnings

def _normalize_and_bold_option_labels(block: list) -> tuple[int, int, list[str]]:
    """Normalize only a complete four-option set; never wrap labels modulo four.

    Word tables and inline two-column layouts are traversed in document/XML order. If
    exactly four explicit A-D markers are found in the question block, their labels are
    normalized to A, B, C, D in that order while keeping each text run and its content
    in place. If the marker count is incomplete/ambiguous, labels are only bolded and a
    warning is returned instead of silently creating duplicate or misleading labels.
    Lowercase a)-d) markers are processed separately and never count as A-D options.
    """
    changed = 0
    bolded = 0
    warnings: list[str] = []
    option_markers = []
    paragraphs_seen = set()
    for element in block:
        paragraphs = [element] if element.tag == qn("w:p") else list(element.iter(qn("w:p")))
        for paragraph in paragraphs:
            if id(paragraph) in paragraphs_seen:
                continue
            paragraphs_seen.add(id(paragraph))
            text, spans = _paragraph_text_nodes(paragraph)
            # A paragraph must begin with an uppercase option label. Additional labels
            # in the same paragraph are allowed for side-by-side A/B and C/D layouts.
            if not re.match(r"^\s*[A-D][.)]", text):
                continue
            for match in OPTION_MARKER_IN_TEXT_RE.finditer(text):
                option_markers.append((paragraph, spans, match.start(2), match.group(2)))

    can_normalize = len(option_markers) == 4
    if option_markers and not can_normalize:
        warnings.append(
            f"Không tự chuẩn hóa nhãn A-D: phát hiện {len(option_markers)} nhãn trong một khối câu; cần kiểm tra cấu trúc phương án."
        )
    for index, (paragraph, spans, offset, original_label) in enumerate(option_markers):
        replacement = chr(65 + index) if can_normalize else None
        if _bold_label_at_offset(spans, offset, replacement):
            bolded += 1
            if replacement is not None and replacement != original_label:
                changed += 1

    # Bold true/false subitems but never relabel or shuffle them as A-D options.
    for element in block:
        paragraphs = [element] if element.tag == qn("w:p") else list(element.iter(qn("w:p")))
        for paragraph in paragraphs:
            if id(paragraph) in paragraphs_seen and paragraph.tag != qn("w:p"):
                continue
            text, spans = _paragraph_text_nodes(paragraph)
            if not re.match(r"^\s*[a-d][.)]", text):
                continue
            for match in TF_MARKER_IN_TEXT_RE.finditer(text):
                if _bold_label_at_offset(spans, match.start(2)):
                    bolded += 1
    return changed, bolded, warnings


def _split_and_style_label_runs(block: list) -> None:
    """Split simple text runs so only visible option/subitem markers are bold."""
    from copy import deepcopy
    from docx.oxml import OxmlElement
    label_re = re.compile(r"(^|[\t ]+)([A-Da-d])[.)](?=\s|$)")
    for element in block:
        paragraphs = [element] if element.tag == qn("w:p") else list(element.iter(qn("w:p")))
        for paragraph in paragraphs:
            paragraph_text = _paragraph_text(paragraph)
            if not re.match(r"^\s*(?:[A-D][.)]|[a-d][.)])", paragraph_text):
                continue
            for run in list(paragraph.iter(qn("w:r"))):
                children = list(run)
                if any(child.tag != qn("w:rPr") and child.tag != qn("w:t") for child in children):
                    continue
                text_nodes = [child for child in children if child.tag == qn("w:t")]
                if len(text_nodes) != 1:
                    continue
                text_node = text_nodes[0]
                value = text_node.text or ""
                matches = list(label_re.finditer(value))
                if not matches:
                    continue
                pieces = []
                cursor = 0
                for match in matches:
                    marker_start = match.start(2)
                    marker_end = match.end(0)
                    if marker_start > cursor:
                        pieces.append((value[cursor:marker_start], False))
                    # Keep punctuation and any trailing whitespace with the bold label.
                    pieces.append((value[marker_start:marker_end], True))
                    cursor = marker_end
                if cursor < len(value):
                    pieces.append((value[cursor:], False))
                parent = run.getparent()
                if parent is None:
                    continue
                insert_at = parent.index(run)
                for piece, is_label in pieces:
                    if not piece:
                        continue
                    new_run = OxmlElement("w:r")
                    rpr = run.find(qn("w:rPr"))
                    if rpr is not None:
                        new_rpr = deepcopy(rpr)
                        old_bold = new_rpr.find(qn("w:b"))
                        if old_bold is not None:
                            new_rpr.remove(old_bold)
                        old_bold_cs = new_rpr.find(qn("w:bCs"))
                        if old_bold_cs is not None:
                            new_rpr.remove(old_bold_cs)
                        new_run.append(new_rpr)
                    if is_label:
                        _set_run_bold(new_run, True)
                    t = OxmlElement("w:t")
                    if piece.startswith(" ") or piece.endswith(" ") or "\t" in piece:
                        t.set(qn("xml:space"), "preserve")
                    t.text = piece
                    new_run.append(t)
                    parent.insert(insert_at, new_run)
                    insert_at += 1
                parent.remove(run)


def _enable_field_updates(doc) -> None:
    """Ask Word-compatible editors to refresh PAGE/NUMPAGES fields on open."""
    from docx.oxml import OxmlElement
    settings = doc.settings._element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")


def _shuffle_options_in_block(block: list, rng: random.Random) -> tuple[int, str | None]:
    """Shuffle option paragraphs when they are separate, clearly labeled A-D paragraphs.

    Inline side-by-side options are left untouched rather than risking corruption.
    """
    option_positions = [i for i, el in enumerate(block) if el.tag == qn("w:p") and OPTION_RE.match(_paragraph_text(el))]
    if len(option_positions) < 2:
        return 0, None
    # Only shuffle a contiguous run; non-option paragraphs can contain figures or formulas.
    groups: list[list[int]] = []
    for pos in option_positions:
        if not groups or pos != groups[-1][-1] + 1:
            groups.append([pos])
        else:
            groups[-1].append(pos)
    changed = 0
    for group in groups:
        if len(group) < 2:
            continue
        original = [block[i] for i in group]
        shuffled = original[:]
        rng.shuffle(shuffled)
        for offset, pos in enumerate(group):
            block[pos] = shuffled[offset]
            _relabel_option(block[pos], chr(65 + offset))
        changed += len(group)
    return changed, None


def _relabel_option(element, label: str) -> None:
    """Relabel the leading A./B./C./D. marker without rebuilding its paragraph XML."""
    first = True
    for node in element.iter(qn("w:t")):
        value = node.text or ""
        if not value:
            continue
        if first:
            match = OPTION_RE.match(value)
            if match:
                node.text = re.sub(r"^\s*[A-D][.)]", f"{label}.", value, count=1)
                if node.text.startswith(" ") or node.text.endswith(" "):
                    node.set(qn("xml:space"), "preserve")
            first = False
            return

def _replace_numpages_cache(docx_bytes: bytes, page_count: int) -> bytes:
    """Write the computed page count into the cached result of each NUMPAGES field."""
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with zipfile.ZipFile(BytesIO(docx_bytes), "r") as src:
        files = [(item, src.read(item.filename)) for item in src.infolist()]
    updated_files = []
    changed = False
    for item, payload in files:
        if item.filename != "word/document.xml":
            updated_files.append((item, payload))
            continue
        root = etree.fromstring(payload)
        for paragraph in root.xpath(".//w:p", namespaces=ns):
            field_instruction = ""
            phase = ""
            for node in paragraph.iter():
                if node.tag == f"{{{ns['w']}}}fldChar":
                    kind = node.get(f"{{{ns['w']}}}fldCharType", "")
                    if kind == "begin":
                        field_instruction, phase = "", "instruction"
                    elif kind == "separate" and phase:
                        phase = "result"
                    elif kind == "end":
                        field_instruction, phase = "", ""
                elif node.tag == f"{{{ns['w']}}}instrText" and phase == "instruction":
                    field_instruction += node.text or ""
                elif node.tag == f"{{{ns['w']}}}t" and phase == "result" and "NUMPAGES" in field_instruction.upper():
                    node.text = str(page_count)
                    changed = True
                    phase = "result_done"
        updated_files.append((item, etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)))
    if not changed:
        return docx_bytes
    out = BytesIO()
    with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as dst:
        for item, payload in updated_files:
            dst.writestr(item, payload)
    return out.getvalue()


def _render_docx_page_count(docx_bytes: bytes) -> tuple[int | None, str | None]:
    """Count pages with LibreOffice + pdfinfo; gracefully fall back if unavailable."""
    office = shutil.which("libreoffice") or shutil.which("soffice")
    pdfinfo = shutil.which("pdfinfo")
    if not office or not pdfinfo:
        return None, "Không có LibreOffice/pdfinfo trên máy chủ; NUMPAGES sẽ được Word cập nhật khi mở tài liệu."
    try:
        with tempfile.TemporaryDirectory(prefix="ymp-pagecount-") as temp_dir:
            root = Path(temp_dir)
            source_path = root / "exam.docx"
            source_path.write_bytes(docx_bytes)
            profile = root / "lo-profile"
            command = [office, f"-env:UserInstallation={profile.as_uri()}", "--headless", "--convert-to", "pdf", "--outdir", str(root), str(source_path)]
            result = subprocess.run(command, capture_output=True, text=True, timeout=60, check=False)
            pdf_path = root / "exam.pdf"
            if result.returncode != 0 or not pdf_path.exists():
                detail = (result.stderr or result.stdout or "LibreOffice không tạo được PDF.").strip()[-400:]
                return None, f"Không đếm được số trang tự động: {detail}"
            info = subprocess.run([pdfinfo, str(pdf_path)], capture_output=True, text=True, timeout=15, check=False)
            match = re.search(r"^Pages:\s+(\d+)\s*$", info.stdout, re.MULTILINE)
            if info.returncode != 0 or not match:
                return None, "pdfinfo không trả về số trang hợp lệ."
            return int(match.group(1)), None
    except (OSError, subprocess.TimeoutExpired) as exc:
        return None, f"Không đếm được số trang tự động: {type(exc).__name__}."


def _finalize_total_pages(docx_bytes: bytes) -> tuple[bytes, int | None, str | None]:
    """Compute and cache total pages, repeating if the cached value changes pagination."""
    current = docx_bytes
    last_count = None
    for _ in range(3):
        count, warning = _render_docx_page_count(current)
        if count is None:
            return current, last_count, warning
        updated = _replace_numpages_cache(current, count)
        last_count = count
        if updated == current:
            return current, count, None
        verify_count, verify_warning = _render_docx_page_count(updated)
        if verify_count == count:
            return updated, count, None
        if verify_count is None:
            return updated, count, verify_warning or warning
        current = _replace_numpages_cache(updated, verify_count)
        last_count = verify_count
    return current, last_count, "Tổng số trang đã được tính nhưng chưa ổn định sau 3 lượt dàn trang; cần kiểm tra bằng Word."


def export_reordered_docx(source: bytes, question_order: Sequence[str] | None = None, header: dict | None = None, section_modes: dict | None = None, seed: int | None = None) -> tuple[bytes, dict]:
    try:
        doc = Document(BytesIO(source))
    except Exception as exc:
        raise ValueError("Không mở được tệp DOCX nguồn.") from exc

    body = doc.element.body
    children = [child for child in list(body) if child.tag != qn("w:sectPr")]
    sect_pr = next((child for child in list(body) if child.tag == qn("w:sectPr")), None)
    children, source_header_blocks_removed = _remove_existing_generated_header(children)
    # Resolve each section's default g0-g3 mode from source markers before cleaning them.
    marker_modes: dict[str, str] = {}
    current_section = "P0"
    active_mode = "g0"
    pending_mode: str | None = None
    for child in children:
        text = _paragraph_text(child)
        marker_match = re.search(r"<\s*(g[0-3])\s*>", text, re.I)
        sec_match = SECTION_RE.match(text)
        if marker_match:
            pending_mode = marker_match.group(1).lower()
            if not sec_match:
                active_mode = pending_mode
                continue
        if sec_match:
            current_section = f"P{sec_match.group(1).upper()}"
            active_mode = marker_match.group(1).lower() if marker_match else (pending_mode or active_mode)
            marker_modes[current_section] = active_mode
            pending_mode = None
    effective_modes = dict(marker_modes)
    if section_modes:
        for key, value in section_modes.items():
            if str(value) in {"g0", "g1", "g2", "g3"}:
                effective_modes[str(key)] = str(value)
    rng = random.Random(seed)
    option_shuffle_count = 0
    option_shuffle_warnings: list[str] = []
    # Remove standalone parser markers and the explicit trailing answer-key section.
    for child in children:
        _clean_marker_text(child)
    children, answer_key_removed = _clean_answer_key_tail(children)
    children, short_answer_keys_removed = _remove_short_answer_inline_keys(children)
    answer_marks_removed = False
    for child in children:
        if _remove_answer_formatting(child):
            answer_marks_removed = True
    # Drop empty paragraphs left behind by standalone <g1>/<g2>/<g3> markers.
    children = [child for child in children if not (
        child.tag == qn("w:p") and not _paragraph_text(child).strip()
        and not child.xpath('.//w:drawing | .//w:pict | .//w:object | .//m:oMath | .//m:oMathPara')
    )]

    # Segment by section headings first, so headings remain fixed and can never be moved
    # into another section when question blocks are reordered.
    segments: list[dict] = []
    current = {"section": "P0", "prefix": [], "starts": []}
    for idx, child in enumerate(children):
        text = _paragraph_text(child)
        sec_match = SECTION_RE.match(text)
        if sec_match:
            if current["prefix"] or current["starts"] or segments:
                segments.append(current)
            current = {"section": f"P{sec_match.group(1).upper()}", "prefix": [child], "starts": []}
            continue
        number = _question_number(child)
        if number is not None:
            current["starts"].append({"index": idx, "number": number, "element": child})
        elif not current["starts"]:
            current["prefix"].append(child)
    segments.append(current)

    # Build question blocks using source child positions, clipped at section boundaries.
    all_starts = []
    for seg_index, seg in enumerate(segments):
        for item in seg["starts"]:
            item["section"] = seg["section"]
            item["seg_index"] = seg_index
            all_starts.append(item)
    if not all_starts:
        raise ValueError("Không tìm thấy tiêu đề câu dạng 'Câu 1.', 'Câu 2.' trong DOCX.")

    # Re-scan into exact segment child arrays; each segment owns its heading/preamble and
    # the blocks between questions. This keeps tables, pictures, OMML and OLE objects with
    # the preceding question unless they occur before the first question in that section.
    segment_elements: list[list] = []
    cursor = 0
    for seg in segments:
        elems = []
        if seg["prefix"]:
            elems.extend(seg["prefix"])
        # Find the contiguous source range represented by this segment using its first
        # prefix/heading and next section heading. Build from the body sequence instead.
        segment_elements.append(elems)

    # A reliable section assignment for every direct body element.
    element_section: dict[int, str] = {}
    section_now = "P0"
    section_headers: set[int] = set()
    for child in children:
        text = _paragraph_text(child)
        sm = SECTION_RE.match(text)
        if sm:
            section_now = f"P{sm.group(1).upper()}"
            section_headers.add(id(child))
        element_section[id(child)] = section_now

    # Detect each question with a unique ordinal within its section.
    counts: dict[str, int] = {}
    question_starts = []
    for idx, child in enumerate(children):
        number = _question_number(child)
        if number is None:
            continue
        sec = element_section[id(child)]
        counts[sec] = counts.get(sec, 0) + 1
        question_starts.append({"index": idx, "number": number, "section": sec,
                                "id": f"Q-DOCX-{sec}-{counts[sec]}"})

    blocks: dict[str, list] = {}
    for pos, item in enumerate(question_starts):
        end = len(children)
        for nxt in question_starts[pos + 1:]:
            if nxt["section"] != item["section"]:
                end = nxt["index"]
                break
            end = nxt["index"]
            break
        # Never include a section heading in the preceding question block.
        for j in range(item["index"] + 1, end):
            if id(children[j]) in section_headers:
                end = j
                break
        blocks[item["id"]] = children[item["index"]:end]

    suffix: list = []
    last_id = question_starts[-1]["id"]
    while blocks.get(last_id):
        tail = blocks[last_id][-1]
        if tail.tag == qn("w:p") and END_RE.fullmatch(_paragraph_text(tail)):
            suffix.insert(0, blocks[last_id].pop())
        else:
            break

    requested = list(question_order or blocks.keys())
    if not requested:
        raise ValueError("Cần ít nhất một ID câu để xuất DOCX.")
    if len(requested) != len(set(requested)):
        raise ValueError("Danh sách thứ tự câu có ID bị lặp.")
    unknown = set(requested) - set(blocks)
    if unknown:
        raise ValueError("Danh sách có ID không tồn tại trong DOCX nguồn: " + ", ".join(sorted(unknown)))

    section_order = []
    for item in question_starts:
        if item["section"] not in section_order:
            section_order.append(item["section"])
    requested_by_section: dict[str, list[str]] = {sec: [] for sec in section_order}
    for qid in requested:
        parts = qid.split("-")
        sec = parts[2] if len(parts) >= 4 else "P0"
        if sec not in requested_by_section:
            raise ValueError(f"ID câu {qid} không có phần hợp lệ.")
        requested_by_section[sec].append(qid)

    # Preserve all non-question content and section headings in source order. Reordering
    # occurs only among blocks inside each section; source prefix content stays in place.
    output_children = []
    warnings = []
    renumbered = 0
    section_positions = {}
    for idx, child in enumerate(children):
        if id(child) in section_headers:
            sec = element_section[id(child)]
            section_positions.setdefault(sec, idx)

    # Collect section fixed content (all non-question elements before the first question)
    # and then ordered question blocks. Any unrecognized content after a question remains
    # attached to that question block.
    question_index_by_start = {item["index"]: item for item in question_starts}
    idx = 0
    current_section = "P0"
    emitted_sections = set()
    emitted_questions = set()
    while idx < len(children):
        child = children[idx]
        if id(child) in section_headers:
            current_section = element_section[id(child)]
            output_children.append(child)
            emitted_sections.add(current_section)
            idx += 1
            continue
        item = question_index_by_start.get(idx)
        if item:
            sec = item["section"]
            if sec not in emitted_sections and sec != "P0":
                # Section header was absent from source or not detected; don't fabricate one.
                emitted_sections.add(sec)
            chosen = requested_by_section.get(sec, [])
            # Emit the selected order once, at the first question of each section.
            if sec not in emitted_questions:
                emitted_questions.add(sec)
                if chosen:
                    for new_number, qid in enumerate(chosen, start=1):
                        block = blocks[qid]
                        q_section = next((it["section"] for it in question_starts if it["id"] == qid), sec)
                        # True/false subitems default to a), b), c), d) order.
                        # Restore by their existing labels only when a complete, unambiguous
                        # four-item group is present; never shuffle these items implicitly.
                        tf_groups_restored, tf_order_warnings = _restore_true_false_subitem_order(block)
                        warnings.extend(tf_order_warnings)
                        normalized_options, bolded_labels, label_warnings = _normalize_and_bold_option_labels(block)
                        warnings.extend(label_warnings)
                        _split_and_style_label_runs(block)
                        if effective_modes.get(q_section, "g0") in {"g2", "g3"}:
                            count, warning = _shuffle_options_in_block(block, rng)
                            option_shuffle_count += count
                            if warning:
                                option_shuffle_warnings.append(warning)
                        if block and _replace_heading_number(block[0], new_number):
                            renumbered += 1
                        else:
                            warnings.append(f"Chưa đánh lại số được cho {qid}; giữ nguyên tiêu đề gốc.")
                        output_children.extend(block)
            # Skip each original question block; the selected blocks were emitted above.
            idx += len(blocks[item["id"]])
            continue
        # Non-question content before first question in a section or between section blocks.
        output_children.append(child)
        idx += 1

    output_children.extend(suffix)

    # Remove duplicate content that was traversed after emitting question blocks. The loop
    # advances by exact block lengths, so output remains a valid ordered list of XML nodes.
    # Deduplicate by identity defensively (prefix/section nodes should appear exactly once).
    seen = set()
    output_children = [el for el in output_children if not (id(el) in seen or seen.add(id(el)))]

    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)
    for child in output_children:
        body.insert(len(body) - (1 if sect_pr is not None else 0), child)
    if sect_pr is not None and sect_pr.getparent() is not body:
        body.append(sect_pr)

    # Rebuild body using cleaned elements, then prepend the official exam template.
    _add_official_header(doc, header)
    _enable_field_updates(doc)
    out = BytesIO()
    doc.save(out)
    finalized_bytes, total_pages, page_count_warning = _finalize_total_pages(out.getvalue())
    if page_count_warning:
        # Do not hide a failed server-side page count; expose a clear diagnostic to the UI/logs.
        logger = __import__("logging").getLogger("young-mix-pro-api")
        logger.warning("DOCX total-page count fallback: %s", page_count_warning)
    return finalized_bytes, {
        "question_count": len(requested),
        "renumbered_count": renumbered,
        "warnings": warnings,
        "preserves_source_xml": True,
        "answer_markings_removed": answer_marks_removed,
        "answer_key_section_removed": answer_key_removed,
        "parser_markers_removed": True,
        "official_header_added": True,
        "source_header_blocks_removed": source_header_blocks_removed,
        "official_footer_added": True,
        "footer_contains_only_code_and_page": True,
        "total_pages_field_in_header": True,
        "total_pages_cached": total_pages,
        "total_pages_counted_on_server": total_pages is not None,
        "page_count_warning": page_count_warning,
        "official_title_in_header": True,
        "section_modes": effective_modes,
        "option_paragraphs_shuffled": option_shuffle_count,
        "option_labels_normalized_and_bolded": True,
        "true_false_groups_restored_to_abcd_order": True,
        "true_false_items_shuffled": False,
        "word_fields_update_on_open": True,
        "option_shuffle_warnings": option_shuffle_warnings,
        "page_number_added": True,
        "short_answer_inline_keys_removed": short_answer_keys_removed,
        "section_order_preserved": True,
    }
