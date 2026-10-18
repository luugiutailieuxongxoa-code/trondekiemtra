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
from docx import Document
from docx.oxml.ns import qn

QUESTION_RE = re.compile(r"^\s*Câu\s*(\d+)\s*[.:)]?", re.IGNORECASE)
SECTION_RE = re.compile(r"^\s*PHẦN\s+([IVX]+|\d+)\b", re.IGNORECASE)
END_RE = re.compile(r"\s*(?:[-—– ]*HẾT[-—– ]*|[-—– ]*KẾT THÚC[-—– ]*)\s*", re.I)


def _paragraph_text(element) -> str:
    return "".join(node.text or "" for node in element.iter(qn("w:t"))).strip() if element.tag == qn("w:p") else ""


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


def export_reordered_docx(source: bytes, question_order: Sequence[str] | None = None) -> tuple[bytes, dict]:
    try:
        doc = Document(BytesIO(source))
    except Exception as exc:
        raise ValueError("Không mở được tệp DOCX nguồn.") from exc

    body = doc.element.body
    children = [child for child in list(body) if child.tag != qn("w:sectPr")]
    sect_pr = next((child for child in list(body) if child.tag == qn("w:sectPr")), None)

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

    out = BytesIO()
    doc.save(out)
    return out.getvalue(), {
        "question_count": len(requested),
        "renumbered_count": renumbered,
        "warnings": warnings,
        "preserves_source_xml": True,
        "answer_markings_removed": False,
        "section_order_preserved": True,
    }
