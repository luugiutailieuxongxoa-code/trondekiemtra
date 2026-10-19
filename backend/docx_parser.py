"""DOCX structure scanner for Young Mix Pro.

The parser creates a reviewable question index while the original DOCX remains the
source of truth for images, equations, tables, drawings and embedded objects.
"""
from __future__ import annotations
import re
from io import BytesIO
from docx import Document
from docx.oxml.ns import qn

QRE = re.compile(r"^\s*Câu\s*(\d+)\s*[.:)]?\s*(.*)$", re.I)
ORE = re.compile(r"^\s*([A-D])\s*[.)]\s*(.*)$", re.I)
TFRE = re.compile(r"^\s*([a-d])\s*[.)]\s*(.*)$", re.I)
SECTION_RE = re.compile(r"^\s*(?:<g\d+>\s*)?PHẦN\s+([IVXLCDM]+|\d+)\b", re.I)
END_RE = re.compile(r"^\s*(?:[-—– ]*HẾT[-—– ]*|[-—– ]*KẾT THÚC[-—– ]*)\s*$", re.I)


def _text(el) -> str:
    return "".join((node.text or "") for node in el.iter(qn("w:t"))).strip()


def _has_drawing(el) -> bool:
    return any(el.iter(qn(tag)) for tag in ("w:drawing", "w:pict", "w:object"))


def _body_blocks(doc):
    body = doc.element.body
    for el in body:
        if el.tag != qn("w:sectPr"):
            yield el


def _section_type(section: str) -> str:
    if section == "PII":
        return "true_false"
    if section == "PIII":
        return "short_answer"
    return "multiple_choice"


def parse_docx(content: bytes):
    try:
        doc = Document(BytesIO(content))
    except Exception as exc:
        raise ValueError("Không mở được DOCX") from exc

    blocks = list(_body_blocks(doc))
    questions = []
    section = "P0"
    section_titles: dict[str, str] = {}
    section_counts: dict[str, int] = {}
    current = None
    all_text = []
    table_count = 0

    def flush():
        nonlocal current
        if current is not None:
            current["text"] = "\n".join(x for x in current.pop("_lines", []) if x).strip()
            questions.append(current)
            current = None

    for block in blocks:
        is_table = block.tag == qn("w:tbl")
        if is_table:
            table_count += 1
            # Include visible text and mark media inside tables as belonging to the active
            # question. The original table remains intact in the source DOCX used for export.
            table_text = " | ".join(t.strip() for t in _text(block).splitlines() if t.strip())
            if current:
                current["has_table"] = True
                if _has_drawing(block):
                    current["has_image"] = True
                if table_text:
                    current["_lines"].append(table_text)
            continue

        text = _text(block)
        has_image = _has_drawing(block)
        if text:
            all_text.append(text)
        sec = SECTION_RE.match(text) if text else None
        if sec:
            flush()
            roman = sec.group(1).upper()
            section = f"P{roman}"
            section_titles[section] = text
            continue

        q = QRE.match(text) if text else None
        if q:
            flush()
            section_counts[section] = section_counts.get(section, 0) + 1
            ordinal = section_counts[section]
            number = int(q.group(1))
            current = {
                "id": f"Q-DOCX-{section}-{ordinal}",
                "source_id": f"Q-DOCX-{section}-{ordinal}",
                "number": number,
                "source_number": number,
                "section": section,
                "section_title": section_titles.get(section, ""),
                "type": _section_type(section),
                "topic": "Chưa phân loại",
                "level": "Chưa xác định",
                "text": "",
                "options": [],
                "answer": None,
                "points": 0,
                "group": "g0",
                "has_image": has_image,
                "has_table": False,
                "review_required": True,
                "_lines": [q.group(2).strip()] if q.group(2).strip() else [],
            }
            continue

        if current:
            if has_image:
                current["has_image"] = True
            if text:
                option = TFRE.match(text) if current["type"] == "true_false" else ORE.match(text)
                if option:
                    current["options"].append({
                        "id": f"{current['id']}-O{option.group(1).upper()}",
                        "label": option.group(1).upper(),
                        "text": option.group(2).strip(),
                    })
                elif not END_RE.match(text):
                    current["_lines"].append(text)

    flush()

    try:
        image_count = len(doc.part._element.xpath(".//pic:pic"))
    except Exception:
        image_count = sum(1 for b in blocks if _has_drawing(b))

    section_summary = []
    for sec in dict.fromkeys(q["section"] for q in questions):
        qs = [q for q in questions if q["section"] == sec]
        section_summary.append({
            "id": sec,
            "title": section_titles.get(sec) or ("Phần chưa xác định" if sec == "P0" else f"Phần {sec[1:]}"),
            "question_count": len(qs),
            "question_numbers": [q["source_number"] for q in qs],
            "type": qs[0]["type"],
            "image_count": sum(1 for q in qs if q["has_image"]),
            "table_count": sum(1 for q in qs if q["has_table"]),
        })

    warnings = [
        "Đây là nhận diện cấu trúc để hỗ trợ trộn; giáo viên cần đối chiếu với DOCX gốc trước khi phát hành.",
        "Hình ảnh, công thức OMML/MathType, bảng và đối tượng nhúng được bảo toàn bằng cách xuất lại từ DOCX nguồn; phần xem trước không tái dựng chính xác các đối tượng này.",
        "Đáp án đúng không được tự suy luận. Cần xác nhận đáp án trước khi phát hành đề.",
    ]
    if image_count:
        warnings.append(f"Phát hiện {image_count} hình/đồ họa trong tài liệu nguồn.")
    if table_count:
        warnings.append(f"Phát hiện {table_count} bảng trong tài liệu nguồn.")
    if not questions:
        warnings.append("Không nhận diện được tiêu đề câu trong đoạn văn chính. Tài liệu có thể dùng hộp văn bản, cột hoặc bảng để chứa câu hỏi; cần kiểm tra thủ công.")

    return {
        "question_count": len(questions),
        "image_count": image_count,
        "table_count": table_count,
        "sections": [s["id"] for s in section_summary],
        "section_summary": section_summary,
        "questions": questions,
        "warnings": warnings,
        "review_required": True,
        "source_preserving_export_available": True,
    }
