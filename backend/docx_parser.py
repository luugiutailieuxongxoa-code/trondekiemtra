"""Conservative DOCX structure scanner for Young Mix Pro.

Extracts review metadata from Word body blocks while leaving the original DOCX bytes
available for source-preserving export. It does not claim OCR or semantic formula parsing.
"""
from __future__ import annotations
import re
from io import BytesIO
from docx import Document
from docx.oxml.ns import qn

QRE = re.compile(r"^\s*Câu\s*(\d+)\s*[.:)]?\s*(.*)$", re.I)
ORE = re.compile(r"^\s*([A-D])\s*[.)]\s*(.*)$", re.I)
SECTION_RE = re.compile(r"^\s*PHẦN\s+([IVX]+|\d+)\b", re.I)


def _text(el) -> str:
    return "".join((node.text or "") for node in el.iter(qn("w:t"))).strip()


def _has_drawing(el) -> bool:
    return any(el.iter(qn(tag)) for tag in ("w:drawing", "w:pict", "w:object"))


def _body_blocks(doc):
    """Yield direct body elements in their original order (paragraphs and tables)."""
    body = doc.element.body
    for el in body:
        if el.tag == qn("w:sectPr"):
            continue
        yield el


def parse_docx(content: bytes):
    try:
        doc = Document(BytesIO(content))
    except Exception as exc:
        raise ValueError("Không mở được DOCX") from exc

    blocks = list(_body_blocks(doc))
    questions = []
    section = "P0"
    section_counts: dict[str, int] = {}
    current = None
    all_text = []
    image_count = 0
    table_count = 0

    def flush():
        nonlocal current
        if current is not None:
            current["text"] = "\n".join(x for x in current["_lines"] if x).strip()
            current.pop("_lines", None)
            questions.append(current)
            current = None

    for block in blocks:
        if block.tag == qn("w:tbl"):
            table_count += 1
            # Tables remain attached to the current question on export; include readable
            # table text in preview metadata only when a question is active.
            table_text = " | ".join(t.strip() for t in _text(block).splitlines() if t.strip())
            if current and table_text:
                current["_lines"].append(table_text)
            continue

        text = _text(block)
        if _has_drawing(block):
            image_count += 1
        if text:
            all_text.append(text)
        sec = SECTION_RE.match(text) if text else None
        if sec:
            flush()
            roman = sec.group(1).upper()
            section = f"P{roman}"
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
                "type": "multiple_choice",
                "topic": "Chưa phân loại",
                "level": "Chưa xác định",
                "text": "",
                "options": [],
                "answer": None,
                "points": 0,
                "group": "g0",
                "has_image": False,
                "review_required": True,
                "_lines": [q.group(2).strip()] if q.group(2).strip() else [],
            }
            continue
        if current:
            if _has_drawing(block):
                current["has_image"] = True
            opt = ORE.match(text) if text else None
            if opt:
                current["options"].append({
                    "id": f"{current['id']}-O{opt.group(1).upper()}",
                    "label": opt.group(1).upper(),
                    "text": opt.group(2).strip(),
                })
            elif text:
                current["_lines"].append(text)

    flush()
    # Images embedded in the main document XML, including diagrams inside tables.
    try:
        image_count = len(doc.part._element.xpath(".//pic:pic"))
    except Exception:
        image_count = 0

    warnings = [
        "Đây là nhận diện cấu trúc sơ bộ; giáo viên cần đối chiếu với DOCX gốc trước khi trộn.",
        "Công thức OMML/MathType, ảnh nổi, biểu đồ và nội dung OLE được giữ trong DOCX nguồn khi xuất theo chế độ bảo toàn nguồn; chúng không được chuyển đổi đầy đủ thành văn bản xem trước.",
        "Đáp án đúng không được tự suy luận. Hãy xác nhận đáp án và cách đánh dấu đáp án trước khi phát hành đề.",
    ]
    if image_count:
        warnings.append(f"Phát hiện khoảng {image_count} đối tượng hình/đồ họa; cần kiểm tra hình có đi cùng đúng câu sau khi xuất.")
    if table_count:
        warnings.append(f"Phát hiện {table_count} bảng trong DOCX; nội dung bảng được giữ ở cấp XML nhưng cần kiểm tra trực quan.")
    if not questions:
        warnings.append("Chưa nhận diện được tiêu đề câu theo dạng 'Câu 1.', 'Câu 2.'; có thể tài liệu dùng bố cục hoặc hộp văn bản chưa hỗ trợ.")

    return {
        "question_count": len(questions),
        "image_count": image_count,
        "table_count": table_count,
        "sections": sorted({q["section"] for q in questions}, key=lambda s: (s == "P0", s)),
        "questions": questions,
        "warnings": warnings,
        "review_required": True,
        "source_preserving_export_available": True,
    }
