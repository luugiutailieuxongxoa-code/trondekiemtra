import re
from io import BytesIO
from docx import Document
QRE=re.compile(r"^\s*Câu\s*(\d+)\s*[\.:)]?\s*(.*)$",re.I)
ORE=re.compile(r"^\s*([A-D])[\.\)]\s*(.*)$",re.I)
def parse_docx(content:bytes):
    doc=Document(BytesIO(content))
    lines=[p.text.strip() for p in doc.paragraphs if p.text.strip()]
    try: images=len(doc.part._element.xpath(".//pic:pic"))
    except Exception: images=0
    questions=[]; cur=None
    for line in lines:
        m=QRE.match(line)
        if m:
            if cur: questions.append(cur)
            cur={"id":f"Q-DOCX-{m.group(1)}","number":int(m.group(1)),"type":"multiple_choice","topic":"Chưa phân loại","level":"Biết","text":m.group(2).strip(),"options":[],"answer":None,"points":.25,"group":"g0"}
        elif cur:
            o=ORE.match(line)
            if o: cur["options"].append({"id":f'{cur["id"]}-O{o.group(1).upper()}',"text":o.group(2).strip()})
            else: cur["text"] += ("\n" if cur["text"] else "")+line
    if cur: questions.append(cur)
    warnings=["Parser này chỉ trích xuất văn bản đoạn cơ bản; cần đối chiếu DOCX gốc.","Không bảo đảm bảo toàn OMML/MathType, ảnh nổi, biểu đồ, bảng hoặc bố cục.","Đáp án không được tự suy luận; cần giáo viên xác nhận."]
    if images: warnings.append(f"Phát hiện {images} đối tượng hình ảnh; MVP chưa gắn ảnh vào câu hỏi.")
    if doc.tables: warnings.append(f"Phát hiện {len(doc.tables)} bảng; nội dung bảng chưa được phân tích.")
    return {"question_count":len(questions),"image_count":images,"table_count":len(doc.tables),"questions":questions,"warnings":warnings}
