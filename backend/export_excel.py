from openpyxl import Workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.utils import get_column_letter
def export_answer_key(codes,path):
    # Workbook tổng quát, chưa phải mẫu đáp án chính thức 81 cột.
    wb=Workbook(); ws=wb.active; ws.title="Đáp án"
    ws.append(["Mã đề","Số câu","ID câu hỏi","Dạng câu","Đáp án","Điểm"])
    for code in codes:
        for q in code.get("questions",[]):
            ans=q.get("answer")
            if isinstance(ans,list): ans=", ".join(map(str,ans))
            ws.append([str(code.get("code","")),q.get("number",""),q.get("id",""),q.get("type",""),ans or "",q.get("points","")])
    for c in ws[1]:
        c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="176B52"); c.alignment=Alignment(horizontal="center")
    for i,w in enumerate([14,12,24,20,38,12],1): ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions; wb.save(path)
