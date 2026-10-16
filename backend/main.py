import os
from fastapi import FastAPI,File,UploadFile,HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from models import MixRequest
from engine import mix_exam
from docx_parser import parse_docx
from export_excel import export_answer_key
app=FastAPI(title="Young Mix Pro API",version="0.1.0")
origins=[x.strip() for x in os.getenv("CORS_ORIGINS","http://localhost:5173,http://127.0.0.1:5173").split(",") if x.strip()]
app.add_middleware(CORSMiddleware,allow_origins=origins,allow_credentials=True,allow_methods=["GET","POST","OPTIONS"],allow_headers=["*"])
@app.get("/api/health")
def health(): return {"ok":True,"service":"young-mix-pro-api","version":"0.1.0"}
@app.post("/api/exams/parse")
async def parse(file:UploadFile=File(...)):
    if not file.filename or not file.filename.lower().endswith(".docx"): raise HTTPException(400,"Endpoint này hiện chỉ nhận DOCX.")
    data=await file.read()
    if len(data)>20*1024*1024: raise HTTPException(413,"Tệp vượt giới hạn 20 MB.")
    try: return parse_docx(data)
    except Exception as e: raise HTTPException(400,f"Không đọc được DOCX: {e}")
@app.post("/api/exams/mix")
def mix(req:MixRequest):
    if len(req.questions)>1000: raise HTTPException(400,"MVP giới hạn 1000 câu mỗi yêu cầu.")
    return mix_exam(req)
@app.post("/api/exams/export-answer-key")
def export(payload:dict):
    codes=payload.get("codes",[])
    if not isinstance(codes,list): raise HTTPException(400,"codes phải là danh sách.")
    path="/tmp/young-mix-pro-answer-key.xlsx"; export_answer_key(codes,path)
    return FileResponse(path,filename="young-mix-pro-answer-key.xlsx",media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
