"""Young Mix Pro API - phase 1 foundation.

Run from this directory:
    uvicorn main:app --host 0.0.0.0 --port 8000
"""
from __future__ import annotations

import logging
import os
import tempfile
from pathlib import Path
from typing import Any

from fastapi import FastAPI, File, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse
from starlette.background import BackgroundTask
from pydantic import ValidationError

from docx_parser import parse_docx
from engine import mix_exam
from export_excel import export_answer_key
from word_export import export_reordered_docx
import json
from models import MixRequest

VERSION = "0.4.0"
MAX_UPLOAD_BYTES = int(os.getenv("MAX_UPLOAD_BYTES", str(20 * 1024 * 1024)))
logging.basicConfig(level=os.getenv("LOG_LEVEL", "INFO"))
logger = logging.getLogger("young-mix-pro-api")

app = FastAPI(
    title="Young Mix Pro API",
    version=VERSION,
    description="API nền tảng cho nhập DOCX, trộn đề và xuất bảng đáp án.",
)

origins = [
    value.strip().rstrip("/")
    for value in os.getenv(
        "CORS_ORIGINS", "http://localhost:5173,http://127.0.0.1:5173"
    ).split(",")
    if value.strip()
]
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["Content-Type", "Authorization", "X-Request-ID"],
)


@app.middleware("http")
async def request_logging(request: Request, call_next):
    """Log method/path/status without logging uploaded content or secrets."""
    try:
        response = await call_next(request)
        logger.info("%s %s -> %s", request.method, request.url.path, response.status_code)
        return response
    except Exception:
        logger.exception("Unhandled API error on %s %s", request.method, request.url.path)
        return JSONResponse(
            status_code=500,
            content={"detail": "Lỗi máy chủ không xác định.", "code": "INTERNAL_ERROR"},
        )


@app.get("/")
def root() -> dict[str, Any]:
    return {
        "service": "young-mix-pro-api",
        "version": VERSION,
        "health": "/api/health",
        "docs": "/docs",
    }


@app.get("/api/health")
def health() -> dict[str, Any]:
    return {"ok": True, "service": "young-mix-pro-api", "version": VERSION}


@app.get("/api/config")
def public_config() -> dict[str, Any]:
    """Non-secret diagnostics to help identify frontend/backend misconfiguration."""
    return {
        "service": "young-mix-pro-api",
        "version": VERSION,
        "features": {"docx_import": True, "pdf_import": False, "ocr": False,
                     "mix_api": True, "excel_export_basic": True,
                     "docx_export_source_preserving": True, "docx_section_aware_parse": True, "pdf_export": False,
                     "database": False, "authentication": False},
        "max_upload_bytes": MAX_UPLOAD_BYTES,
    }


async def read_limited_upload(file: UploadFile) -> bytes:
    if not file.filename:
        raise HTTPException(status_code=400, detail="Tệp tải lên không có tên.")
    if Path(file.filename).suffix.lower() != ".docx":
        raise HTTPException(
            status_code=415,
            detail="Phiên bản hiện tại chỉ nhận tệp DOCX. PDF/OCR chưa được triển khai.",
        )

    chunks: list[bytes] = []
    total = 0
    while True:
        chunk = await file.read(1024 * 1024)
        if not chunk:
            break
        total += len(chunk)
        if total > MAX_UPLOAD_BYTES:
            raise HTTPException(
                status_code=413,
                detail=f"Tệp vượt giới hạn {MAX_UPLOAD_BYTES // (1024 * 1024)} MB.",
            )
        chunks.append(chunk)
    data = b"".join(chunks)
    if not data:
        raise HTTPException(status_code=400, detail="Tệp DOCX đang trống.")
    return data


@app.post("/api/exams/parse")
async def parse_exam(file: UploadFile = File(...)) -> dict[str, Any]:
    try:
        data = await read_limited_upload(file)
        try:
            result = parse_docx(data)
            result["source_filename"] = Path(file.filename or "de.docx").name
            result["review_required"] = True
            return result
        except (ValueError, ValidationError) as exc:
            logger.info("DOCX parse rejected: %s", type(exc).__name__)
            raise HTTPException(
                status_code=422,
                detail="Không phân tích được cấu trúc DOCX. Hãy kiểm tra tệp hoặc dùng DOCX hợp lệ.",
            ) from exc
        except Exception as exc:
            # Do not return internal paths or implementation details to the browser.
            logger.exception("DOCX parser failed")
            raise HTTPException(
                status_code=422,
                detail="Không đọc được tệp DOCX. Tệp có thể bị hỏng hoặc chứa cấu trúc chưa hỗ trợ.",
            ) from exc
    finally:
        await file.close()


@app.post("/api/exams/mix")
def mix(request: MixRequest) -> dict[str, Any]:
    if not request.questions:
        raise HTTPException(status_code=422, detail="Cần ít nhất một câu hỏi để trộn.")
    if len(request.questions) > 1000:
        raise HTTPException(status_code=413, detail="Giới hạn 1000 câu mỗi yêu cầu.")
    try:
        return mix_exam(request)
    except Exception as exc:
        logger.exception("Mix engine failed")
        raise HTTPException(status_code=500, detail="Không thể trộn đề với dữ liệu đã gửi.") from exc



@app.post("/api/exams/export-docx")
async def export_docx(
    file: UploadFile = File(...),
    question_order: str | None = None,
) -> FileResponse:
    """Export a DOCX by reordering source question blocks without rebuilding them.

    `question_order` is an optional JSON array, e.g. ["Q-DOCX-2", "Q-DOCX-1"].
    The order may contain a subset of detected question IDs; omitted question blocks
    are excluded from the exported paper.
    """
    temp_path: str | None = None
    try:
        data = await read_limited_upload(file)
        order = None
        if question_order:
            try:
                parsed = json.loads(question_order)
            except json.JSONDecodeError as exc:
                raise HTTPException(status_code=422, detail="question_order phải là mảng JSON hợp lệ.") from exc
            if not isinstance(parsed, list) or not all(isinstance(item, str) for item in parsed):
                raise HTTPException(status_code=422, detail="question_order phải là danh sách ID câu dạng chuỗi.")
            order = parsed
        try:
            exported, report = export_reordered_docx(data, order)
        except ValueError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc
        temp = tempfile.NamedTemporaryFile(prefix="ymp-exam-", suffix=".docx", delete=False)
        temp_path = temp.name
        temp.close()
        Path(temp_path).write_bytes(exported)
        # Use a neutral filename; original user filename is not trusted as a path.
        return FileResponse(
            temp_path,
            filename="young-mix-pro-exam.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "X-YMP-Question-Count": str(report["question_count"]),
                "X-YMP-Renumbered-Count": str(report["renumbered_count"]),
                "X-YMP-Answer-Markings-Removed": "false",
                "X-YMP-Section-Order-Preserved": "true",
                "X-YMP-Warning-Count": str(len(report["warnings"])),
            },
            background=BackgroundTask(Path(temp_path).unlink, missing_ok=True),
        )
    except HTTPException:
        if temp_path:
            Path(temp_path).unlink(missing_ok=True)
        raise
    except Exception as exc:
        if temp_path:
            Path(temp_path).unlink(missing_ok=True)
        logger.exception("DOCX export failed")
        raise HTTPException(status_code=500, detail="Không thể xuất DOCX.") from exc
    finally:
        await file.close()


@app.post("/api/exams/export-answer-key")
def export(payload: dict[str, Any]) -> FileResponse:
    codes = payload.get("codes", [])
    if not isinstance(codes, list):
        raise HTTPException(status_code=422, detail="Trường codes phải là một danh sách.")
    if not codes:
        raise HTTPException(status_code=422, detail="Không có mã đề để xuất.")

    temp = tempfile.NamedTemporaryFile(prefix="ymp-answer-key-", suffix=".xlsx", delete=False)
    temp_path = temp.name
    temp.close()
    try:
        export_answer_key(codes, temp_path)
    except Exception as exc:
        Path(temp_path).unlink(missing_ok=True)
        logger.exception("Excel export failed")
        raise HTTPException(status_code=500, detail="Không thể tạo tệp Excel từ dữ liệu đã gửi.") from exc
    return FileResponse(
        temp_path,
        filename="young-mix-pro-answer-key.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        background=BackgroundTask(Path(temp_path).unlink, missing_ok=True),
    )
