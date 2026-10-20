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
import re
import zipfile
from io import BytesIO
from models import MixRequest

VERSION = "0.8.0"
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
                     "docx_export_source_preserving": True, "docx_section_aware_parse": True, "docx_source_mix_ui": True,
                     "docx_official_header": True, "docx_answer_mark_cleanup": True, "docx_official_footer": True, "docx_page_numbers": True, "short_answer_inline_key_cleanup": True, "section_g_modes": True, "option_paragraph_shuffle": True, "editable_docx_header": True, "pdf_export": False,
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
    header: str | None = None,
    section_modes: str | None = None,
    seed: int | None = None,
) -> FileResponse:
    """Export a DOCX using the source document, standard exam header and answer-mark cleanup.

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
        header_config = None
        if header:
            try:
                header_config = json.loads(header)
            except json.JSONDecodeError as exc:
                raise HTTPException(status_code=422, detail="header phải là JSON hợp lệ.") from exc
            if not isinstance(header_config, dict):
                raise HTTPException(status_code=422, detail="header phải là một đối tượng JSON.")
        section_mode_config = None
        if section_modes:
            try:
                section_mode_config = json.loads(section_modes)
            except json.JSONDecodeError as exc:
                raise HTTPException(status_code=422, detail="section_modes phải là JSON hợp lệ.") from exc
            if not isinstance(section_mode_config, dict):
                raise HTTPException(status_code=422, detail="section_modes phải là một đối tượng JSON.")
        try:
            exported, report = export_reordered_docx(data, order, header_config, section_mode_config, seed)
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
                "X-YMP-Answer-Markings-Removed": str(report["answer_markings_removed"]).lower(),
                "X-YMP-Answer-Key-Section-Removed": str(report["answer_key_section_removed"]).lower(),
                "X-YMP-Parser-Markers-Removed": str(report["parser_markers_removed"]).lower(),
                "X-YMP-Official-Header-Added": str(report["official_header_added"]).lower(),
                "X-YMP-Official-Footer-Added": str(report["official_footer_added"]).lower(),
                "X-YMP-Page-Number-Added": str(report["page_number_added"]).lower(),
                "X-YMP-Short-Answer-Keys-Removed": str(report["short_answer_inline_keys_removed"] > 0).lower(),
                "X-YMP-Section-Order-Preserved": "true",
                "X-YMP-Option-Paragraphs-Shuffled": str(report.get("option_paragraphs_shuffled", 0)),
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


@app.post("/api/exams/export-bundle")
async def export_bundle(
    file: UploadFile = File(...),
    question_order: str | None = None,
    header: str | None = None,
    section_modes: str | None = None,
    answers: str | None = None,
    exam_code: str = "001",
    seed: int | None = None,
    answer_layout: str = "vertical",
) -> FileResponse:
    """Return one ZIP containing the source-preserving Word exam and its Excel key.

    The exam code is user-editable and is used consistently in the Word header,
    Excel answer sheet and filenames. Answers are never inferred by this endpoint.
    """
    temp_path: str | None = None
    try:
        code = str(exam_code or "").strip()
        if not code or len(code) > 32 or not re.fullmatch(r"[A-Za-z0-9_-]+", code):
            raise HTTPException(status_code=422, detail="Mã đề chỉ được gồm chữ, số, dấu gạch ngang hoặc gạch dưới (tối đa 32 ký tự).")
        data = await read_limited_upload(file)

        def parse_json_param(value: str | None, name: str, expected: type):
            if not value:
                return None
            try:
                parsed = json.loads(value)
            except json.JSONDecodeError as exc:
                raise HTTPException(status_code=422, detail=f"{name} phải là JSON hợp lệ.") from exc
            if not isinstance(parsed, expected):
                raise HTTPException(status_code=422, detail=f"{name} có kiểu dữ liệu không hợp lệ.")
            return parsed

        order = parse_json_param(question_order, "question_order", list)
        if order is not None and not all(isinstance(item, str) for item in order):
            raise HTTPException(status_code=422, detail="question_order phải chứa ID câu hỏi dạng chuỗi.")
        header_config = parse_json_param(header, "header", dict) or {}
        header_config["exam_code"] = code
        modes = parse_json_param(section_modes, "section_modes", dict) or {}
        answer_rows = parse_json_param(answers, "answers", list)
        if not answer_rows:
            raise HTTPException(status_code=422, detail="Không có dữ liệu đáp án để xuất. Hãy tạo mã đề và kiểm tra đáp án trước.")
        normalized_answers = []
        for index, item in enumerate(answer_rows):
            if not isinstance(item, dict):
                raise HTTPException(status_code=422, detail="Mỗi đáp án phải là một đối tượng.")
            normalized_answers.append({
                "number": item.get("number", index + 1),
                "type": item.get("type", "multiple_choice"),
                "answer": item.get("answer") if item.get("answer") not in ("Chưa xác nhận", "") else None,
                "points": item.get("points", ""),
            })

        try:
            docx_bytes, report = export_reordered_docx(data, order, header_config, modes, seed)
        except ValueError as exc:
            raise HTTPException(status_code=422, detail=str(exc)) from exc

        with tempfile.TemporaryDirectory(prefix="ymp-bundle-") as tmp:
            tmpdir = Path(tmp)
            docx_path = tmpdir / f"De-thi-ma-{code}.docx"
            xlsx_path = tmpdir / f"Dap-an-ma-{code}.xlsx"
            zip_path = tmpdir / f"Young-Mix-Pro-ma-{code}.zip"
            docx_path.write_bytes(docx_bytes)
            export_answer_key([{"code": code, "questions": normalized_answers}], xlsx_path, layout=answer_layout)
            with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
                archive.write(docx_path, arcname=docx_path.name)
                archive.write(xlsx_path, arcname=xlsx_path.name)
                if report.get("warnings"):
                    warning_path = tmpdir / "CAN-KIEM-TRA.txt"
                    warning_path.write_text("Các cảnh báo khi xuất đề:\n" + "\n".join(f"- {w}" for w in report["warnings"]), encoding="utf-8")
                    archive.write(warning_path, arcname=warning_path.name)
            # Move out of temporary directory so FileResponse can stream it, then clean after response.
            handle = tempfile.NamedTemporaryFile(prefix="ymp-bundle-", suffix=".zip", delete=False)
            temp_path = handle.name
            handle.close()
            Path(temp_path).write_bytes(zip_path.read_bytes())
        return FileResponse(
            temp_path,
            filename=f"Young-Mix-Pro-ma-{code}.zip",
            media_type="application/zip",
            headers={
                "X-YMP-Exam-Code": code,
                "X-YMP-Question-Count": str(report.get("question_count", len(normalized_answers))),
                "X-YMP-Warning-Count": str(len(report.get("warnings", []))),
                "X-YMP-Answer-Layout": answer_layout,
            },
            background=BackgroundTask(Path(temp_path).unlink, missing_ok=True),
        )
    except HTTPException:
        if temp_path:
            Path(temp_path).unlink(missing_ok=True)
        raise
    except ValueError as exc:
        if temp_path:
            Path(temp_path).unlink(missing_ok=True)
        raise HTTPException(status_code=422, detail=str(exc)) from exc
    except Exception as exc:
        if temp_path:
            Path(temp_path).unlink(missing_ok=True)
        logger.exception("Bundle export failed")
        raise HTTPException(status_code=500, detail="Không thể tạo ZIP đề Word và đáp án Excel.") from exc
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
