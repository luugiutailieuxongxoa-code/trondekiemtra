from io import BytesIO
from pathlib import Path
import sys

from docx import Document
from fastapi.testclient import TestClient

BACKEND_DIR = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(BACKEND_DIR))
from main import app  # noqa: E402

client = TestClient(app)


def make_docx() -> bytes:
    doc = Document()
    doc.add_paragraph("Câu 1. Nước có công thức là gì?")
    doc.add_paragraph("A. H2O")
    doc.add_paragraph("B. CO2")
    doc.add_paragraph("C. O2")
    doc.add_paragraph("D. NaCl")
    buffer = BytesIO()
    doc.save(buffer)
    return buffer.getvalue()


def test_health_endpoint():
    response = client.get("/api/health")
    assert response.status_code == 200
    assert response.json()["ok"] is True
    assert response.json()["service"] == "young-mix-pro-api"


def test_config_does_not_expose_secrets():
    response = client.get("/api/config")
    assert response.status_code == 200
    assert response.json()["features"]["pdf_import"] is False
    assert "GEMINI_API_KEY" not in response.text


def test_docx_parse_endpoint():
    response = client.post(
        "/api/exams/parse",
        files={"file": ("sample.docx", make_docx(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")},
    )
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["question_count"] == 1
    assert len(body["questions"][0]["options"]) == 4
    assert body["review_required"] is True


def test_pdf_is_rejected_with_clear_message():
    response = client.post("/api/exams/parse", files={"file": ("sample.pdf", b"%PDF", "application/pdf")})
    assert response.status_code == 415
    assert "PDF/OCR" in response.json()["detail"]


def test_mix_rejects_empty_question_list():
    response = client.post("/api/exams/mix", json={"questions": [], "mode": "g3", "code_count": 2})
    assert response.status_code == 422
