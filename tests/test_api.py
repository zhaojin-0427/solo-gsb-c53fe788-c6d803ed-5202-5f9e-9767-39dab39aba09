"""HTTP-level tests for the FastAPI service."""
import io
import sys
import os

import pytest
from fastapi.testclient import TestClient

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
sys.path.insert(0, os.path.dirname(__file__))

from app.main import app
from builder import DocxBuilder
from docx_audit.errors import AuditError

client = TestClient(app)


def _docx_bytes():
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1)
    b.heading("目标标题", level=1, bookmark="h")
    b.paragraph(*b.ref_field_runs("h", "w", "1"))
    return b.build()


def test_health():
    r = client.get("/health")
    assert r.status_code == 200
    assert r.json() == {"status": "ok"}


def test_audit_success_envelope():
    r = client.post(
        "/audit",
        files={"file": ("sample.docx", _docx_bytes(),
                        "application/vnd.openxmlformats-officedocument."
                        "wordprocessingml.document")},
    )
    assert r.status_code == 200
    body = r.json()
    assert body["filename"] == "sample.docx"
    assert set(body) == {"filename", "accept_view", "reject_view"}
    f = body["accept_view"]["fields"][0]
    assert f["status"] == "correct"
    assert f["current_value"] == "1"
    assert f["expected_value"] == "1"
    assert f["switch"] == "w"
    assert f["bookmark"] == "h"
    assert {"paragraph_index", "field_index", "xpath", "table_path",
            "paragraph_text", "logical_paragraph_index"} == set(f["position"])


def test_audit_rejects_non_zip_with_422():
    r = client.post("/audit", files={"file": ("a.docx", b"not a zip",
                                              "application/octet-stream")})
    assert r.status_code == 422
    assert r.json()["error"]["code"] == "not_a_docx"


def test_audit_rejects_doc_extension():
    r = client.post("/audit", files={"file": ("old.doc", b"x",
                                              "application/octet-stream")})
    assert r.status_code == 422
    assert r.json()["error"]["code"] == "not_a_docx"


def test_duplicate_bookmark_is_422_no_partial_report():
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1)
    b.heading("一", level=1, bookmark="dup")
    b.heading("二", level=1, bookmark="dup")
    r = client.post("/audit", files={"file": ("d.docx", b.build(),
                                              "application/octet-stream")})
    assert r.status_code == 422
    body = r.json()
    assert body["error"]["code"] == "duplicate_bookmark"
    assert "accept_view" not in body and "reject_view" not in body


def test_no_r_context_is_422():
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1)
    b.paragraph(*b.ref_field_runs("h", "r", "1"))
    b.heading("h", level=1, bookmark="h")
    r = client.post("/audit", files={"file": ("d.docx", b.build(),
                                              "application/octet-stream")})
    assert r.status_code == 422
    assert r.json()["error"]["code"] == "no_r_context"


def test_encrypted_docx_422():
    # Construct an OOXML package that declares the encrypted content type.
    import zipfile
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr(
            "[Content_Types].xml",
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/'
            'content-types"><Default Extension="rels" ContentType="x"/>'
            '<Override PartName="/word/document.xml" ContentType='
            '"application/vnd.openxmlformats-officedocument.wordprocessingml'
            '.document.encrypted"/></Types>',
        )
        zf.writestr("word/document.xml",
                    '<w:document xmlns:w="http://schemas.openxmlformats.org/'
                    'wordprocessingml/2006/main"/>')
    r = client.post("/audit", files={"file": ("e.docx", buf.getvalue(),
                                              "application/octet-stream")})
    assert r.status_code == 422
    assert r.json()["error"]["code"] == "encrypted_docx"
