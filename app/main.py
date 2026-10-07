"""FastAPI application exposing the DOCX cross-reference audit service."""

from __future__ import annotations

from fastapi import FastAPI, File, UploadFile
from fastapi.responses import JSONResponse

from .docx_audit.auditor import audit_document
from .docx_audit.errors import AuditError

app = FastAPI(
    title="DOCX 条款交叉引用审查 API",
    description=(
        "接收未加密 DOCX，核对正文中指向自动编号标题（Heading 1~3）的 REF "
        "编号域（支持 \\n、\\r、\\w 开关），分别给出接受修订与拒绝修订视图。"
    ),
    version="1.0.0",
)

MAX_UPLOAD_BYTES = 50 * 1024 * 1024


@app.exception_handler(AuditError)
async def audit_error_handler(request, exc: AuditError):
    return JSONResponse(status_code=422, content=exc.detail())


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.post("/audit")
async def audit(file: UploadFile = File(...)):
    if not (
        file.filename is None
        or file.filename.lower().endswith(".docx")
        or file.content_type
        in (
            None,
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "application/octet-stream",
        )
    ):
        return JSONResponse(
            status_code=422,
            content={
                "error": {
                    "code": "not_a_docx",
                    "message": "仅支持 .docx 文件（不支持旧版 .doc 或加密文档）",
                }
            },
        )

    data = await file.read()
    if len(data) > MAX_UPLOAD_BYTES:
        return JSONResponse(
            status_code=422,
            content={"error": {"code": "file_too_large", "message": "文件超过 50 MiB"}},
        )

    report = audit_document(data)
    return {
        "filename": file.filename,
        "accept_view": report["accept"],
        "reject_view": report["reject"],
    }
