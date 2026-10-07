"""FastAPI 入口：POST /audit 接收未加密 DOCX，返回双视图交叉引用审查报告。"""

from __future__ import annotations

from fastapi import FastAPI, File, Request, UploadFile
from fastapi.responses import JSONResponse

from .audit import audit_docx
from .errors import AuditError

app = FastAPI(
    title="DOCX 条款交叉引用审查 API",
    version="1.0.0",
    description=(
        "核对 DOCX 正文及表格中带书签的 Heading 1~3 自动编号标题与指向它们的 "
        "REF 编号域（\\n、\\r、\\w 开关），分别给出接受全部修订与拒绝全部修订"
        "两个视图下每处域的位置、现值、应值与状态。"
    ),
)


@app.exception_handler(AuditError)
async def audit_error_handler(_request, exc: AuditError):
    # 可预期的文档问题一律 422，且不返回任何部分报告
    return JSONResponse(
        status_code=422,
        content={"detail": {"error": exc.code, "message": exc.message}},
    )


@app.get("/")
async def index():
    return {
        "service": "DOCX 条款交叉引用审查 API",
        "usage": "POST /audit（multipart 表单字段 file，或直接以请求体提交 DOCX 字节）",
        "docs": "/docs",
    }


@app.post("/audit")
async def audit(request: Request, file: UploadFile | None = File(default=None)):
    data = await file.read() if file is not None else await request.body()
    return audit_docx(data)
