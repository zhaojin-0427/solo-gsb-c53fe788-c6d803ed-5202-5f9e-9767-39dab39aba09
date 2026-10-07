"""Audit error model.

Every structural or semantic rejection defined by the service contract raises
:class:`AuditError`, which FastAPI maps to HTTP 422 with no partial report.
"""


class AuditError(Exception):
    """Raised when a document cannot be audited under the service contract."""

    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message

    def detail(self) -> dict:
        return {"error": {"code": self.code, "message": self.message}}
