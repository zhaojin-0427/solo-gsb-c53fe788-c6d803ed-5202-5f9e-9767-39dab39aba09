"""业务错误：一律映射为 HTTP 422，且不返回任何部分报告。"""

from __future__ import annotations


class AuditError(Exception):
    """审查过程中的可预期错误。

    code 用于机器可读的错误类别，message 面向调用者描述原因。
    """

    def __init__(self, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message
