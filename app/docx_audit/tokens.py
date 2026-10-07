"""Per-paragraph tokenizer: extracts runs, fields and bookmarks.

Word fields can be spread across many runs (``w:fldChar`` begin / separate /
end plus ``w:instrText`` instructions), or expressed compactly as
``w:fldSimple``.  Both forms are supported, and every part of one field must
be enclosed by the same revision wrapper (or none) -- otherwise the document
is rejected as structurally corrupt.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .errors import AuditError
from .namespaces import w


@dataclass
class FieldSpan:
    ordinal: int
    instr: str = ""
    cached: str = ""
    simple: bool = False
    wrapper: object = None  # the enclosing w:ins / w:del element, or None


@dataclass
class BookmarkEvent:
    name: str
    kind: str  # "start" | "end"
    wrapper: object = None


@dataclass
class _OpenField:
    span: FieldSpan
    stage: str = "instr"  # instr -> result -> closed


@dataclass
class ParagraphTokens:
    fields: list = field(default_factory=list)
    bookmarks: list = field(default_factory=list)
    # Concatenated visible text per revision view key: None / "ins" / "del".
    text_by_wrapper: dict = field(default_factory=dict)
    # w:ins / w:del elements that wrap run-level content.
    content_wrappers: list = field(default_factory=list)
    pilcrow: object = None  # w:ins / w:del attached to the paragraph mark
    p_pr: object = None


class _Tokenizer:
    def __init__(self, p):
        self.p = p
        self.tokens = ParagraphTokens(p_pr=p.find(w("pPr")))
        self.stack: list[_OpenField] = []
        self._field_ordinal = 0

    def run(self):
        self._read_pilcrow()
        self._walk(self.p, wrapper=None, depth=0)
        if self.stack:
            raise AuditError(
                "malformed_field",
                "文档中的域缺少 w:fldChar end，域结构损坏",
            )
        return self.tokens

    # ------------------------------------------------------------- pilcrow
    def _read_pilcrow(self):
        p_pr = self.tokens.p_pr
        if p_pr is None:
            return
        rpr = p_pr.find(w("rPr"))
        if rpr is None:
            return
        ins = rpr.find(w("ins"))
        dele = rpr.find(w("del"))
        if ins is not None and dele is not None:
            raise AuditError(
                "conflicting_pilcrow_revision",
                "同一段落标记同时存在插入与删除修订，无法生成一致的接受/拒绝视图",
            )
        self.tokens.pilcrow = ins if ins is not None else dele

    # -------------------------------------------------------------- walker
    def _walk(self, el, wrapper, depth):
        for child in el:
            tag = child.tag
            if not isinstance(tag, str) or not tag.startswith("{"):
                continue
            local = tag.rsplit("}", 1)[-1]

            if local == "pPr":
                continue  # handled separately
            if local in ("ins", "del"):
                if wrapper is None:
                    self.tokens.content_wrappers.append(child)
                self._walk(child, wrapper=child, depth=depth + 1)
                continue
            if local == "r":
                self._run(child, wrapper)
                continue
            if local == "fldSimple":
                self._simple_field(child, wrapper)
                continue
            if local == "bookmarkStart":
                self.tokens.bookmarks.append(
                    BookmarkEvent(name=child.get(w("name"), ""), kind="start", wrapper=wrapper)
                )
                continue
            if local == "bookmarkEnd":
                self.tokens.bookmarks.append(
                    BookmarkEvent(name=child.get(w("name"), ""), kind="end", wrapper=wrapper)
                )
                continue
            # Structural containers (hyperlink, smartTag, sdtContent ...):
            # descend without changing the revision context.
            self._walk(child, wrapper=wrapper, depth=depth + 1)

    # --------------------------------------------------------------- runs
    def _run(self, r, wrapper):
        for child in r:
            tag = child.tag
            if not isinstance(tag, str) or not tag.startswith("{"):
                continue
            local = tag.rsplit("}", 1)[-1]

            if local == "fldChar":
                fld_type = child.get(w("fldCharType"))
                if fld_type == "begin":
                    span = FieldSpan(ordinal=self._field_ordinal, wrapper=wrapper)
                    self._field_ordinal += 1
                    self.stack.append(_OpenField(span=span))
                elif fld_type == "separate":
                    if not self.stack:
                        raise AuditError("malformed_field", "孤立的域分隔符，域结构损坏")
                    self.stack[-1].stage = "result"
                elif fld_type == "end":
                    if not self.stack:
                        raise AuditError("malformed_field", "孤立的域结束符，域结构损坏")
                    open_field = self.stack.pop()
                    self.tokens.fields.append(open_field.span)
                else:
                    # beginSeparate / endSeparate inside legacy field
                    # constructions are not supported.
                    raise AuditError(
                        "malformed_field", f"不支持的 fldCharType: {fld_type}"
                    )
                continue

            if local == "instrText":
                if not self.stack:
                    raise AuditError(
                        "malformed_field", "域指令出现在 w:fldChar begin 之外"
                    )
                open_field = self.stack[-1]
                if open_field.span.wrapper != wrapper:
                    raise AuditError(
                        "malformed_field",
                        "域跨越了不同的修订容器，域结构损坏",
                    )
                open_field.span.instr += child.text or ""
                continue

            text = self._leaf_text(local, child)
            if text is not None:
                self._append_text(wrapper, text)
                for open_field in self.stack:
                    if open_field.span.wrapper != wrapper:
                        raise AuditError(
                            "malformed_field",
                            "域跨越了不同的修订容器，域结构损坏",
                        )
                    if open_field.stage == "instr":
                        if local != "t":
                            # literal text between begin and separate must be
                            # instrText only
                            raise AuditError(
                                "malformed_field",
                                "域指令区中出现普通文本，域结构损坏",
                            )
                    else:
                        open_field.span.cached += text

    def _simple_field(self, fld, wrapper):
        span = FieldSpan(
            ordinal=self._field_ordinal,
            instr=fld.get(w("instr"), ""),
            simple=True,
            wrapper=wrapper,
        )
        self._field_ordinal += 1
        # Descend with the simple field implicitly open so its result text
        # lands in `cached`.
        open_field = _OpenField(span=span, stage="result")
        before = len(self.stack)
        self.stack.append(open_field)
        self._walk(fld, wrapper=wrapper, depth=1)
        if self.stack[before] is not open_field:
            raise AuditError(
                "malformed_field", "fldSimple 内嵌套的复杂域结构损坏"
            )
        self.stack.pop()
        self.tokens.fields.append(span)

    @staticmethod
    def _leaf_text(local, child):
        if local == "t":
            return child.text or ""
        if local == "delText":
            return child.text or ""
        if local == "tab":
            return "\t"
        if local in ("br", "cr"):
            return "\n"
        if local == "noBreakHyphen":
            return "‑"
        if local == "softHyphen":
            return ""
        return None

    def _append_text(self, wrapper, text):
        key = self._wrapper_key(wrapper)
        self.tokens.text_by_wrapper[key] = (
            self.tokens.text_by_wrapper.get(key, "") + text
        )

    @staticmethod
    def _wrapper_key(wrapper):
        if wrapper is None:
            return None
        return wrapper.tag  # w:ins or w:del


def tokenize_paragraph(p) -> ParagraphTokens:
    return _Tokenizer(p).run()
