"""Tracked-revision validation.

Only the two revision classes named by the service contract may appear in
the audited parts:

* run-content insertions/deletions (``w:ins`` / ``w:del`` wrapping runs);
* paragraph-mark insertions/deletions (``w:ins`` / ``w:del`` inside
  ``w:pPr/rPr``).

``w:rPrChange`` and ``w:pPrChange`` (formatting revisions) are the only
other tolerated marks.  Everything else -- moves, table/numbering/cell
changes, comments, merge fields and any further revision type -- aborts the
request with HTTP 422.
"""

from __future__ import annotations

from .errors import AuditError
from .namespaces import w

# Revision marks explicitly outside the supported set.
_FORBIDDEN = {
    "moveFrom",
    "moveTo",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
    "cellIns",
    "cellDel",
    "cellMerge",
    "trIns",
    "trDel",
    "trPrChange",
    "tblPrChange",
    "tblPrExChange",
    "tcPrChange",
    "tblGridChange",
    "numberingChange",
    "customXmlInsRangeStart",
    "customXmlInsRangeEnd",
    "customXmlDelRangeStart",
    "customXmlDelRangeEnd",
    "customXmlMoveFromRangeStart",
    "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart",
    "customXmlMoveToRangeEnd",
    "commentRangeStart",
    "commentRangeEnd",
    "commentReference",
    "commentContentStart",
    "commentContentEnd",
    "commentEx",
    "moveFrom",
    "pPrChange",
    "rPrChange",
    "sectPrChange",
    "footnoteRef",
}

_TOLERATED_CHANGE = {"pPrChange", "rPrChange"}

_ALLOWED_INSIDE_CHANGE = {"r", "bookmarkStart", "bookmarkEnd", "proofErr"}

# Parents that may legitimately contain a run-level w:ins / w:del, including
# block-level containers (whole paragraph inserted/deleted) and inline flow
# containers.
_RUN_CHANGE_PARENTS = {
    "p",
    "hyperlink",
    "smartTag",
    "sdtContent",
    "body",
    "tc",
}


def validate_no_unsupported_revisions(*roots):
    """Reject any revision type beyond text/pilcrow inserts and deletes."""
    for root in roots:
        if root is None:
            continue
        for el in root.iter():
            if not el.tag.startswith("{"):
                continue
            local = el.tag.rsplit("}", 1)[-1]
            if local in _FORBIDDEN and local not in _TOLERATED_CHANGE:
                raise AuditError(
                    "unsupported_revision",
                    f"文档包含本服务不支持的修订类型 w:{local}；"
                    "仅支持文字插入/删除与段落标记插入/删除修订",
                )
            if local in ("ins", "del"):
                _validate_change_wrapper(el)


def _validate_change_wrapper(el):
    parent = el.getparent()
    parent_local = parent.tag.rsplit("}", 1)[-1] if parent is not None else ""

    # Pilcrow revision: w:pPr/w:rPr/w:ins|w:del.
    if parent_local == "rPr":
        rpr_parent = parent.getparent()
        if rpr_parent is None or rpr_parent.tag != w("pPr"):
            raise AuditError(
                "unsupported_revision",
                f"w:{el.tag.rsplit('}', 1)[-1]} 出现在不允许的位置",
            )
        return

    # Run-content revision: must live in a paragraph flow container.
    if parent_local not in _RUN_CHANGE_PARENTS:
        raise AuditError(
            "unsupported_revision",
            f"w:{el.tag.rsplit('}', 1)[-1]} 出现在块级或其他不允许的位置",
        )

    # Nested revision wrappers (ins inside ins/del etc.) are not produced by
    # Word and have no unambiguous accept/reject semantics.
    ancestor = parent
    while ancestor is not None:
        if ancestor.tag in (w("ins"), w("del")):
            raise AuditError(
                "unsupported_revision",
                "检测到嵌套的 w:ins/w:del 修订容器，无法生成一致视图",
            )
        ancestor = ancestor.getparent()

    for child in el:
        local = child.tag.rsplit("}", 1)[-1] if child.tag.startswith("{") else ""
        if local not in _ALLOWED_INSIDE_CHANGE:
            raise AuditError(
                "unsupported_revision",
                f"修订容器 w:{el.tag.rsplit('}', 1)[-1]} 内出现不支持的元素 w:{local}",
            )
