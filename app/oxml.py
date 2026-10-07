"""OOXML 基础层：命名空间、DOCX 部件加载、修订类型与书签的预先校验。"""

from __future__ import annotations

import io
import zipfile

from lxml import etree

from .errors import AuditError

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = "{%s}" % W_NS

# 除“文字插入/删除、段落标记插入/删除”之外的修订类型，一律拒绝（422）。
UNSUPPORTED_REVISION_TAGS = {
    "moveFrom",
    "moveTo",
    "moveFromRangeStart",
    "moveFromRangeEnd",
    "moveToRangeStart",
    "moveToRangeEnd",
    "pPrChange",
    "rPrChange",
    "tblPrChange",
    "trPrChange",
    "tcPrChange",
    "tblGridChange",
    "sectPrChange",
    "numberingChange",
    "cellIns",
    "cellDel",
    "cellMerge",
    "customXmlInsRangeStart",
    "customXmlInsRangeEnd",
    "customXmlDelRangeStart",
    "customXmlDelRangeEnd",
    "customXmlMoveFromRangeStart",
    "customXmlMoveFromRangeEnd",
    "customXmlMoveToRangeStart",
    "customXmlMoveToRangeEnd",
}

# w:ins / w:del 作为“运行级”修订时允许出现的父元素。
_RUN_CONTAINER_PARENTS = {
    "p",
    "hyperlink",
    "smartTag",
    "sdtContent",
    "customXml",
    "fldSimple",
    "bdo",
    "dir",
}


def localname(el) -> str:
    return etree.QName(el).localname


def wval(el, attr: str = "val"):
    """读取 w:<attr> 属性；元素为 None 时返回 None。"""
    if el is None:
        return None
    return el.get(W + attr)


def load_parts(data: bytes):
    """解包 DOCX，返回 (document, styles, numbering) 三棵 XML 树（后两者可为 None）。"""
    if not data:
        raise AuditError("invalid_docx", "请求体为空，未收到 DOCX 文件")
    try:
        zf = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile:
        raise AuditError(
            "invalid_docx",
            "无法按 ZIP 解包：文件已加密或不是有效的 DOCX",
        )
    names = set(zf.namelist())
    if "word/document.xml" not in names:
        raise AuditError("invalid_docx", "包内缺少 word/document.xml，不是有效的 DOCX")

    parser = etree.XMLParser(resolve_entities=False, no_network=True)

    def parse(name: str):
        try:
            return etree.fromstring(zf.read(name), parser)
        except (etree.XMLSyntaxError, ValueError) as exc:
            raise AuditError("invalid_docx", f"XML 解析失败（{name}）：{exc}")

    document = parse("word/document.xml")
    styles = parse("word/styles.xml") if "word/styles.xml" in names else None
    numbering = parse("word/numbering.xml") if "word/numbering.xml" in names else None
    return document, styles, numbering


def scan_revisions(document) -> None:
    """仅允许文字/段落标记的插入与删除，其余修订类型一律 422。"""
    for el in document.iter():
        if not isinstance(el.tag, str):
            continue
        name = localname(el)
        if name in UNSUPPORTED_REVISION_TAGS:
            raise AuditError(
                "unsupported_revision",
                f"检测到不支持的修订类型 w:{name}，仅支持文字与段落标记的插入/删除",
            )
        if name in ("ins", "del"):
            parent = el.getparent()
            pname = localname(parent) if parent is not None else ""
            if pname in _RUN_CONTAINER_PARENTS:
                continue  # 运行级插入/删除
            if pname == "rPr":
                grand = parent.getparent()
                if grand is not None and localname(grand) == "pPr":
                    continue  # 段落标记插入/删除
            raise AuditError(
                "unsupported_revision",
                f"w:{name} 出现在不支持的位置（父元素 w:{pname}）",
            )


def check_duplicate_bookmarks(document) -> None:
    """同一书签名出现两次即拒绝（无论是否处于修订内容中）。"""
    seen: set[str] = set()
    for el in document.iter(W + "bookmarkStart"):
        name = el.get(W + "name")
        if name is None:
            continue
        if name in seen:
            raise AuditError("duplicate_bookmark", f"书签名重复：{name!r}")
        seen.add(name)
