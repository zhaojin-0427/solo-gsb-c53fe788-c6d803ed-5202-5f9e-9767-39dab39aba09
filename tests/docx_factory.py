"""测试用 DOCX 构造工具：直接拼装 word/*.xml 部件。"""

from __future__ import annotations

import io
import zipfile

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'


def make_docx(document: str, styles: str = None, numbering: str = None) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("word/document.xml", XML_DECL + document)
        if styles is not None:
            z.writestr("word/styles.xml", XML_DECL + styles)
        if numbering is not None:
            z.writestr("word/numbering.xml", XML_DECL + numbering)
    return buf.getvalue()


def styles_xml(*style_defs: str) -> str:
    return f'<w:styles {W_NS}>{"".join(style_defs)}</w:styles>'


def heading_style(sid: str, num_id: int = None, ilvl: int = 0, based_on: str = None) -> str:
    """内置标题样式定义；num_id 为 None 时不带编号。"""
    parts = [f'<w:name w:val="heading {sid[-1]}"/>']
    if based_on:
        parts.append(f'<w:basedOn w:val="{based_on}"/>')
    if num_id is not None:
        parts.append(
            f"<w:pPr><w:numPr><w:ilvl w:val=\"{ilvl}\"/>"
            f'<w:numId w:val="{num_id}"/></w:numPr></w:pPr>'
        )
    return (
        f'<w:style w:type="paragraph" w:styleId="{sid}">{"".join(parts)}</w:style>'
    )


def custom_style(sid: str, name: str, based_on: str) -> str:
    return (
        f'<w:style w:type="paragraph" w:styleId="{sid}">'
        f'<w:name w:val="{name}"/><w:basedOn w:val="{based_on}"/>'
        f"</w:style>"
    )


def numbering_xml(abstracts: str, nums: str) -> str:
    return f"<w:numbering {W_NS}>{abstracts}{nums}</w:numbering>"


def abstract(aid: int, levels: list) -> str:
    """levels: (ilvl, fmt, lvlText, start=None, restart=None, pstyle=None)"""
    lvls = []
    for spec in levels:
        ilvl, fmt, text = spec[0], spec[1], spec[2]
        start = spec[3] if len(spec) > 3 else None
        restart = spec[4] if len(spec) > 4 else None
        pstyle = spec[5] if len(spec) > 5 else None
        parts = []
        if start is not None:
            parts.append(f'<w:start w:val="{start}"/>')
        parts.append(f'<w:numFmt w:val="{fmt}"/>')
        if pstyle:
            parts.append(f'<w:pStyle w:val="{pstyle}"/>')
        parts.append(f'<w:lvlText w:val="{text}"/>')
        if restart is not None:
            parts.append(f'<w:lvlRestart w:val="{restart}"/>')
        lvls.append(f'<w:lvl w:ilvl="{ilvl}">{"".join(parts)}</w:lvl>')
    return f'<w:abstractNum w:abstractNumId="{aid}">{"".join(lvls)}</w:abstractNum>'


def num(num_id: int, abstract_id: int, overrides: str = "") -> str:
    return (
        f'<w:num w:numId="{num_id}">'
        f'<w:abstractNumId w:val="{abstract_id}"/>{overrides}</w:num>'
    )


def lvl_override(ilvl: int, start_override: int = None, lvl: str = None) -> str:
    parts = []
    if start_override is not None:
        parts.append(f'<w:startOverride w:val="{start_override}"/>')
    if lvl:
        parts.append(lvl)
    return f'<w:lvlOverride w:ilvl="{ilvl}">{"".join(parts)}</w:lvlOverride>'


# ---------------------------------------------------------------- 段落构造


def run(text: str) -> str:
    return f"<w:r><w:t>{text}</w:t></w:r>"


def ppr_xml(style: str = None, num_id: int = None, ilvl: int = 0, mark: str = None) -> str:
    parts = []
    if style:
        parts.append(f'<w:pStyle w:val="{style}"/>')
    if num_id is not None:
        parts.append(
            f'<w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr>'
        )
    if mark:
        parts.append(f'<w:rPr><w:{mark} w:id="1" w:author="t" w:date="2026-01-01"/></w:rPr>')
    return f"<w:pPr>{''.join(parts)}</w:pPr>" if parts else ""


def para(*runs: str, style: str = None, num_id: int = None, ilvl: int = 0, mark: str = None) -> str:
    return f"<w:p>{ppr_xml(style, num_id, ilvl, mark)}{''.join(runs)}</w:p>"


def heading(text: str, sid: str, bm: str = None, bm_id: int = 1) -> str:
    start = f'<w:bookmarkStart w:id="{bm_id}" w:name="{bm}"/>' if bm else ""
    end = f'<w:bookmarkEnd w:id="{bm_id}"/>' if bm else ""
    return para(f"{start}{run(text)}{end}", style=sid)


def ref_para(bookmark: str, *switches: str, current: str = "", style: str = None) -> str:
    sw = " ".join(f"\\{s}" for s in switches)
    instr = f" REF {bookmark} {sw} ".strip()
    return para(
        run("详见"),
        f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
        f'<w:r><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>',
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
        run(current),
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
        style=style,
    )


def doc(*body: str) -> str:
    return f"<w:document {W_NS}><w:body>{''.join(body)}</w:body></w:document>"
