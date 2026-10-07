"""审查编排：构建双视图、计算标题编号、核对 REF 域并生成报告。"""

from __future__ import annotations

import re

from .errors import AuditError
from .numbering import NO_NUMBER_FMTS, Numbering, NumberingEngine
from .oxml import (
    W,
    check_duplicate_bookmarks,
    load_parts,
    scan_revisions,
    wval,
)
from .styles import Styles
from .views import build_view, collect_groups, parse_fields

_REF_RE = re.compile(
    r'^\s*REF\s+(?:"(?P<q>[^"]+)"|(?P<b>[^\s\\]+))(?P<rest>.*)$',
    re.IGNORECASE | re.DOTALL,
)
_SWITCH_RE = re.compile(r"\\([A-Za-z])")

_PREVIEW_LEN = 80


def _preview(text: str) -> str:
    return text if len(text) <= _PREVIEW_LEN else text[: _PREVIEW_LEN - 1] + "…"


def _pstyle_of(ppr):
    if ppr is None:
        return None
    return wval(ppr.find(W + "pStyle"))


def _resolve_numref(ppr, sid, styles: Styles, numbering: Numbering):
    """按 Word 优先级解析段落编号来源：直接 numPr → 样式继承 → pStyle 链接。"""
    if ppr is not None:
        numpr = ppr.find(W + "numPr")
        if numpr is not None:
            num_id = wval(numpr.find(W + "numId"))
            if num_id == "0":
                return None  # 显式取消编号
            if num_id:
                ilvl_text = wval(numpr.find(W + "ilvl"))
                ilvl = int(ilvl_text) if ilvl_text is not None else 0
                return (num_id, ilvl) if numbering.has_num(num_id) else None
    inherited = styles.style_numpr(sid)
    if inherited == "disabled":
        return None
    if inherited is not None:
        num_id, ilvl = inherited
        return (num_id, ilvl) if numbering.has_num(num_id) else None
    for cur in styles.chain(sid):
        if cur in numbering.pstyle_map:
            return numbering.pstyle_map[cur]
    return None


def _relative_text(ctx_raw, ctx_fmt, t_raw, t_fmt):
    """相对编号：去掉与引用点标题相同的公共前缀；目标为上下文前缀时回退完整编号。"""
    i = 0
    while i < len(ctx_raw) and i < len(t_raw) and ctx_raw[i] == t_raw[i]:
        i += 1
    if i == len(t_raw):
        i = 0  # 目标编号是上下文编号的前缀（或相同），无法相对化，取完整编号
    return ".".join(f for f in t_fmt[i:] if f != "")


def _audit_view(paras, styles: Styles, numbering: Numbering) -> dict:
    # ---- 第一遍：标题编号 ------------------------------------------------
    engine = NumberingEngine(numbering)
    infos = []
    for para in paras:
        sid = _pstyle_of(para.ppr)
        info = {"hlevel": styles.heading_level(sid), "numbered": False}
        numref = _resolve_numref(para.ppr, sid, styles, numbering)
        if numref is not None:
            num_id, ilvl = numref
            if numbering.fmt(num_id, ilvl) not in NO_NUMBER_FMTS:
                raw, fmt = engine.advance(num_id, ilvl)
                info.update(
                    numbered=True,
                    raw=raw,
                    fmt=fmt,
                    text=".".join(x for x in fmt if x != ""),
                )
        infos.append(info)

    # ---- 第二遍：书签、域、段落文本 --------------------------------------
    bookmarks: dict = {}  # name -> {para, text}
    open_bm: dict = {}  # id -> {name, para, parts}
    para_fields: list = []
    para_texts: list = []
    for idx, para in enumerate(paras):
        fields = parse_fields(para.items)  # 结构损坏在此抛 422
        para_fields.append(fields)
        parts = []
        for it in para.items:
            if it.kind == "text":
                parts.append(it.text)
                for ob in open_bm.values():
                    ob["parts"].append(it.text)
            elif it.kind == "bm_start":
                open_bm[it.bm_id] = {"name": it.name, "para": idx, "parts": []}
            elif it.kind == "bm_end":
                ob = open_bm.pop(it.bm_id, None)
                if ob is not None and ob["name"] not in bookmarks:
                    bookmarks[ob["name"]] = {
                        "para": ob["para"],
                        "text": "".join(ob["parts"]),
                    }
        para_texts.append("".join(parts))
    for ob in open_bm.values():  # 未闭合书签：按空书签处理
        if ob["name"] not in bookmarks:
            bookmarks[ob["name"]] = {"para": ob["para"], "text": ""}

    # ---- 第三遍：核对 REF 域 ---------------------------------------------
    headings = [(i, inf) for i, inf in enumerate(infos) if inf["hlevel"] and inf["numbered"]]

    def context_for(idx):
        ctx = None
        for i, inf in headings:
            if i <= idx:
                ctx = inf
            else:
                break
        return ctx

    out_fields = []
    for idx, fields in enumerate(para_fields):
        for fidx, fld in enumerate(fields):
            m = _REF_RE.match(fld.instr)
            if not m:
                continue  # 仅核对 REF 域
            name = m.group("q") if m.group("q") is not None else m.group("b")
            switches = _SWITCH_RE.findall(m.group("rest") or "")
            mode = None
            for s in switches:
                if s in ("n", "r", "w"):
                    mode = s  # 多个开关时后者优先
            current = fld.result if fld.has_sep else ""
            bm = bookmarks.get(name)
            entry = {
                "paragraph_index": idx,
                "field_index": fidx,
                "paragraph_text": _preview(para_texts[idx]),
                "instruction": fld.instr.strip(),
                "bookmark": name,
                "switches": switches,
                "mode": mode,
                "current": current,
                "expected": None,
                "status": "target_missing",
            }
            if bm is not None:
                tinfo = infos[bm["para"]]
                if mode is None:
                    # 无 n/r/w 开关：REF 取书签文本
                    entry["expected"] = bm["text"]
                elif not tinfo["numbered"]:
                    entry["expected"] = None  # 目标不是可见的编号标题
                elif mode == "n":
                    entry["expected"] = tinfo["fmt"][-1]
                elif mode == "w":
                    entry["expected"] = tinfo["text"]
                else:  # mode == "r"
                    ctx = context_for(idx)
                    if ctx is None:
                        raise AuditError(
                            "missing_r_context",
                            f"第 {idx} 段的 REF \\r 域（书签 {name!r}）之前没有可见的编号标题",
                        )
                    entry["expected"] = _relative_text(
                        ctx["raw"], ctx["fmt"], tinfo["raw"], tinfo["fmt"]
                    )
                if entry["expected"] is not None:
                    entry["status"] = (
                        "correct" if entry["expected"] == current else "stale"
                    )
            out_fields.append(entry)
    return {"field_count": len(out_fields), "fields": out_fields}


def audit_docx(data: bytes) -> dict:
    """审查入口：返回双视图报告；任何可预期错误抛 AuditError（422）。"""
    document, styles_root, numbering_root = load_parts(data)
    scan_revisions(document)
    check_duplicate_bookmarks(document)

    styles = Styles(styles_root)
    numbering = Numbering(numbering_root)

    body = document.find(W + "body")
    if body is None:
        raise AuditError("invalid_docx", "word/document.xml 缺少 w:body")

    groups = collect_groups(body)
    return {
        "views": {
            "accept": _audit_view(build_view(groups, "accept"), styles, numbering),
            "reject": _audit_view(build_view(groups, "reject"), styles, numbering),
        }
    }
