"""视图构建：把 document.xml 展开为“接受全部修订 / 拒绝全部修订”两个视图。

- 文字修订：w:ins 仅出现于接受视图，w:del 仅出现于拒绝视图；
- 段落标记修订：w:pPr/w:rPr/w:del 的段落标记在接受视图中被删除，
  该段落与下一段合并（合并后取下一段的段落属性）；w:ins 的段落标记
  在拒绝视图中不存在，同样与下一段合并；
- 域（fldChar/instrText/fldSimple）与书签（bookmarkStart/End）跨 run
  出现时，按文档顺序的事件流进行配对。
"""

from __future__ import annotations

from dataclasses import dataclass, field as dc_field

from .errors import AuditError
from .oxml import W, localname

# 视图内事件项的可见性
BOTH = "both"
ACCEPT_ONLY = "accept"
REJECT_ONLY = "reject"


@dataclass
class Item:
    """段落内的一个可见性事件：文本 / 域控制符 / 指令 / 书签边界。"""

    kind: str  # text | fld_begin | fld_sep | fld_end | instr | bm_start | bm_end
    text: str = ""
    name: str = ""
    bm_id: str = ""
    vis: str = BOTH


@dataclass
class Para:
    items: list = dc_field(default_factory=list)
    ppr: object = None  # 幸存段落标记所属段落的 w:pPr 元素
    mark: str = ""  # "" | "ins" | "del"（段落标记修订状态）


@dataclass
class Field:
    instr: str
    result: str
    has_sep: bool


# ---------------------------------------------------------------- 段落提取


def _extract_run(r_el, vis: str, items: list) -> None:
    for child in r_el:
        if not isinstance(child.tag, str):
            continue
        tag = localname(child)
        if tag in ("t", "delText"):
            items.append(Item("text", text=child.text or "", vis=vis))
        elif tag == "tab":
            items.append(Item("text", text="\t", vis=vis))
        elif tag in ("br", "cr"):
            items.append(Item("text", text="\n", vis=vis))
        elif tag == "noBreakHyphen":
            items.append(Item("text", text="‑", vis=vis))
        elif tag == "instrText":
            items.append(Item("instr", text=child.text or "", vis=vis))
        elif tag == "fldChar":
            ftype = child.get(W + "fldCharType")
            kind = {
                "begin": "fld_begin",
                "separate": "fld_sep",
                "end": "fld_end",
            }.get(ftype)
            if kind:
                items.append(Item(kind, vis=vis))
        # rPr、drawing、pict、softHyphen 等与审查无关，忽略


def _extract_items(el, vis: str, items: list) -> None:
    for child in el:
        if not isinstance(child.tag, str):
            continue
        tag = localname(child)
        if tag == "pPr":
            continue
        if tag == "r":
            _extract_run(child, vis, items)
        elif tag == "ins":
            _extract_items(child, ACCEPT_ONLY, items)
        elif tag == "del":
            _extract_items(child, REJECT_ONLY, items)
        elif tag in ("hyperlink", "smartTag"):
            _extract_items(child, vis, items)
        elif tag == "sdt":
            content = child.find(W + "sdtContent")
            if content is not None:
                _extract_items(content, vis, items)
        elif tag == "fldSimple":
            # 简单域展开为 开始/指令/分隔/结果/结束 事件序列
            items.append(Item("fld_begin", vis=vis))
            items.append(Item("instr", text=child.get(W + "instr") or "", vis=vis))
            items.append(Item("fld_sep", vis=vis))
            _extract_items(child, vis, items)
            items.append(Item("fld_end", vis=vis))
        elif tag == "bookmarkStart":
            items.append(
                Item(
                    "bm_start",
                    name=child.get(W + "name") or "",
                    bm_id=child.get(W + "id") or "",
                    vis=vis,
                )
            )
        elif tag == "bookmarkEnd":
            items.append(Item("bm_end", bm_id=child.get(W + "id") or "", vis=vis))
        # proofErr、permStart、commentRangeStart 等忽略


def _extract_para(p_el) -> Para:
    ppr = p_el.find(W + "pPr")
    mark = ""
    if ppr is not None:
        rpr = ppr.find(W + "rPr")
        if rpr is not None:
            if rpr.find(W + "ins") is not None:
                mark = "ins"
            elif rpr.find(W + "del") is not None:
                mark = "del"
    items: list = []
    _extract_items(p_el, BOTH, items)
    return Para(items=items, ppr=ppr, mark=mark)


# ---------------------------------------------------------------- 容器收集


def collect_groups(container) -> list:
    """按文档顺序收集段落，同容器连续段落为一组（合并不跨组）。"""
    groups: list = []
    current: list = []

    def flush():
        nonlocal current
        if current:
            groups.append(current)
            current = []

    for child in container:
        if not isinstance(child.tag, str):
            continue
        tag = localname(child)
        if tag == "p":
            current.append(_extract_para(child))
        elif tag == "tbl":
            flush()
            for tr in child.findall(W + "tr"):
                for tc in tr.findall(W + "tc"):
                    groups.extend(collect_groups(tc))
        elif tag == "sdt":
            flush()
            content = child.find(W + "sdtContent")
            if content is not None:
                groups.extend(collect_groups(content))
        # sectPr、body 级 bookmarkStart 等忽略
    flush()
    return groups


# ---------------------------------------------------------------- 视图构建


def _absorbs_next(mark: str, view: str) -> bool:
    """该段落的段落标记在指定视图中是否被删除（从而与下一段合并）。"""
    return (view == "accept" and mark == "del") or (view == "reject" and mark == "ins")


def build_view(groups: list, view: str) -> list:
    """生成指定视图下可见段落的有序列表（含表格内段落）。"""
    out: list = []
    for group in groups:
        merged: list = []
        for para in group:
            items = [it for it in para.items if it.vis in (BOTH, view)]
            if merged and _absorbs_next(merged[-1].mark, view):
                # 段落标记被删除：与下一段合并，幸存标记属于下一段
                merged[-1].items = merged[-1].items + items
                merged[-1].ppr = para.ppr
                merged[-1].mark = para.mark
            else:
                merged.append(Para(items=list(items), ppr=para.ppr, mark=para.mark))
        out.extend(merged)
    return out


# ---------------------------------------------------------------- 域解析


def parse_fields(items: list) -> list:
    """解析单段落内的域结构，返回顶层域列表；结构损坏抛 AuditError。"""
    fields: list = []
    stack: list = []

    def broken(msg: str) -> AuditError:
        return AuditError("broken_field", f"域结构损坏：{msg}")

    for it in items:
        if it.kind == "fld_begin":
            stack.append({"instr": [], "result": [], "sep": False})
        elif it.kind == "instr":
            if not stack:
                raise broken("instrText 出现在域之外")
            if stack[-1]["sep"]:
                raise broken("instrText 出现在分隔符之后")
            stack[-1]["instr"].append(it.text)
        elif it.kind == "fld_sep":
            if not stack:
                raise broken("分隔符缺少对应的域开始符")
            stack[-1]["sep"] = True
        elif it.kind == "fld_end":
            if not stack:
                raise broken("结束符缺少对应的域开始符")
            top = stack.pop()
            fld = Field(
                instr="".join(top["instr"]),
                result="".join(top["result"]),
                has_sep=top["sep"],
            )
            if stack:
                # 嵌套域的显示结果计入外层域的结果文本
                if stack[-1]["sep"]:
                    stack[-1]["result"].append(fld.result)
            else:
                fields.append(fld)
        elif it.kind == "text":
            if stack and stack[-1]["sep"]:
                stack[-1]["result"].append(it.text)
    if stack:
        raise broken("域开始符缺少对应的结束符")
    return fields
