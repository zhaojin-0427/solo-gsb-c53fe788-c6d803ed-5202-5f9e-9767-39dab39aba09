"""styles.xml 解析：样式继承链、Heading 1~3 识别、样式级编号。"""

from __future__ import annotations

import re

from .oxml import W, wval

_HEADING_NAME_RE = re.compile(r"^\s*heading\s*([1-9])\s*$", re.IGNORECASE)
_HEADING_ID_RE = re.compile(r"^Heading([1-9])$")


class Styles:
    def __init__(self, root):
        # styleId -> {name, based_on, has_numpr, num_id, ilvl}
        self._styles: dict = {}
        if root is None:
            return
        for st in root.findall(W + "style"):
            sid = st.get(W + "styleId")
            if not sid:
                continue
            name = wval(st.find(W + "name")) or ""
            based_on = wval(st.find(W + "basedOn"))
            has_numpr = False
            num_id = None
            ilvl = None
            ppr = st.find(W + "pPr")
            if ppr is not None:
                numpr = ppr.find(W + "numPr")
                if numpr is not None:
                    has_numpr = True
                    num_id = wval(numpr.find(W + "numId"))
                    ilvl_text = wval(numpr.find(W + "ilvl"))
                    ilvl = int(ilvl_text) if ilvl_text is not None else None
            self._styles[sid] = {
                "name": name,
                "based_on": based_on,
                "has_numpr": has_numpr,
                "num_id": num_id,
                "ilvl": ilvl,
            }

    def chain(self, sid):
        """从 sid 出发沿 basedOn 向上，逐级产出 styleId（防循环）。"""
        seen = set()
        cur = sid
        while cur and cur in self._styles and cur not in seen:
            seen.add(cur)
            yield cur
            cur = self._styles[cur]["based_on"]

    def heading_level(self, sid):
        """识别 Heading 1~3（含基于它们的自定义样式），返回级别或 None。"""
        for cur in self.chain(sid):
            info = self._styles[cur]
            m = _HEADING_NAME_RE.match(info["name"]) or _HEADING_ID_RE.match(cur)
            if m:
                return int(m.group(1))
        return None

    def style_numpr(self, sid):
        """沿继承链解析样式级编号。

        返回 (num_id, ilvl)；返回 "disabled" 表示链上显式取消编号
        （numId 缺失或为 0）；返回 None 表示链上没有编号信息。
        """
        for cur in self.chain(sid):
            info = self._styles[cur]
            if info["has_numpr"]:
                if not info["num_id"] or info["num_id"] == "0":
                    return "disabled"
                return (info["num_id"], info["ilvl"] if info["ilvl"] is not None else 0)
        return None
