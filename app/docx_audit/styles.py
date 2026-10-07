"""WordprocessingML styles model: style inheritance and heading detection."""

from __future__ import annotations

from .namespaces import w

# Common built-in (English) and localized (Chinese / generated) heading
# style identifiers and display names.
_HEADING_STYLE_IDS = {
    "heading1": 1,
    "heading2": 2,
    "heading3": 3,
    "1": 1,
    "2": 2,
    "3": 3,  # built-in styleId commonly emitted by Word
}

_HEADING_NAME_LEVEL = {
    "heading 1": 1,
    "heading 2": 2,
    "heading 3": 3,
    "标题 1": 1,
    "标题 2": 2,
    "标题 3": 3,
    "标题1": 1,
    "标题2": 2,
    "标题3": 3,
}


class Styles:
    def __init__(self, root):
        self._styles: dict[str, dict] = {}
        if root is not None:
            self._parse(root)

    def _parse(self, root):
        for style in root.findall(w("style")):
            style_id = style.get(w("styleId"))
            if not style_id:
                continue
            based = style.find(w("basedOn"))
            name_el = style.find(w("name"))
            ppr = style.find(w("pPr"))
            self._styles[style_id] = {
                "type": style.get(w("type"), "paragraph"),
                "name": name_el.get(w("val"), "") if name_el is not None else "",
                "based_on": based.get(w("val")) if based is not None else None,
                "pPr": ppr,
            }

    def chain(self, style_id, seen=None):
        """Return the inheritance chain (style ids, most-derived first)."""
        if style_id is None or style_id not in self._styles:
            return []
        seen = seen if seen is not None else set()
        if style_id in seen:
            return []
        seen.add(style_id)
        style = self._styles[style_id]
        chain = [style_id]
        parent = style["based_on"]
        if parent:
            chain.extend(self.chain(parent, seen))
        return chain

    def heading_level(self, style_id):
        """Return 1/2/3 when the style (or an inherited one) is such a heading."""
        for sid in self.chain(style_id):
            style = self._styles.get(sid)
            if style is None:
                continue
            if sid.lower() in _HEADING_STYLE_IDS:
                level = _HEADING_STYLE_IDS[sid.lower()]
                if level <= 3:
                    return level
            name = style["name"].strip().lower()
            if name in _HEADING_NAME_LEVEL:
                return _HEADING_NAME_LEVEL[name]
        return None

    def numbering(self, style_id):
        """Effective ``(numId, ilvl)`` declared on the style chain, or None."""
        for sid in self.chain(style_id):
            style = self._styles.get(sid)
            if style is None or style["pPr"] is None:
                continue
            num_pr = style["pPr"].find(w("numPr"))
            if num_pr is None:
                continue
            num_id_el = num_pr.find(w("numId"))
            ilvl_el = num_pr.find(w("ilvl"))
            num_id = num_id_el.get(w("val")) if num_id_el is not None else None
            if num_id is None:
                continue
            return int(num_id), int(ilvl_el.get(w("val"), "0")) if ilvl_el is not None else 0
        return None

    def outline_level(self, style_id):
        """Effective ``w:outlineLvl`` (0-based) declared on the style chain."""
        for sid in self.chain(style_id):
            style = self._styles.get(sid)
            if style is None or style["pPr"] is None:
                continue
            outline = style["pPr"].find(w("outlineLvl"))
            if outline is not None:
                try:
                    return int(outline.get(w("val"), "0"))
                except ValueError:
                    return None
        return None
