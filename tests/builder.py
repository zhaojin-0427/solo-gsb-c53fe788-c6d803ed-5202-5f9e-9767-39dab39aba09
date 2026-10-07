"""Minimal OOXML DOCX builder used by tests and the sample generator."""

from __future__ import annotations

import io
import zipfile

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def we(tag):
    return f"{{{W}}}{tag}"


def re(tag):
    return f"{{{R}}}{tag}"


def el(tag, attrs=None, text=None):
    node = etree.Element(we(tag))
    for key, value in (attrs or {}).items():
        node.set(we(key), str(value))
    if text is not None:
        node.text = text
    return node


class DocxBuilder:
    def __init__(self):
        self.body = etree.Element(we("body"))
        self._bookmark_seq = 0
        # numbering definitions: num_id -> dict(abstract_id, levels, overrides)
        self._nums = {}
        self._styles = {}
        self._next_num_id = 1
        self._next_abstract_id = 0

    # ------------------------------------------------------------ styles
    def add_style(self, style_id, name=None, based_on=None, outline=None,
                  num_id=None, ilvl=0):
        self._styles[style_id] = {
            "name": name or style_id,
            "based_on": based_on,
            "outline": outline,
            "num_id": num_id,
            "ilvl": ilvl,
        }

    def add_default_heading_styles(self):
        for i in (1, 2, 3):
            self.add_style(f"Heading{i}", name=f"Heading {i}", outline=i - 1)

    # ---------------------------------------------------------- numbering
    def add_multilevel_numbering(
        self,
        num_id=1,
        texts=None,
        fmts=None,
        starts=None,
        restarts=None,
        is_lgl=False,
        style_levels=None,
        start_overrides=None,
    ):
        texts = texts or ["%1", "%1.%2", "%1.%2.%3"]
        fmts = fmts or ["decimal", "decimal", "decimal"]
        starts = starts or [1, 1, 1]
        restarts = restarts if restarts is not None else [0, 0, 0]
        abstract_id = self._next_abstract_id
        self._next_abstract_id += 1
        self._nums[num_id] = {
            "abstract": abstract_id,
            "texts": texts,
            "fmts": fmts,
            "starts": starts,
            "restarts": restarts,
            "is_lgl": is_lgl,
            "style_levels": style_levels or {},
            "start_overrides": start_overrides or {},
        }
        return num_id

    # ------------------------------------------------------------ blocks
    def paragraph(self, *runs, style=None, num_id=None, ilvl=0,
                  pilcrow=None, bookmark_starts=None, bookmark_ends=None,
                  append_to=None, ppr_num_id=..., ppr_ilvl=None):
        p = el("p")
        ppr = el("pPr")
        if style:
            ppr.append(el("pStyle", {"val": style}))
        if num_id is not None:
            numpr = el("numPr")
            if ppr_ilvl is not None:
                numpr.append(el("ilvl", {"val": ppr_ilvl}))
            elif ilvl:
                numpr.append(el("ilvl", {"val": ilvl}))
            numpr.append(el("numId", {"val": num_id}))
            ppr.append(numpr)
        if pilcrow is not None:
            kind, author = pilcrow
            rpr = el("rPr")
            rpr.append(el(kind, {"id": 900, "author": author, "date": "2026-01-01T00:00:00Z"}))
            ppr.append(rpr)
        p.append(ppr)

        for name in bookmark_starts or []:
            self._bookmark_seq += 1
            p.append(el("bookmarkStart", {
                "id": self._bookmark_seq, "name": name,
            }))
        for run in runs:
            self._append_run(p, run)
        for bid in bookmark_ends or []:
            self._bookmark_seq += 1
            p.append(el("bookmarkEnd", {"id": self._bookmark_seq}))

        target = append_to if append_to is not None else self.body
        target.append(p)
        return p

    def heading(self, text, level=1, num_id=1, ilvl=None, style=None,
                bookmark=None, pilcrow=None, runs=None, append_to=None):
        ilvl = level - 1 if ilvl is None else ilvl
        style = style or f"Heading{level}"
        use_runs = runs if runs is not None else [self.text_run(text)]
        return self.paragraph(
            *use_runs,
            style=style,
            num_id=num_id,
            ilvl=ilvl,
            pilcrow=pilcrow,
            bookmark_starts=[bookmark] if bookmark else None,
            append_to=append_to,
        )

    def _append_run(self, target, run):
        """Append a run tuple produced by the run factories below."""
        kind = run[0]
        if kind == "wrap":
            _, change, author, children = run
            wrapper = el(change, {"id": abs(hash(text_id(author))) % 100000,
                                  "author": author,
                                  "date": "2026-01-01T00:00:00Z"})
            for child in children:
                self._append_run(wrapper, child)
            target.append(wrapper)
            return
        r = el("r")
        if kind == "text":
            _, value = run
            t = el("t", {"space": "preserve"}, text=value)
            r.append(t)
        elif kind == "deltext":
            _, value = run
            r.append(el("delText", {"space": "preserve"}, text=value))
        elif kind == "fld":
            _, fldtype = run
            r.append(el("fldChar", {"fldCharType": fldtype}))
        elif kind == "instr":
            _, value = run
            r.append(el("instrText", {"space": "preserve"}, text=value))
        elif kind == "bookmarkStart":
            self._bookmark_seq += 1
            target.append(el("bookmarkStart", {"id": self._bookmark_seq, "name": run[1]}))
            return
        elif kind == "bookmarkEnd":
            self._bookmark_seq += 1
            target.append(el("bookmarkEnd", {"id": self._bookmark_seq}))
            return
        else:
            raise ValueError(kind)
        target.append(r)

    # ------------------------------------------------------- run factories
    @staticmethod
    def text_run(value):
        return ("text", value)

    @staticmethod
    def deleted_text(value):
        return ("wrap", "del", "tester", [("deltext", value)])

    @staticmethod
    def inserted_text(value):
        return ("wrap", "ins", "tester", [("text", value)])

    @staticmethod
    def wrap(change, author, *children):
        return ("wrap", change, author, list(children))

    @staticmethod
    def fld(fldtype):
        return ("fld", fldtype)

    @staticmethod
    def instr(value):
        return ("instr", value)

    @staticmethod
    def bookmark_run(name):
        return ("bookmarkStart", name)

    def ref_field_runs(self, name, switch, cached, hyperlink=False, split=False):
        instruction = f' REF {name} \\{switch}' + (" \\h" if hyperlink else "") + " "
        if not split:
            return [
                self.fld("begin"),
                self.instr(instruction),
                self.fld("separate"),
                self.text_run(cached),
                self.fld("end"),
            ]
        # Split every dynamic part into its own run.
        return [
            self.fld("begin"),
            self.instr(" REF "),
            self.instr(name),
            self.instr(f" \\{switch} "),
            self.fld("separate"),
            self.text_run(cached[: len(cached) // 2]),
            self.text_run(cached[len(cached) // 2 :]),
            self.fld("end"),
        ]

    def ref_field_wrapped(self, change, name, switch, cached):
        return self.wrap(
            change,
            "tester",
            self.fld("begin"),
            self.instr(f" REF {name} \\{switch} "),
            self.fld("separate"),
            ("text" if change == "ins" else "deltext", cached),
            self.fld("end"),
        )

    # -------------------------------------------------------------- tables
    def table(self, rows, style=None):
        """rows: list[list[list[paragraph-runs/kwargs dicts]]]."""
        tbl = el("tbl")
        tblpr = el("tblPr")
        tblpr.append(el("tblStyle", {"val": style or "TableGrid"}))
        tbl.append(tblpr)
        for row in rows:
            tr = el("tr")
            for cell_paragraphs in row:
                tc = el("tc")
                for spec in cell_paragraphs:
                    kwargs = dict(spec)
                    runs = kwargs.pop("runs", [])
                    self.paragraph(*runs, append_to=tc, **kwargs)
                tr.append(tc)
            tbl.append(tr)
        self.body.append(tbl)
        return tbl

    # ------------------------------------------------------------- export
    def build(self) -> bytes:
        document = self._document_xml()
        styles_xml = self._styles_xml()
        numbering_xml = self._numbering_xml()

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("[Content_Types].xml", self._content_types())
            zf.writestr("_rels/.rels", self._root_rels())
            zf.writestr("word/_rels/document.xml.rels", self._doc_rels())
            zf.writestr("word/document.xml", document)
            zf.writestr("word/styles.xml", styles_xml)
            zf.writestr("word/numbering.xml", numbering_xml)
        return buf.getvalue()

    def _document_xml(self):
        root = el("document")
        self.body.append(el("sectPr"))
        root.append(self.body)
        return etree.tostring(
            root, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    def _styles_xml(self):
        root = el("styles")
        # Default Normal.
        normal = el("style", {"type": "paragraph", "styleId": "Normal",
                              "default": "1"})
        normal.append(el("name", {"val": "Normal"}))
        root.append(normal)
        for sid, info in self._styles.items():
            style = el("style", {"type": "paragraph", "styleId": sid})
            style.append(el("name", {"val": info["name"]}))
            if info["based_on"]:
                style.append(el("basedOn", {"val": info["based_on"]}))
            ppr = el("pPr")
            if info["outline"] is not None:
                ppr.append(el("outlineLvl", {"val": info["outline"]}))
            if info["num_id"] is not None:
                numpr = el("numPr")
                numpr.append(el("ilvl", {"val": info["ilvl"]}))
                numpr.append(el("numId", {"val": info["num_id"]}))
                ppr.append(numpr)
            if len(ppr):
                style.append(ppr)
            root.append(style)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                              standalone=True)

    def _numbering_xml(self):
        root = el("numbering")
        for num_id, info in self._nums.items():
            abstract = el("abstractNum", {"abstractNumId": info["abstract"]})
            abstract.append(el("multiLevelType", {"val": "multilevel"}))
            for lvl in range(9):
                attrs = {"ilvl": lvl}
                level_el = el("lvl", attrs)
                if lvl < len(info["texts"]):
                    level_el.append(el("start", {"val": info["starts"][lvl]}))
                    level_el.append(el("numFmt", {"val": info["fmts"][lvl]}))
                    level_el.append(el("lvlText", {"val": info["texts"][lvl]}))
                    level_el.append(el("suff", {"val": "space"}))
                    if lvl < len(info["restarts"]):
                        level_el.append(el("lvlRestart", {"val": info["restarts"][lvl]}))
                    if info["is_lgl"]:
                        level_el.append(el("isLgl"))
                else:
                    level_el.append(el("start", {"val": 1}))
                    level_el.append(el("numFmt", {"val": "decimal"}))
                    level_el.append(el("lvlText", {"val": f"%{lvl + 1}"}))
                style = next((s for s, il in info["style_levels"].items() if il == lvl), None)
                if style:
                    level_el.append(el("pStyle", {"val": style}))
                abstract.append(level_el)
            root.append(abstract)

            num = el("num", {"numId": num_id})
            num.append(el("abstractNumId", {"val": info["abstract"]}))
            for ilvl, value in info["start_overrides"].items():
                override = el("lvlOverride", {"ilvl": ilvl})
                override.append(el("startOverride", {"val": value}))
                num.append(override)
            root.append(num)
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8",
                              standalone=True)

    @staticmethod
    def _content_types():
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType='
            '"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '<Override PartName="/word/styles.xml" ContentType='
            '"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
            '<Override PartName="/word/numbering.xml" ContentType='
            '"application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
            "</Types>"
        )

    @staticmethod
    def _root_rels():
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type='
            '"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
            ' Target="word/document.xml"/></Relationships>'
        )

    @staticmethod
    def _doc_rels():
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type='
            '"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"'
            ' Target="styles.xml"/>'
            '<Relationship Id="rId2" Type='
            '"http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"'
            ' Target="numbering.xml"/></Relationships>'
        )


_counter = [0]


def text_id(seed):
    _counter[0] += 1
    return seed + str(_counter[0])
