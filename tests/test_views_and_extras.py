"""Tests for pilcrow revisions, tables, styles and fldSimple."""
import pytest

from builder import DocxBuilder
from docx_audit.auditor import audit_document, ACCEPT, REJECT
from docx_audit.errors import AuditError


def fields_for(report, view):
    return report[view]["fields"]


# ============================================================ pilcrow
def test_deleted_pilcrow_merges_into_following_paragraph():
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1, start_overrides={0: 4})
    b.heading("第四章", level=1)
    # Heading with a deleted pilcrow: accept -> merged into following body
    # paragraph (numbering of that body paragraph comes from the heading);
    # reject -> heading stays its own numbered paragraph.
    b.heading(
        "被合并标题",
        level=2,
        bookmark="h",
        pilcrow=("del", "tester"),
    )
    b.paragraph(b.text_run("尾随正文，引用见 "), *b.ref_field_runs("h", "w", "4.1"))
    report = audit_document(b.build())

    accept = fields_for(report, ACCEPT)[0]
    reject = fields_for(report, REJECT)[0]
    assert accept["expected_value"] == "4.1"
    assert reject["expected_value"] == "4.1"
    # In accept view the field's logical paragraph is the merged one.
    assert accept["position"]["logical_paragraph_index"] == 1
    assert reject["position"]["logical_paragraph_index"] == 2


def test_inserted_pilcrow_reject_merges_into_previous():
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1)
    b.paragraph(b.text_run("前段正文"))
    # Inserted paragraph mark plus inserted content (as Word records a
    # tracked Enter): reject -> the new paragraph is gone entirely;
    # accept -> standalone numbered heading.
    b.heading(
        "新标题",
        level=1,
        pilcrow=("ins", "tester"),
        runs=[b.wrap("ins", "tester", b.bookmark_run("h"), ("text", "新标题"))],
    )
    b.paragraph(*b.ref_field_runs("h", "n", "1"))
    report = audit_document(b.build())
    accept = fields_for(report, ACCEPT)[0]
    reject = fields_for(report, REJECT)[0]
    assert accept["expected_value"] == "1"
    # In reject view the heading disappears (rejected insertion) -> missing.
    assert reject["status"] == "missing_target"
    assert accept["status"] == "correct"


def test_conflicting_pilcrow_revisions_422():
    from lxml import etree
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1)
    # heading() with pilcrow=ins creates w:rPr/w:ins; append w:del as well.
    p = b.heading("x", level=1, pilcrow=("ins", "tester"))
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    rpr = p.find(W + "pPr").find(W + "rPr")
    rpr.append(etree.SubElement(rpr, W + "del",
                                {W + "id": "2", W + "author": "a"}))
    with pytest.raises(AuditError) as exc:
        audit_document(b.build())
    assert exc.value.code == "conflicting_pilcrow_revision"


# =============================================================== tables
def test_headings_and_fields_inside_table_are_audited():
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1, start_overrides={0: 3})
    b.heading("第三章", level=1)
    b.table([
        [
            [
                {"runs": [b.bookmark_run("t")], "style": "Heading2",
                 "num_id": 1, "ilvl": 1},
            ],
            [
                {"runs": [*b.ref_field_runs("t", "w", "3.1")]},
            ],
        ],
    ])
    report = audit_document(b.build())
    f = fields_for(report, ACCEPT)[0]
    assert f["expected_value"] == "3.1"
    assert f["status"] == "correct"
    assert f["position"]["table_path"] is not None
    assert "tbl" in f["position"]["table_path"]
    assert "tc" in f["position"]["table_path"]


def test_numbering_continues_across_table():
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1)
    b.heading("第一章", level=1)
    b.table([[[{"runs": [b.text_run("cell")]}]]])
    b.heading("第二章", level=1, bookmark="h")
    b.paragraph(*b.ref_field_runs("h", "n", "2"))
    report = audit_document(b.build())
    assert fields_for(report, ACCEPT)[0]["expected_value"] == "2"


# =============================================================== styles
def test_style_inheritance_for_heading_level():
    b = DocxBuilder()
    # Custom style based on Heading1.
    b.add_style("MyH1", name="My Heading 1", based_on="Heading1", outline=0)
    b.add_style("Heading1", name="heading 1", outline=0)
    b.add_multilevel_numbering(num_id=1)
    b.heading("自定义标题", level=1, style="MyH1", bookmark="h")
    b.paragraph(*b.ref_field_runs("h", "w", "1"))
    report = audit_document(b.build())
    assert fields_for(report, ACCEPT)[0]["status"] == "correct"


def test_numbering_declared_on_style():
    b = DocxBuilder()
    b.add_style("Heading1", name="heading 1", outline=0, num_id=2, ilvl=0)
    b.add_multilevel_numbering(num_id=2)
    b.paragraph(b.text_run("x"), style="Heading1")
    b.paragraph(b.text_run("y"), style="Heading1", bookmark_starts=["h"])
    b.paragraph(*b.ref_field_runs("h", "n", "2"))
    report = audit_document(b.build())
    f = fields_for(report, ACCEPT)[0]
    assert f["expected_value"] == "2"
    assert f["status"] == "correct"


def test_direct_numid_zero_cancels_style_numbering():
    b = DocxBuilder()
    b.add_style("Heading1", name="heading 1", outline=0, num_id=2, ilvl=0)
    b.add_multilevel_numbering(num_id=2)
    b.paragraph(b.text_run("x"), style="Heading1")
    # numId=0 cancellation
    p = b.paragraph(b.text_run("y"), style="Heading1", bookmark_starts=["h"],
                    num_id=0)
    b.paragraph(*b.ref_field_runs("h", "w", ""))
    report = audit_document(b.build())
    assert fields_for(report, ACCEPT)[0]["status"] == "missing_target"


# ============================================================ fldSimple
def test_fldsimple_ref_is_audited():
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1)
    b.heading("目标", level=1, bookmark="h")
    p = b.paragraph()
    # Append fldSimple manually.
    from lxml import etree
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    simple = etree.SubElement(p, W + "fldSimple", {W + "instr": " REF h \\w "})
    r = etree.SubElement(simple, W + "r")
    etree.SubElement(r, W + "t", {W + "space": "preserve"}).text = "1"
    report = audit_document(b.build())
    f = fields_for(report, ACCEPT)[0]
    assert f["expected_value"] == "1"
    assert f["status"] == "correct"


# ============================================================ revision 422
def test_pprchange_is_tolerated():
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1)
    b.heading("目标", level=1, bookmark="h")
    b.paragraph(*b.ref_field_runs("h", "w", "1"))
    data = _inject(b, "pPrChange")
    report = audit_document(data)
    assert fields_for(report, ACCEPT)[0]["status"] == "correct"


def test_comment_revision_returns_422():
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1)
    b.heading("目标", level=1, bookmark="h")
    data = _inject(b, "commentRangeStart", attrs_needed=True)
    with pytest.raises(AuditError) as exc:
        audit_document(data)
    assert exc.value.code == "unsupported_revision"


def test_numbering_change_returns_422():
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1)
    b.heading("目标", level=1, bookmark="h")
    b.paragraph(*b.ref_field_runs("h", "w", "1"))
    data = _inject_numbering_change(b)
    with pytest.raises(AuditError) as exc:
        audit_document(data)
    assert exc.value.code == "unsupported_revision"


# ---------------------------------------------------------------------------
# style-linked numbering / text boxes / block-level revisions
# ---------------------------------------------------------------------------


def test_style_linked_numbering_via_abstract_pstyle():
    # abstractNum/lvl/pStyle links a custom style to a numbering level.
    b = DocxBuilder()
    b.add_style("Heading1", name="heading 1", outline=0)
    b.add_style("MyNumberedHeading", name="My Numbered Heading",
                based_on="Heading1", outline=0)
    b.add_multilevel_numbering(
        num_id=5,
        style_levels={"MyNumberedHeading": 0},
    )
    b.paragraph(b.text_run("一"), style="MyNumberedHeading")
    b.paragraph(b.text_run("二"), style="MyNumberedHeading",
                bookmark_starts=["h"])
    b.paragraph(*b.ref_field_runs("h", "n", "2"))
    report = audit_document(b.build())
    f = fields_for(report, ACCEPT)[0]
    assert f["expected_value"] == "2"
    assert f["status"] == "correct"


def test_paragraphs_inside_textbox_are_ignored():
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1)
    # A "numbered heading" buried inside a text box must not consume a
    # counter; the real body heading must number as 1.
    from lxml import etree
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    txbx = etree.Element(W + "txbxContent")
    fake = etree.SubElement(txbx, W + "p")
    ppr = etree.SubElement(fake, W + "pPr")
    etree.SubElement(ppr, W + "pStyle", {W + "val": "Heading1"})
    numpr = etree.SubElement(ppr, W + "numPr")
    etree.SubElement(numpr, W + "ilvl", {W + "val": "0"})
    etree.SubElement(numpr, W + "numId", {W + "val": "1"})
    container = etree.Element(W + "pict")
    container.append(txbx)
    run_holder = b.paragraph()
    r = etree.SubElement(run_holder, W + "r")
    r.append(container)

    b.heading("真正标题", level=1, bookmark="h")
    b.paragraph(*b.ref_field_runs("h", "w", "1"))
    report = audit_document(b.build())
    assert fields_for(report, ACCEPT)[0]["expected_value"] == "1"


def test_block_level_inserted_paragraph():
    # w:ins directly under body wrapping runs of the whole paragraph.
    import zipfile, io
    from lxml import etree
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1)
    b.heading("保留标题", level=1, bookmark="h")
    data = b.build()
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    zf = zipfile.ZipFile(io.BytesIO(data))
    root = etree.fromstring(zf.read("word/document.xml"))
    body = root.find(W + "body")
    ins = etree.Element(W + "ins", {W + "id": "7", W + "author": "q"})
    r = etree.SubElement(ins, W + "r")
    etree.SubElement(r, W + "t").text = "整段插入正文"
    body.insert(list(body).index(body.find(W + "sectPr")), ins)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in zf.namelist():
            content = zf.read(item)
            if item == "word/document.xml":
                content = etree.tostring(root)
            z.writestr(item, content)
    report = audit_document(out.getvalue())
    # No numbered fields in the inserted run, but the document must audit
    # without raising an "unsupported revision" error.
    assert fields_for(report, ACCEPT) == []
    assert fields_for(report, REJECT) == []


def _inject(b, tag_name, attrs_needed=False):
    import zipfile
    import io
    from lxml import etree
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    zf = zipfile.ZipFile(io.BytesIO(b.build()))
    root = etree.fromstring(zf.read("word/document.xml"))
    body = root.find(W + "body")
    if tag_name == "pPrChange":
        ppr = body.find(W + "p").find(W + "pPr")
        node = etree.SubElement(ppr, W + tag_name, {W + "id": "5", W + "author": "q"})
        etree.SubElement(node, W + "pStyle", {W + "val": "Heading1"})
    else:
        etree.SubElement(body, W + tag_name)
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in zf.namelist():
            content = zf.read(item)
            if item == "word/document.xml":
                content = etree.tostring(root)
            z.writestr(item, content)
    return out.getvalue()


def _inject_numbering_change(b):
    import zipfile
    import io
    from lxml import etree
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    zf = zipfile.ZipFile(io.BytesIO(b.build()))
    root = etree.fromstring(zf.read("word/document.xml"))
    ppr = root.find(W + "body").find(W + "p").find(W + "pPr")
    numpr = ppr.find(W + "numPr")
    etree.SubElement(numpr, W + "numberingChange",
                     {W + "id": "9", W + "author": "q", W + "old": "1", W + "new": "2"})
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in zf.namelist():
            content = zf.read(item)
            if item == "word/document.xml":
                content = etree.tostring(root)
            z.writestr(item, content)
    return out.getvalue()
