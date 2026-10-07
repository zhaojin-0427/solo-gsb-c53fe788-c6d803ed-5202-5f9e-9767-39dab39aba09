"""Integration tests for the audit pipeline and its accept/reject views."""
import pytest

from builder import DocxBuilder, we
from docx_audit.auditor import audit_document, ACCEPT, REJECT
from docx_audit.errors import AuditError
from lxml import etree


def make_doc():
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1, start_overrides={0: 4})
    return b


def fields_for(report, view):
    return report[view]["fields"]


def by_bookmark(report, view):
    return {f["bookmark"] + ":" + f["switch"]: f for f in report[view]["fields"]}


def test_three_switches_basic():
    b = make_doc()
    b.heading("第四章", level=1)                       # 4
    b.heading("a", level=2)                            # 4.1
    b.heading("b", level=2)                            # 4.2
    b.heading("c", level=2)                            # 4.3
    b.heading("源标题", level=3, bookmark="src")       # 4.3.1
    b.paragraph(
        b.text_run("前向引用见 "),
        *b.ref_field_runs("tgt", "r", "5.2"),
        b.text_run("；n="),
        *b.ref_field_runs("tgt", "n", "2"),
        b.text_run("；w="),
        *b.ref_field_runs("tgt", "w", "4.5.2"),
    )
    b.heading("d", level=2)                            # 4.4
    b.heading("e", level=2)                            # 4.5
    b.heading("e1", level=3)                           # 4.5.1
    b.heading("目标标题", level=3, bookmark="tgt")     # 4.5.2

    report = audit_document(b.build())
    for view in (ACCEPT, REJECT):
        entries = by_bookmark(report, view)
        assert entries["tgt:r"]["expected_value"] == "5.2"
        assert entries["tgt:r"]["status"] == "correct"
        assert entries["tgt:n"]["expected_value"] == "2"
        assert entries["tgt:n"]["status"] == "correct"
        assert entries["tgt:w"]["expected_value"] == "4.5.2"
        assert entries["tgt:w"]["status"] == "correct"


def test_stale_field_detected():
    b = make_doc()
    b.heading("第四章", level=1)
    b.heading("a", level=2)
    b.paragraph(*b.ref_field_runs("h", "w", "1.99"))
    b.heading("目标", level=2, bookmark="h")           # 4.2
    report = audit_document(b.build())
    f = fields_for(report, ACCEPT)[0]
    assert f["status"] == "stale"
    assert f["current_value"] == "1.99"
    assert f["expected_value"] == "4.2"


def test_missing_target_bookmark_deleted_in_view():
    b = make_doc()
    b.heading("第四章", level=1)
    # Entire heading paragraph text deleted; bookmark sits inside deleted run.
    b.heading(
        "删除目标",
        level=2,
        runs=[b.wrap("del", "tester", b.bookmark_run("gone"), ("deltext", "删除目标"))],
    )
    b.paragraph(*b.ref_field_runs("gone", "w", "4.1"))

    report = audit_document(b.build())
    accept = fields_for(report, ACCEPT)[0]
    reject = fields_for(report, REJECT)[0]
    assert accept["status"] == "missing_target"
    assert accept["expected_value"] is None
    assert reject["status"] == "correct"
    assert reject["expected_value"] == "4.1"
    # Views never borrow targets from one another.
    assert accept["expected_value"] != reject["expected_value"]


def test_inserted_target_visible_only_in_accept_view():
    b = make_doc()
    b.heading("第四章", level=1)
    b.paragraph(*b.ref_field_runs("new", "n", "1"))
    b.heading(
        "新增目标",
        level=2,
        runs=[b.wrap("ins", "tester", b.bookmark_run("new"), ("text", "新增目标"))],
    )
    report = audit_document(b.build())
    accept = fields_for(report, ACCEPT)[0]
    reject = fields_for(report, REJECT)[0]
    assert accept["status"] == "correct"
    assert accept["expected_value"] == "1"  # first level2 under chapter 4
    assert reject["status"] == "missing_target"


def test_inserted_field_audited_only_in_accept():
    b = make_doc()
    b.heading("第四章", level=1)
    b.heading("目标", level=2, bookmark="h")          # 4.1
    b.paragraph(b.ref_field_wrapped("ins", "h", "w", "4.1"))
    report = audit_document(b.build())
    assert len(fields_for(report, ACCEPT)) == 1
    assert fields_for(report, ACCEPT)[0]["status"] == "correct"
    assert fields_for(report, REJECT) == []


def test_deleted_field_audited_only_in_reject():
    b = make_doc()
    b.heading("第四章", level=1)
    b.heading("目标", level=2, bookmark="h")
    b.paragraph(b.ref_field_wrapped("del", "h", "w", "4.1"))
    report = audit_document(b.build())
    assert fields_for(report, ACCEPT) == []
    assert len(fields_for(report, REJECT)) == 1
    assert fields_for(report, REJECT)[0]["status"] == "correct"


def test_cross_run_field_matches_bookmark():
    b = make_doc()
    b.heading("第四章", level=1)
    b.heading("目标", level=2, bookmark="h")
    b.paragraph(
        *[
            b.fld("begin"),
            b.instr(" RE"),
            b.instr("F h \\"),
            b.instr("w "),
            b.fld("separate"),
            b.text_run("4."),
            b.text_run("1"),
            b.fld("end"),
        ]
    )
    report = audit_document(b.build())
    f = fields_for(report, ACCEPT)[0]
    assert f["current_value"] == "4.1"
    assert f["expected_value"] == "4.1"
    assert f["status"] == "correct"


def test_bookmark_span_across_runs():
    b = make_doc()
    b.heading("第四章", level=1)
    # bookmarkStart in one run-group, heading text split across runs
    p = b.heading(
        "目标标题",
        level=2,
        bookmark=None,
        runs=[
            b.bookmark_run("h"),
            b.text_run("目标"),
            b.text_run("标题"),
        ],
    )
    b.paragraph(*b.ref_field_runs("h", "w", "4.1"))
    report = audit_document(b.build())
    assert fields_for(report, ACCEPT)[0]["status"] == "correct"


def test_duplicate_bookmark_returns_422():
    b = make_doc()
    b.heading("第四章", level=1)
    b.heading("目标一", level=2, bookmark="dup")
    b.heading("目标二", level=2, bookmark="dup")
    with pytest.raises(AuditError) as exc:
        audit_document(b.build())
    assert exc.value.code == "duplicate_bookmark"


def test_unknown_bookmark_is_missing_target_not_422():
    b = make_doc()
    b.heading("第四章", level=1)
    b.paragraph(*b.ref_field_runs("ghost", "w", "9.9"))
    report = audit_document(b.build())
    f = fields_for(report, ACCEPT)[0]
    assert f["status"] == "missing_target"
    assert fields_for(report, REJECT)[0]["status"] == "missing_target"


def test_field_missing_end_returns_422():
    b = make_doc()
    b.heading("第四章", level=1)
    b.paragraph(
        b.fld("begin"), b.instr(" REF h \\w "), b.fld("separate"),
        b.text_run("4.1"),
    )
    with pytest.raises(AuditError) as exc:
        audit_document(b.build())
    assert exc.value.code == "malformed_field"


def test_field_crossing_revision_wrapper_422():
    b = make_doc()
    b.heading("第四章", level=1)
    b.heading("目标", level=2, bookmark="h")
    b.paragraph(
        b.wrap("ins", "tester", b.fld("begin"), b.instr(" REF h \\w ")),
        b.fld("separate"), b.text_run("4.1"), b.fld("end"),
    )
    with pytest.raises(AuditError) as exc:
        audit_document(b.build())
    assert exc.value.code == "malformed_field"


def test_move_revision_returns_422():
    b = make_doc()
    b.heading("第四章", level=1)
    b.heading("目标", level=2, bookmark="h")
    b.paragraph(*b.ref_field_runs("h", "w", "4.1"))
    with pytest.raises(AuditError) as exc:
        audit_document(_rebuild_with_move(b))
    assert exc.value.code == "unsupported_revision"


def _rebuild_with_move(b):
    data = b.build()
    import zipfile
    import io
    from docx_audit.namespaces import w as W
    zf = zipfile.ZipFile(io.BytesIO(data))
    doc_root = etree.fromstring(zf.read("word/document.xml"))
    move = etree.SubElement(doc_root.find(W("body")), W("moveTo"))
    move.set(W("author"), "x")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in zf.namelist():
            content = zf.read(item)
            if item == "word/document.xml":
                content = etree.tostring(doc_root)
            z.writestr(item, content)
    return out.getvalue()


def test_no_r_context_returns_422():
    b = make_doc()
    # Field appears before any numbered heading at all.
    b.paragraph(*b.ref_field_runs("h", "r", "1"))
    b.heading("第四章", level=1)
    b.heading("目标", level=2, bookmark="h")
    with pytest.raises(AuditError) as exc:
        audit_document(b.build())
    assert exc.value.code == "no_r_context"


def test_r_context_deleted_in_view_422_for_that_view():
    b = make_doc()
    b.heading(
        "被删上下文章",
        level=1,
        runs=[b.wrap("del", "tester", b.bookmark_run("x"), ("deltext", "旧章"))],
    )
    # Field is inserted, so it exists only in accept view, where the context
    # chapter is gone -> 422 (no partial report).
    b.paragraph(b.ref_field_wrapped("ins", "h", "r", "1"))
    b.heading("目标", level=1, bookmark="h")
    with pytest.raises(AuditError) as exc:
        audit_document(b.build())
    assert exc.value.code == "no_r_context"


def test_r_relative_after_heading_insertion_changes_numbering():
    b = make_doc()
    b.heading("第四章", level=1)
    b.heading("a", level=2)
    b.heading(
        "新增小节",
        level=2,
        runs=[b.wrap("ins", "tester", b.bookmark_run("new"), ("text", "新增小节"))],
    )
    b.heading("c", level=2, bookmark="c")       # accept 4.3 / reject 4.2
    b.paragraph(*b.ref_field_runs("c", "n", "3"))
    report = audit_document(b.build())
    accept = fields_for(report, ACCEPT)[0]
    reject = fields_for(report, REJECT)[0]
    assert accept["expected_value"] == "3"
    assert reject["expected_value"] == "2"
    assert accept["status"] == "correct"
    assert reject["status"] == "stale"


def test_unnumbered_heading_target_is_missing():
    b = DocxBuilder()
    b.add_default_heading_styles()
    b.add_multilevel_numbering(num_id=1)
    b.heading("无编号标题", level=1, num_id=None, bookmark="h")
    b.paragraph(*b.ref_field_runs("h", "w", ""))
    report = audit_document(b.build())
    f = fields_for(report, ACCEPT)[0]
    assert f["status"] == "missing_target"


def test_non_heading_bookmark_target_is_missing():
    b = make_doc()
    b.heading("第四章", level=1)
    b.paragraph(b.bookmark_run("note"), b.text_run("普通正文"))
    b.paragraph(*b.ref_field_runs("note", "w", "x"))
    report = audit_document(b.build())
    assert fields_for(report, ACCEPT)[0]["status"] == "missing_target"


def test_plain_ref_without_number_switch_is_ignored():
    b = make_doc()
    b.heading("第四章", level=1)
    b.heading("目标", level=2, bookmark="h")
    b.paragraph(
        b.fld("begin"), b.instr(" REF h \\h "), b.fld("separate"),
        b.text_run("目标"), b.fld("end"),
    )
    report = audit_document(b.build())
    assert fields_for(report, ACCEPT) == []
    assert fields_for(report, REJECT) == []


def test_multiple_number_switches_422():
    b = make_doc()
    b.heading("第四章", level=1)
    b.heading("目标", level=2, bookmark="h")
    b.paragraph(
        b.fld("begin"), b.instr(" REF h \\n \\w "), b.fld("separate"),
        b.text_run("1"), b.fld("end"),
    )
    with pytest.raises(AuditError) as exc:
        audit_document(b.build())
    assert exc.value.code == "malformed_field"


def test_unknown_switch_422():
    b = make_doc()
    b.heading("第四章", level=1)
    b.heading("目标", level=2, bookmark="h")
    b.paragraph(
        b.fld("begin"), b.instr(" REF h \\z "), b.fld("separate"),
        b.text_run("1"), b.fld("end"),
    )
    with pytest.raises(AuditError) as exc:
        audit_document(b.build())
    assert exc.value.code == "unsupported_ref_switch"
