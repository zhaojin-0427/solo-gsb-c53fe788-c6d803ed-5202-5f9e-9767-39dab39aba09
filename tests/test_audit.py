"""端到端测试：通过 TestClient 调用 POST /audit。"""

import pytest
from fastapi.testclient import TestClient

from app.main import app
from tests.docx_factory import (
    abstract,
    custom_style,
    doc,
    heading,
    heading_style,
    lvl_override,
    make_docx,
    num,
    numbering_xml,
    para,
    ppr_xml,
    ref_para,
    run,
    styles_xml,
)

client = TestClient(app)

# ---------------------------------------------------------------- 通用部件

STYLES = styles_xml(
    '<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>',
    heading_style("Heading1", num_id=10, ilvl=0),
    heading_style("Heading2", num_id=10, ilvl=1),
    heading_style("Heading3", num_id=10, ilvl=2),
)

NUMBERING = numbering_xml(
    abstract(
        1,
        [
            (0, "decimal", "%1."),
            (1, "decimal", "%1.%2"),
            (2, "decimal", "%1.%2.%3"),
        ],
    ),
    num(10, 1),
)


def post(docx: bytes):
    return client.post(
        "/audit",
        content=docx,
        headers={"Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"},
    )


def post_ok(docx: bytes):
    resp = post(docx)
    assert resp.status_code == 200, resp.text
    return resp.json()["views"]


def post_422(docx: bytes, code: str):
    resp = post(docx)
    assert resp.status_code == 422, resp.text
    assert resp.json()["detail"]["error"] == code


# ---------------------------------------------------------------- 基础核对


def test_basic_correct_stale_missing():
    docx = make_docx(
        doc(
            heading("第一章", "Heading1", bm="_TocA", bm_id=1),
            heading("第一节", "Heading2", bm="_TocB", bm_id=2),
            ref_para("_TocA", "\\w", current="1"),
            ref_para("_TocB", "\\w", current="1.2"),   # 应为 1.1 → 过期
            ref_para("_TocB", "\\n", current="1"),
            ref_para("_TocNone", "\\w", current="?"),  # 目标缺失
        ),
        STYLES,
        NUMBERING,
    )
    views = post_ok(docx)
    fields = views["accept"]["fields"]
    assert [f["status"] for f in fields] == ["correct", "stale", "correct", "target_missing"]
    assert fields[0]["expected"] == "1"
    assert fields[1]["expected"] == "1.1"
    assert fields[1]["current"] == "1.2"
    assert fields[2]["expected"] == "1"  # \n 取当前级（第 2 级）编号
    assert fields[3]["expected"] is None
    # 拒绝视图与接受视图一致（无修订）
    assert views["reject"]["fields"] == fields


def test_ref_without_switch_uses_bookmark_text():
    docx = make_docx(
        doc(
            heading("总则", "Heading1", bm="_TocA", bm_id=1),
            ref_para("_TocA", current="总则"),
            ref_para("_TocA", current="第一章"),  # 过期
        ),
        STYLES,
        NUMBERING,
    )
    fields = post_ok(docx)["accept"]["fields"]
    assert [f["status"] for f in fields] == ["correct", "stale"]
    assert fields[0]["expected"] == "总则"


def test_relative_ref_example():
    """任务书示例：源标题 4.3.1 指向 4.5.2，\\r 应显示 5.2。"""
    body = []
    for i in range(1, 5):  # 4 个一级标题
        body.append(heading(f"第{i}章", "Heading1", bm=f"_TocH{i}", bm_id=i))
    # 第 4 章下：4.1 ~ 4.3
    body.append(heading("4.1 节", "Heading2"))
    body.append(heading("4.2 节", "Heading2"))
    body.append(heading("4.3 节", "Heading2", bm="_Toc43", bm_id=43))
    body.append(heading("4.3.1 条", "Heading3", bm="_Toc431", bm_id=431))
    # 引用点：4.3.1 标题之后、4.4 之前 → 上下文为 4.3.1
    body.append(ref_para("_Toc452", "\\r", current="5.2"))
    body.append(ref_para("_Toc452", "\\n", current="2"))
    body.append(ref_para("_Toc452", "\\w", current="4.5.2"))
    body.append(heading("4.4 节", "Heading2"))
    body.append(heading("4.5 节", "Heading2", bm="_Toc45", bm_id=45))
    body.append(heading("4.5.1 条", "Heading3"))
    body.append(heading("4.5.2 条", "Heading3", bm="_Toc452", bm_id=452))
    docx = make_docx(doc(*body), STYLES, NUMBERING)
    fields = post_ok(docx)["accept"]["fields"]
    assert [f["expected"] for f in fields] == ["5.2", "2", "4.5.2"]
    assert all(f["status"] == "correct" for f in fields)


def test_relative_ref_to_prefix_falls_back_to_full():
    """\\r 目标编号是上下文编号的前缀时，回退为完整编号。"""
    docx = make_docx(
        doc(
            heading("第一章", "Heading1", bm="_TocH1", bm_id=1),
            heading("1.1 节", "Heading2", bm="_Toc11", bm_id=2),
            ref_para("_TocH1", "\\r", current="1"),   # 目标是上下文的前缀 → 完整编号
        ),
        STYLES,
        NUMBERING,
    )
    fields = post_ok(docx)["accept"]["fields"]
    assert fields[0]["expected"] == "1"
    assert fields[0]["status"] == "correct"


def test_missing_r_context_is_422():
    docx = make_docx(
        doc(
            para(run("正文")),  # 无编号标题
            ref_para("_TocA", "\\r", current="1"),
            heading("第一章", "Heading1", bm="_TocA", bm_id=1),
        ),
        STYLES,
        NUMBERING,
    )
    post_422(docx, "missing_r_context")


# ---------------------------------------------------------------- 修订视图


def test_deleted_heading_only_in_reject_view():
    """删除的标题：接受视图目标缺失，拒绝视图正常，互不借用。"""
    deleted_heading = (
        "<w:p>"
        + ppr_xml("Heading2", mark="del")
        + '<w:del w:id="2" w:author="t" w:date="2026-01-01">'
        + '<w:bookmarkStart w:id="9" w:name="_TocDel"/>'
        + "<w:r><w:delText>被删节</w:delText></w:r>"
        + '<w:bookmarkEnd w:id="9"/>'
        + "</w:del>"
        + "</w:p>"
    )
    docx = make_docx(
        doc(
            heading("第一章", "Heading1", bm="_TocH1", bm_id=1),
            heading("1.1 节", "Heading2", bm="_Toc11", bm_id=3),
            deleted_heading,  # 1.2（被删除）
            heading("1.3 节", "Heading2", bm="_Toc13", bm_id=4),
            ref_para("_TocDel", "\\w", current="1.2"),
            ref_para("_Toc13", "\\w", current="1.3"),
        ),
        STYLES,
        NUMBERING,
    )
    views = post_ok(docx)
    acc = views["accept"]["fields"]
    rej = views["reject"]["fields"]
    # 接受视图：被删标题不存在 → 目标缺失；1.3 上移为 1.2
    assert acc[0]["status"] == "target_missing"
    assert acc[1]["expected"] == "1.2"
    assert acc[1]["status"] == "stale"
    # 拒绝视图：1.2 仍在；1.3 保持 1.3
    assert rej[0]["expected"] == "1.2"
    assert rej[0]["status"] == "correct"
    assert rej[1]["expected"] == "1.3"
    assert rej[1]["status"] == "correct"


def test_inserted_heading_only_in_accept_view():
    ins_heading = (
        "<w:p>"
        + ppr_xml("Heading2", mark="ins")
        + '<w:ins w:id="2" w:author="t" w:date="2026-01-01">'
        + '<w:bookmarkStart w:id="9" w:name="_TocIns"/>'
        + "<w:r><w:t>新增节</w:t></w:r>"
        + '<w:bookmarkEnd w:id="9"/>'
        + "</w:ins>"
        + "</w:p>"
    )
    docx = make_docx(
        doc(
            heading("第一章", "Heading1", bm="_TocH1", bm_id=1),
            ins_heading,  # 接受视图中的 1.1
            heading("1.2 节", "Heading2", bm="_Toc12", bm_id=3),
            ref_para("_TocIns", "\\w", current="1.1"),
            ref_para("_Toc12", "\\w", current="1.2"),
        ),
        STYLES,
        NUMBERING,
    )
    views = post_ok(docx)
    acc = views["accept"]["fields"]
    rej = views["reject"]["fields"]
    assert acc[0]["status"] == "correct"          # 1.1
    assert acc[1]["expected"] == "1.2"
    assert acc[1]["status"] == "correct"
    # 拒绝视图：新增标题不存在 → 目标缺失；原 1.2 变为 1.1
    assert rej[0]["status"] == "target_missing"
    assert rej[1]["expected"] == "1.1"
    assert rej[1]["status"] == "stale"


def test_paragraph_mark_insert_delete():
    """段落标记修订：接受视图按现状，拒绝视图合并/还原。"""
    docx = make_docx(
        doc(
            para(run("第一章"), style="Heading1", mark="ins"),
            heading("第二章", "Heading1", bm="_TocH2", bm_id=7),
            para(run("1.1 节"), style="Heading2", mark="del"),
            para(run("内容"), style="Normal"),
            ref_para("_TocH2", "\\w", current="2"),
        ),
        STYLES,
        NUMBERING,
    )
    views = post_ok(docx)
    acc = views["accept"]["fields"]
    rej = views["reject"]["fields"]
    # 接受视图：第一章、第二章（=2）、1.1 节（其段落标记被删，与“内容”合并，
    # 合并后取 Normal 样式 → 不再编号）
    assert acc[0]["expected"] == "2"
    assert acc[0]["status"] == "correct"
    # 拒绝视图：插入的“第一章”段落不存在 → 第二章编号为 1
    assert rej[0]["expected"] == "1"
    assert rej[0]["status"] == "stale"


def test_unsupported_revision_types_422():
    # 移动修订
    docx = make_docx(
        doc(
            "<w:p><w:moveFrom><w:r><w:t>x</w:t></w:r></w:moveFrom></w:p>"
        ),
        STYLES,
        NUMBERING,
    )
    post_422(docx, "unsupported_revision")
    # 属性修订
    docx = make_docx(
        doc(
            '<w:p><w:pPr><w:pPrChange w:id="1" w:author="t" w:date="2026-01-01">'
            "<w:pPr/></w:pPrChange></w:pPr></w:p>"
        ),
        STYLES,
        NUMBERING,
    )
    post_422(docx, "unsupported_revision")
    # 表格行插入修订
    docx = make_docx(
        doc(
            "<w:tbl><w:tr><w:trPr>"
            '<w:ins w:id="1" w:author="t" w:date="2026-01-01"/>'
            "</w:trPr><w:tc><w:p/></w:tc></w:tr></w:tbl>"
        ),
        STYLES,
        NUMBERING,
    )
    post_422(docx, "unsupported_revision")


# ---------------------------------------------------------------- 结构错误


def test_duplicate_bookmark_422():
    docx = make_docx(
        doc(
            heading("一", "Heading1", bm="_TocX", bm_id=1),
            heading("二", "Heading1", bm="_TocX", bm_id=2),
        ),
        STYLES,
        NUMBERING,
    )
    post_422(docx, "duplicate_bookmark")


def test_broken_field_422():
    # 结束符没有开始符
    docx = make_docx(
        doc(
            para(
                run("x"),
                '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
            )
        ),
        STYLES,
        NUMBERING,
    )
    post_422(docx, "broken_field")
    # 开始符未闭合
    docx = make_docx(
        doc(
            para(
                '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
                '<w:r><w:instrText> REF _Toc1 </w:instrText></w:r>',
            )
        ),
        STYLES,
        NUMBERING,
    )
    post_422(docx, "broken_field")
    # 指令文字出现在域之外
    docx = make_docx(
        doc(para('<w:r><w:instrText> REF _Toc1 </w:instrText></w:r>')),
        STYLES,
        NUMBERING,
    )
    post_422(docx, "broken_field")


def test_invalid_docx_422():
    post_422(b"not a zip at all", "invalid_docx")
    post_422(b"", "invalid_docx")


# ---------------------------------------------------------------- 跨 run 配对


def test_field_and_bookmark_spanning_runs():
    """域指令、书签跨多个 run 也能正确配对。"""
    docx = make_docx(
        doc(
            # 书签跨 run
            para(
                '<w:bookmarkStart w:id="5" w:name="_TocA"/>',
                run("第一"),
                run("章"),
                '<w:bookmarkEnd w:id="5"/>',
                style="Heading1",
            ),
            # 域指令拆成两个 run
            para(
                run("见"),
                '<w:r><w:fldChar w:fldCharType="begin"/></w:r>',
                '<w:r><w:instrText xml:space="preserve"> REF _To</w:instrText></w:r>',
                '<w:r><w:instrText xml:space="preserve">cA \\w </w:instrText></w:r>',
                '<w:r><w:fldChar w:fldCharType="separate"/></w:r>',
                run("1"),
                '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
            ),
        ),
        STYLES,
        NUMBERING,
    )
    fields = post_ok(docx)["accept"]["fields"]
    assert fields[0]["bookmark"] == "_TocA"
    assert fields[0]["expected"] == "1"
    assert fields[0]["status"] == "correct"


def test_fld_simple():
    docx = make_docx(
        doc(
            heading("第一章", "Heading1", bm="_TocA", bm_id=1),
            para(
                run("见"),
                '<w:fldSimple w:instr=" REF _TocA \\w "><w:r><w:t>1</w:t></w:r></w:fldSimple>',
            ),
        ),
        STYLES,
        NUMBERING,
    )
    fields = post_ok(docx)["accept"]["fields"]
    assert fields[0]["expected"] == "1"
    assert fields[0]["status"] == "correct"


# ---------------------------------------------------------------- 编号规则


def test_start_override():
    numbering = numbering_xml(
        abstract(1, [(0, "decimal", "%1.")]),
        num(10, 1, lvl_override(0, start_override=3)),
    )
    styles = styles_xml(heading_style("Heading1", num_id=10, ilvl=0))
    docx = make_docx(
        doc(
            heading("三", "Heading1", bm="_TocA", bm_id=1),
            ref_para("_TocA", "\\w", current="3"),
        ),
        styles,
        numbering,
    )
    fields = post_ok(docx)["accept"]["fields"]
    assert fields[0]["expected"] == "3"
    assert fields[0]["status"] == "correct"


def test_level_restart_never():
    """lvlRestart=0：二级编号不随一级标题重启。"""
    numbering = numbering_xml(
        abstract(
            1,
            [
                (0, "decimal", "%1."),
                (1, "decimal", "%1.%2", None, 0),  # 永不重启
            ],
        ),
        num(10, 1),
    )
    styles = styles_xml(
        heading_style("Heading1", num_id=10, ilvl=0),
        heading_style("Heading2", num_id=10, ilvl=1),
    )
    docx = make_docx(
        doc(
            heading("一", "Heading1"),
            heading("1.1", "Heading2"),
            heading("二", "Heading1"),
            heading("2.2", "Heading2", bm="_TocB", bm_id=2),
            ref_para("_TocB", "\\w", current="2.2"),
        ),
        styles,
        numbering,
    )
    fields = post_ok(docx)["accept"]["fields"]
    assert fields[0]["expected"] == "2.2"
    assert fields[0]["status"] == "correct"


def test_style_inheritance_and_pstyle_link():
    """自定义样式基于 Heading1（继承编号）；另一样式靠 pStyle 链接编号。"""
    numbering = numbering_xml(
        abstract(1, [(0, "decimal", "%1.", None, None, "LinkedHead")]),
        num(10, 1),
    )
    styles = styles_xml(
        heading_style("Heading1", num_id=10, ilvl=0),
        custom_style("MyHead", "My Heading", "Heading1"),
        '<w:style w:type="paragraph" w:styleId="LinkedHead">'
        '<w:name w:val="Linked Head"/></w:style>',
    )
    docx = make_docx(
        doc(
            heading("一", "MyHead", bm="_TocA", bm_id=1),      # 继承 Heading1 → 1
            heading("二", "LinkedHead", bm="_TocB", bm_id=2),  # pStyle 链接 → 2
            ref_para("_TocB", "\\w", current="2"),
        ),
        styles,
        numbering,
    )
    fields = post_ok(docx)["accept"]["fields"]
    assert fields[0]["expected"] == "2"
    assert fields[0]["status"] == "correct"


def test_separate_list_instances():
    """两个 numId 指向同一 abstractNum：计数器相互独立。"""
    numbering = numbering_xml(
        abstract(1, [(0, "decimal", "%1.")]),
        num(10, 1) + num(11, 1),
    )
    styles = styles_xml(heading_style("Heading1", num_id=10, ilvl=0))
    docx = make_docx(
        doc(
            heading("一", "Heading1", bm="_TocA", bm_id=1),
            # 另一个列表实例（直接 numPr 指向 numId=11）
            para(run("列表项"), style="Normal", num_id=11, ilvl=0),
            para(run("列表项"), style="Normal", num_id=11, ilvl=0),
            heading("二", "Heading1", bm="_TocB", bm_id=2),
            ref_para("_TocB", "\\w", current="2"),
        ),
        styles,
        numbering,
    )
    fields = post_ok(docx)["accept"]["fields"]
    assert fields[0]["expected"] == "2"  # 实例 10 不受实例 11 影响
    assert fields[0]["status"] == "correct"


def test_number_formats():
    """\\n 按当前级格式取号：罗马数字与字母。"""
    numbering = numbering_xml(
        abstract(
            1,
            [
                (0, "upperRoman", "%1."),
                (1, "lowerLetter", "%1.%2"),
            ],
        ),
        num(10, 1),
    )
    styles = styles_xml(
        heading_style("Heading1", num_id=10, ilvl=0),
        heading_style("Heading2", num_id=10, ilvl=1),
    )
    docx = make_docx(
        doc(
            heading("一", "Heading1", bm="_TocH1", bm_id=1),
            heading("1.1", "Heading2"),
            heading("1.2", "Heading2", bm="_TocB", bm_id=2),
            ref_para("_TocB", "\\n", current="b"),   # 当前级 lowerLetter
            ref_para("_TocB", "\\w", current="I.b"),  # 完整条号
        ),
        styles,
        numbering,
    )
    fields = post_ok(docx)["accept"]["fields"]
    assert fields[0]["expected"] == "b"
    assert fields[1]["expected"] == "I.b"
    assert all(f["status"] == "correct" for f in fields)


# ---------------------------------------------------------------- 表格


def test_headings_and_refs_in_tables():
    docx = make_docx(
        doc(
            "<w:tbl><w:tr><w:tc>"
            + heading("第一章", "Heading1", bm="_TocA", bm_id=1)
            + "</w:tc></w:tr></w:tbl>",
            ref_para("_TocA", "\\w", current="1"),
        ),
        STYLES,
        NUMBERING,
    )
    fields = post_ok(docx)["accept"]["fields"]
    assert fields[0]["expected"] == "1"
    assert fields[0]["status"] == "correct"


def test_field_visibility_across_views():
    """插入/删除的域分别只出现在接受/拒绝视图。"""
    docx = make_docx(
        doc(
            heading("第一章", "Heading1", bm="_TocA", bm_id=1),
            # 插入的 REF（仅接受视图可见）
            (
                "<w:p>"
                '<w:ins w:id="1" w:author="t" w:date="2026-01-01">'
                '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                '<w:r><w:instrText xml:space="preserve"> REF _TocA \\w </w:instrText></w:r>'
                '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                "<w:r><w:t>1</w:t></w:r>"
                '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
                "</w:ins>"
                "</w:p>"
            ),
            # 删除的 REF（仅拒绝视图可见）
            (
                "<w:p>"
                '<w:del w:id="2" w:author="t" w:date="2026-01-01">'
                '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                '<w:r><w:instrText xml:space="preserve"> REF _TocA \\w </w:instrText></w:r>'
                '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                "<w:r><w:delText>9</w:delText></w:r>"
                '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
                "</w:del>"
                "</w:p>"
            ),
        ),
        STYLES,
        NUMBERING,
    )
    views = post_ok(docx)
    acc = views["accept"]["fields"]
    rej = views["reject"]["fields"]
    assert len(acc) == 1 and acc[0]["status"] == "correct"
    assert len(rej) == 1 and rej[0]["current"] == "9" and rej[0]["status"] == "stale"


def test_quoted_bookmark_name_with_spaces():
    docx = make_docx(
        doc(
            heading("第一章", "Heading1", bm="My Bookmark", bm_id=1),
            ref_para('"My Bookmark"', "\\w", current="1"),
        ),
        STYLES,
        NUMBERING,
    )
    fields = post_ok(docx)["accept"]["fields"]
    assert fields[0]["bookmark"] == "My Bookmark"
    assert fields[0]["status"] == "correct"


def test_bookmark_on_unnumbered_paragraph():
    """书签在未编号段落上：\\w 目标缺失；无开关时取书签文本。"""
    docx = make_docx(
        doc(
            para(run("普通段落"), style="Normal").replace(
                "<w:p>", '<w:p><w:bookmarkStart w:id="3" w:name="_TocP"/>'
            ).replace("</w:p>", '<w:bookmarkEnd w:id="3"/></w:p>'),
            ref_para("_TocP", "\\w", current="1.1"),
            ref_para("_TocP", current="普通段落"),
        ),
        STYLES,
        NUMBERING,
    )
    fields = post_ok(docx)["accept"]["fields"]
    assert fields[0]["status"] == "target_missing"
    assert fields[1]["status"] == "correct"
    assert fields[1]["expected"] == "普通段落"
