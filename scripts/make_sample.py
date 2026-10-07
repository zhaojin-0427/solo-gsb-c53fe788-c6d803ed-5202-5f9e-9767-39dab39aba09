"""Generate the demo DOCX used in the README.

Run::

    python scripts/make_sample.py

Produces ``samples/sample.docx`` -- a clause document with Heading 1-3
auto-numbered paragraphs (body + table), bookmarked targets, cross-run REF
fields (\\n / \\r / \\w), one stale field, a tracked insertion, a tracked
deletion and a paragraph-mark deletion.
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "tests"))

from builder import DocxBuilder


def build_sample() -> bytes:
    b = DocxBuilder()
    b.add_default_heading_styles()
    # Chapter numbering starts at 4 to mirror the worked example (4.3.1 -> 4.5.2).
    b.add_multilevel_numbering(num_id=1, start_overrides={0: 4})

    b.heading("范围", level=1)  # 4
    b.heading("通用要求", level=2)  # 4.1
    b.heading("术语和定义", level=2)  # 4.2
    b.heading("编号上下文", level=3, bookmark="bm_ctx")  # 4.2.1

    # A REF \r field located right after 4.2.1, forward-referencing 4.5.2.
    # Cached value "5.1" is stale after the clause was renumbered (5.2).
    b.paragraph(
        b.text_run("见后文 "),
        *b.ref_field_runs("bm_t2", "r", "5.1", split=True),
        b.text_run("（接受视图应值 5.2，缓存已过期）；"),
        b.text_run("完整条号 "),
        *b.ref_field_runs("bm_t2", "w", "4.5.2"),
        b.text_run("；本级编号 "),
        *b.ref_field_runs("bm_t2", "n", "2"),
    )

    b.heading("职责", level=2)  # 4.3
    b.heading("详细设计", level=2)  # 4.4
    b.heading("接口", level=2)  # 4.5
    b.heading("请求", level=3)  # 4.5.1

    # Target heading whose cached reference is deliberately stale.
    b.heading("响应结构", level=3, bookmark="bm_t2")  # 4.5.2
    b.paragraph(
        b.text_run("旧引用值 "),
        *b.ref_field_runs("bm_t1", "w", "4.9.9"),
        b.text_run("（该编号已过期）"),
    )

    b.heading("数据模型", level=2, bookmark="bm_t1")  # 4.6

    # Table: bookmarked heading in a table cell, plus a reference cell.
    b.table(
        [
            [
                [
                    {"runs": [b.bookmark_run("bm_tab")],
                     "style": "Heading3", "num_id": 1, "ilvl": 2},
                ],
                [
                    {"runs": [
                        b.text_run("表内引用 "),
                        *b.ref_field_runs("bm_tab", "w", "4.6.1"),
                    ]},
                ],
            ],
        ]
    )

    # ----- tracked insertion: new heading + new reference (accept only) -----
    b.heading(
        "新增条款",
        level=2,
        runs=[b.wrap("ins", "reviewer", b.bookmark_run("bm_new"),
                     ("text", "新增条款"))],
    )
    b.paragraph(
        b.wrap(
            "ins",
            "reviewer",
            b.fld("begin"), b.instr(" REF bm_new \\n "), b.fld("separate"),
            ("text", "7"), b.fld("end"),
        ),
    )

    # ----- tracked deletion: old heading exists only in the reject view -----
    b.heading(
        "删除条款",
        level=2,
        runs=[b.wrap("del", "reviewer", b.bookmark_run("bm_old"),
                     ("deltext", "删除条款"))],
    )
    b.paragraph(
        b.text_run("对旧条款的引用 "),
        *b.ref_field_runs("bm_old", "w", "4.8"),
    )

    # ----- paragraph-mark deletion: heading merges into the following body
    # paragraph in the accept view.
    b.heading("段落标记删除条款", level=3, bookmark="bm_pilcrow",
              pilcrow=("del", "reviewer"))
    b.paragraph(
        b.text_run("（接受视图中本句并入上方标题）引用 "),
        *b.ref_field_runs("bm_pilcrow", "w", "4.9.1"),
    )

    return b.build()


def main():
    out = os.path.join(os.path.dirname(__file__), "..", "samples", "sample.docx")
    with open(out, "wb") as f:
        f.write(build_sample())
    print(f"wrote {os.path.abspath(out)}")


if __name__ == "__main__":
    main()
