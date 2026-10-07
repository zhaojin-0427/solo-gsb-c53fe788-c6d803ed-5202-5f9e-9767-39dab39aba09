"""Main audit pipeline: physical paragraphs, accept/reject views, numbers."""

from __future__ import annotations

from dataclasses import dataclass, field

from .errors import AuditError
from .fields import parse_instruction
from .namespaces import w
from .numbering import Counter, Numbering, render_full, render_own, render_relative
from .package import load_package
from .revisions import validate_no_unsupported_revisions
from .styles import Styles
from .tokens import tokenize_paragraph

ACCEPT = "accept"
REJECT = "reject"


# ============================================================ data models
@dataclass
class AuditedField:
    ordinal: int
    instr: RefInstruction
    cached: str  # stored per view key in PhysicalParagraph


@dataclass
class PhysicalParagraph:
    ordinal: int  # global depth-first index across body and tables
    p: object
    style_id: str | None
    direct_num: tuple[int, int] | None  # (numId, ilvl) from w:pPr/w:numPr
    pilcrow: object  # None / w:ins / w:del
    fields: list  # list[FieldSpan]
    bookmark_starts: list  # list[(name, wrapper)]
    content_wrappers: list  # w:ins / w:del elements covering run content
    xpath: str
    table_path: str | None
    snippet: str


@dataclass
class LogicalParagraph:
    physicals: list  # merged PhysicalParagraph list
    host: PhysicalParagraph
    ordinal: int  # logical index inside the view
    heading_level: int | None = None
    num_id: int | None = None
    ilvl: int | None = None
    values: dict = field(default_factory=dict)


# ============================================================== parsing
def collect_physical_paragraphs(body, styles: Styles):
    paragraphs = []
    idx = 0
    for p in body.iter(w("p")):
        # Paragraphs inside text boxes / shapes are not part of the main
        # numbered flow (Word does not number them as body headings).
        if _has_ancestor(p, w("txbxContent")):
            continue
        tokens = tokenize_paragraph(p)
        style_id = None
        direct_num = None
        p_pr = tokens.p_pr
        if p_pr is not None:
            p_style = p_pr.find(w("pStyle"))
            if p_style is not None:
                style_id = p_style.get(w("val"))
            num_pr = p_pr.find(w("numPr"))
            if num_pr is not None:
                num_id_el = num_pr.find(w("numId"))
                ilvl_el = num_pr.find(w("ilvl"))
                if num_id_el is not None:
                    num_id = int(num_id_el.get(w("val")))
                    ilvl = int(ilvl_el.get(w("val"), "0")) if ilvl_el is not None else 0
                    direct_num = (num_id, ilvl)

        starts = [
            (bm.name, bm.wrapper) for bm in tokens.bookmarks if bm.kind == "start"
        ]

        # Plain-text snippet (ignoring revision view): normal + ins + del text.
        snippet = "".join(tokens.text_by_wrapper.values())

        paragraphs.append(
            PhysicalParagraph(
                ordinal=idx,
                p=p,
                style_id=style_id,
                direct_num=direct_num,
                pilcrow=tokens.pilcrow,
                fields=tokens.fields,
                bookmark_starts=starts,
                content_wrappers=tokens.content_wrappers,
                xpath=_element_path(p),
                table_path=_table_path(p),
                snippet=_clean_snippet(snippet),
            )
        )
        idx += 1
    return paragraphs


def _clean_snippet(text, limit=120):
    text = text.replace("\t", " ").replace("\n", " ").replace("\xa0", " ")
    text = " ".join(text.split())
    return text[:limit]


def _local(tag):
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) and "}" in tag else tag


def _has_ancestor(el, tag_name):
    cur = el.getparent()
    while cur is not None:
        if cur.tag == tag_name:
            return True
        cur = cur.getparent()
    return False


def _element_path(el):
    """Stable namespace-free XPath with positional predicates."""
    parts = []
    cur = el
    while cur is not None:
        parent = cur.getparent()
        local = _local(cur.tag)
        if parent is None:
            parts.append(local)
        else:
            same = [c for c in parent if isinstance(c.tag, str) and c.tag == cur.tag]
            if len(same) == 1:
                parts.append(local)
            else:
                parts.append(f"{local}[{same.index(cur) + 1}]")
        cur = parent
    return "/" + "/".join(reversed(parts))


def _table_path(p):
    """If the paragraph sits in a table, return its structural path."""
    cur = p.getparent()
    chain = []
    while cur is not None:
        local = _local(cur.tag)
        if local in ("tc", "tr", "tbl", "body", "document"):
            chain.append(local)
        cur = cur.getparent()
    if "tbl" not in chain:
        return None
    return "/".join(reversed(chain))


# ================================================================ views
def wrapper_view(wrapper, view):
    """Whether content inside ``w:ins``/``w:del`` is visible in the view."""
    if wrapper is None:
        return True
    local = _local(wrapper.tag)
    if view == ACCEPT:
        return local != "del"
    return local != "ins"


def paragraph_content_visible(phys, view):
    """Whether the paragraph carries any visible run content in the view.

    All runs, fields and bookmark markers of a tracked inserted/deleted
    paragraph live inside a single run-level revision wrapper.  When that
    wrapper is invisible the whole paragraph disappears from the view (its
    paragraph mark is handled separately for pilcrow merges).
    """
    wrappers = phys.content_wrappers
    if not wrappers:
        return True
    return any(wrapper_view(w, view) for w in wrappers)


def pilcrow_survives(pilcrow, view):
    """Whether the paragraph mark itself survives (keeps paragraph alive)."""
    if pilcrow is None:
        return True
    local = _local(pilcrow.tag)
    if view == ACCEPT:
        # Accepting a deleted pilcrow removes it; accepting an inserted keeps it.
        return local != "del"
    # Rejecting: deletion undone -> survives; insertion undone -> gone.
    return local != "ins"


def build_view(physicals, view):
    """Group physical paragraphs into the logical paragraphs of a view."""
    logical = []
    current = []
    logical_idx = 0

    def flush():
        nonlocal logical_idx, current
        if not current:
            return
        host = current[-1]  # all surviving marks are at the tail
        logical.append(
            LogicalParagraph(
                physicals=list(current),
                host=host,
                ordinal=logical_idx,
                heading_level=None,
                num_id=None,
                ilvl=None,
            )
        )
        logical_idx += 1
        current = []

    for phys in physicals:
        survives = pilcrow_survives(phys.pilcrow, view)
        current.append(phys)
        if survives:
            flush()
    flush()
    return logical


# ============================================================ resolution
def resolve_numbering(phys: PhysicalParagraph, styles: Styles, numbering: Numbering):
    """Return (num_id, ilvl) effective for the paragraph, or None."""
    if phys.direct_num is not None:
        num_id, ilvl = phys.direct_num
        if num_id == 0:  # w:numId w:val=0 cancels list numbering
            return None
        return num_id, ilvl
    if phys.style_id is not None:
        declared = styles.numbering(phys.style_id)
        if declared is not None:
            num_id, ilvl = declared
            if num_id == 0:
                return None
            return num_id, ilvl
        # A w:lvl/w:pStyle inside an abstractNum links the paragraph style
        # to that numbering level (Word "style-linked" numbering).
        linked = numbering.style_link(phys.style_id)
        if linked is not None:
            return linked
    return None


def is_heading(phys: PhysicalParagraph, styles: Styles):
    if phys.style_id is None:
        return None
    return styles.heading_level(phys.style_id)


# ============================================================ auditing
def audit_document(data: bytes):
    pkg = load_package(data)
    doc = pkg.document()
    styles_root = pkg.styles()
    numbering_root = pkg.numbering()
    validate_no_unsupported_revisions(doc, numbering_root)

    styles = Styles(styles_root)
    numbering = Numbering(numbering_root)

    body = doc.find(w("body"))
    if body is None:
        raise AuditError("corrupt_docx", "主文档缺少 w:body")

    physicals = collect_physical_paragraphs(body, styles)
    _validate_bookmarks_unique(physicals)
    _validate_fields(physicals)

    results = {}
    for view in (ACCEPT, REJECT):
        logical = build_view(physicals, view)
        results[view] = audit_view(
            view, logical, physicals, styles, numbering
        )
    return results


def _validate_bookmarks_unique(physicals):
    seen = {}
    for phys in physicals:
        for name, _wrapper in phys.bookmark_starts:
            if name == "":
                continue
            if name in seen:
                raise AuditError(
                    "duplicate_bookmark",
                    f"书签 \"{name}\" 重复定义（段落 {seen[name] + 1} 与 "
                    f"{phys.ordinal + 1}）",
                )
            seen[name] = phys.ordinal


def _validate_fields(physicals):
    for phys in physicals:
        for span in phys.fields:
            span.parsed = parse_instruction(span.instr)


def audit_view(view, logical, physicals, styles: Styles, numbering: Numbering):
    # ------------------------------------------------ view state per p
    logical_of = {}
    for lp in logical:
        for member in lp.physicals:
            logical_of[member.ordinal] = lp

    def phys_present(phys):
        """Whether the physical paragraph contributes to this view.

        * host paragraph (surviving paragraph mark): present iff its run
          content is visible;
        * non-host member merged in via a disappearing pilcrow: present iff
          its own run content is visible (accept deleted-pilcrow keeps both
          pieces; reject inserted-pilcrow drops an inserted first paragraph
          whose host it becomes nowhere).
        """
        lp = logical_of.get(phys.ordinal)
        if lp is None:
            return False
        return paragraph_content_visible(phys, view)

    # Visible field identity set.
    visible_field_phys = set()
    for lp in logical:
        for phys in lp.physicals:
            if not phys_present(phys):
                continue
            for span in phys.fields:
                if wrapper_view(span.wrapper, view):
                    visible_field_phys.add((phys.ordinal, span.ordinal))

    # ------------------------------------ numbering in physical flow order
    counter = Counter(numbering)
    # Physical ordinals that render an auto-numbered Heading 1-3 in this view,
    # mapped to (num_id, ilvl, values, heading_level).
    numbered_phys: dict[int, tuple] = {}
    for phys in physicals:
        if not phys_present(phys):
            continue
        heading = is_heading(phys, styles)
        resolved = resolve_numbering(phys, styles, numbering)
        if resolved is None:
            continue
        num_id, ilvl = resolved
        values = counter.advance(num_id, ilvl)
        if heading in (1, 2, 3):
            numbered_phys[phys.ordinal] = (num_id, ilvl, values, heading)

    # ------------------------------------ visible bookmark -> numbered p
    bookmark_targets: dict[str, tuple] = {}
    for lp in logical:
        host_info = numbered_phys.get(lp.host.ordinal)
        for phys in lp.physicals:
            if not phys_present(phys):
                continue
            for name, wrapper in phys.bookmark_starts:
                if not wrapper_view(wrapper, view):
                    continue
                info = numbered_phys.get(phys.ordinal)
                # Bookmark on a heading merged into a surviving host through
                # a deleted/inserted pilcrow resolves to that heading's own
                # number when present, otherwise the host number.
                if info is None and phys is not lp.host:
                    info = host_info
                if info is None:
                    continue
                if name in bookmark_targets:
                    raise AuditError(
                        "duplicate_bookmark", f"本视图中书签 \"{name}\" 重复"
                    )
                bookmark_targets[name] = info

    # ------------------------------------ evaluate visible REF fields
    findings = []
    for phys in physicals:
        if not phys_present(phys):
            continue
        for span in phys.fields:
            if (phys.ordinal, span.ordinal) not in visible_field_phys:
                continue
            parsed = getattr(span, "parsed", None)
            if (
                parsed is None
                or parsed.field_type != "REF"
                or parsed.number_switch is None
            ):
                continue

            cached = span.cached
            host_lp = logical_of.get(phys.ordinal)
            expected, target = _expected_value(
                switch=parsed.number_switch,
                bookmark=parsed.bookmark,
                bookmark_targets=bookmark_targets,
                field_phys_ordinal=phys.ordinal,
                numbered_phys=numbered_phys,
                physicals=physicals,
                numbering=numbering,
            )

            status = "missing_target" if target is None else (
                "correct" if cached == expected else "stale"
            )

            findings.append(
                {
                    "position": {
                        "paragraph_index": phys.ordinal,
                        "field_index": span.ordinal,
                        "logical_paragraph_index": (
                            host_lp.ordinal if host_lp else None
                        ),
                        "xpath": phys.xpath,
                        "table_path": phys.table_path,
                        "paragraph_text": phys.snippet,
                    },
                    "bookmark": parsed.bookmark,
                    "switch": parsed.number_switch,
                    "current_value": cached,
                    "expected_value": expected,
                    "status": status,
                }
            )

    findings.sort(
        key=lambda f: (
            f["position"]["paragraph_index"],
            f["position"]["field_index"],
        )
    )
    return {"fields": findings}


def _expected_value(
    *,
    switch,
    bookmark,
    bookmark_targets,
    field_phys_ordinal,
    numbered_phys,
    physicals,
    numbering,
):
    target = bookmark_targets.get(bookmark)
    if target is None:
        return None, None
    num_id, ilvl, values, _heading = target

    if switch == "n":
        return render_own(numbering, num_id, ilvl, values), target
    if switch == "w":
        return render_full(numbering, num_id, ilvl, values), target

    # --- r switch: nearest preceding visible numbered heading in flow.
    context = None
    for ordinal, info in numbered_phys.items():
        if ordinal < field_phys_ordinal:
            context = (ordinal, info)
    if context is None:
        raise AuditError(
            "no_r_context",
            f"REF 域 \\r（书签 \"{bookmark}\"，段落 {field_phys_ordinal + 1}）"
            "之前不存在可见的自动编号标题，无法求相对编号",
        )
    ctx_ordinal, (ctx_num_id, _ctx_ilvl, ctx_values, _h) = context
    rendered = render_relative(
        numbering, num_id, ilvl, values, ctx_num_id, ctx_values
    )
    return rendered, target
