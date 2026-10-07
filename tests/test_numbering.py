"""Unit tests for the numbering engine."""
import pytest
from lxml import etree

from builder import DocxBuilder
from docx_audit.numbering import (
    Counter,
    Numbering,
    format_value,
    render_full,
    render_own,
    render_relative,
    Level,
)
from docx_audit.errors import AuditError


def numbering(**kwargs):
    b = DocxBuilder()
    b.add_multilevel_numbering(1, **kwargs)
    root = etree.fromstring(b._numbering_xml())
    return Numbering(root)


def values_sequence(n: Numbering, levels):
    c = Counter(n)
    out = []
    for ilvl in levels:
        out.append(c.advance(1, ilvl))
    return out


def test_basic_multilevel_sequence():
    n = numbering()
    seq = values_sequence(n, [0, 1, 1, 2, 1])
    assert seq[0] == {0: 1}
    assert seq[1] == {0: 1, 1: 1}
    assert seq[2] == {0: 1, 1: 2}
    assert seq[3] == {0: 1, 1: 2, 2: 1}
    # level1 advancement restarts level2
    assert seq[4] == {0: 1, 1: 3}


def test_start_override_one_shot():
    n = numbering(start_overrides={0: 4})
    c = Counter(n)
    assert c.advance(1, 0) == {0: 4}
    assert c.advance(1, 0) == {0: 5}
    # A second instance consumes its own override independently -- but the
    # override was consumed by instance 1; ensure same instance keeps going.
    c2 = Counter(n)
    assert c2.advance(1, 0) == {0: 4}


def test_separate_num_instances_are_independent():
    b = DocxBuilder()
    b.add_multilevel_numbering(1)
    b.add_multilevel_numbering(2)
    root = etree.fromstring(b._numbering_xml())
    n = Numbering(root)
    c = Counter(n)
    assert c.advance(1, 0) == {0: 1}
    assert c.advance(2, 0) == {0: 1}
    assert c.advance(1, 0) == {0: 2}
    assert c.advance(2, 0) == {0: 2}


def test_start_values_other_than_one():
    n = numbering(starts=[3, 2, 5])
    c = Counter(n)
    assert c.advance(1, 0) == {0: 3}
    assert c.advance(1, 1) == {0: 3, 1: 2}
    assert c.advance(1, 2) == {0: 3, 1: 2, 2: 5}


def test_lvl_restart_policy_1():
    # level2 restart=1 -> restarted only when level1 (its direct parent)
    # advances, not when level0 advances.
    n = numbering(restarts=[0, 0, 1])
    seq = values_sequence(n, [0, 1, 2, 0, 1, 2, 2, 1, 2])
    # level0 advances (item 3): level2 policy=1, (2-0)=2 > 1 -> NOT reset
    assert seq[3] == {0: 2}
    assert seq[4] == {0: 2, 1: 1}
    assert seq[5] == {0: 2, 1: 1, 2: 2}  # continues from previous value 1
    assert seq[6] == {0: 2, 1: 1, 2: 3}
    # level1 advances (item 7): (2-1)=1 <= 1 -> reset
    assert seq[7] == {0: 2, 1: 2}
    assert seq[8] == {0: 2, 1: 2, 2: 1}


def test_render_full_own_relative():
    n = numbering()
    vals = {0: 4, 1: 5, 2: 2}
    ctx = {0: 4, 1: 3, 2: 1}
    assert render_full(n, 1, 2, vals) == "4.5.2"
    assert render_own(n, 1, 2, vals) == "2"
    assert render_relative(n, 1, 2, vals, 1, ctx) == "5.2"


def test_relative_different_num_instances_is_full():
    n = numbering()
    vals = {0: 2, 1: 1}
    ctx = {0: 9}
    assert render_relative(n, 1, 1, vals, 2, ctx) == "2.1"


def test_relative_no_common_prefix():
    n = numbering()
    assert render_relative(n, 1, 0, {0: 5}, 1, {0: 2}) == "5"
    assert render_relative(n, 1, 1, {0: 5, 1: 2}, 1, {0: 4, 1: 9}) == "5.2"


def test_relative_same_number_is_full():
    n = numbering()
    vals = {0: 4, 1: 3, 2: 1}
    assert render_relative(n, 1, 2, vals, 1, dict(vals)) == "4.3.1"


def test_roman_and_letter_formats():
    n = numbering(texts=["%1)", "%1.%2", "%1.%2-%3"],
                  fmts=["decimal", "upperRoman", "lowerLetter"])
    vals = {0: 4, 1: 2, 2: 3}
    assert render_full(n, 1, 2, vals) == "4.II-c"
    # Relative rendering keeps lvlText separators from the target level.
    ctx = {0: 4, 1: 1, 2: 9}
    assert render_relative(n, 1, 2, vals, 1, ctx) == "II-c"


def test_is_lgl_forces_decimal_ancestors():
    n = numbering(fmts=["upperRoman", "upperRoman", "decimal"], is_lgl=True)
    vals = {0: 4, 1: 2, 2: 7}
    assert render_full(n, 1, 2, vals) == "4.2.7"


def test_custom_lvl_text_separators():
    n = numbering(texts=["第%1章", "%1（%2）", "%1-%2-%3"])
    vals = {0: 4, 1: 3, 2: 1}
    assert render_full(n, 1, 2, vals) == "4-3-1"
    ctx = {0: 4, 1: 2, 2: 5}
    assert render_relative(n, 1, 2, vals, 1, ctx) == "3-1"


def test_unsupported_format_raises():
    n = numbering(fmts=["bullet", "decimal", "decimal"])
    c = Counter(n)
    with pytest.raises(AuditError) as exc:
        c.advance(1, 0)
    assert exc.value.code == "unsupported_number_format"


def test_ordinal_format():
    assert format_value(Level(ilvl=0, fmt="ordinal"), 22) == "22nd"
    assert format_value(Level(ilvl=0, fmt="decimalZero", start=1), 7) == "07"
    assert format_value(Level(ilvl=0, fmt="upperLetter"), 27) == "AA"
