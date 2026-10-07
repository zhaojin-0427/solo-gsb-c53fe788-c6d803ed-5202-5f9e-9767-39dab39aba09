"""WordprocessingML numbering model.

Implements the subset of the Word numbering rules needed to render the
paragraph numbers of audited headings:

* separate list instances (``w:num``) keep independent counters;
* a numbering level inherits its definition from the abstract numbering
  definition, optionally overridden by ``w:lvlOverride``/``w:lvl``;
* ``w:startOverride`` overrides the start value only for the first item of
  that level in the instance;
* ``w:lvlRestart`` controls which deeper levels are reset when a (more
  significant) level advances; omitted/0 means any more significant level
  restarts this level (the Word default);
* ``w:isLgl`` forces every more significant level to render as decimal;
* ``w:lvlText`` ``%1``..``%9`` placeholders define full clause numbers.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .errors import AuditError
from .namespaces import w

MAX_LEVELS = 9

SUPPORTED_FORMATS = {
    "decimal",
    "decimalZero",
    "lowerRoman",
    "upperRoman",
    "lowerLetter",
    "upperLetter",
    "ordinal",
    "none",
}


def _child(el, name):
    return el.find(w(name))


def _children(el, name):
    return el.findall(w(name))


def _attr(el, name, default=None):
    if el is None:
        return default
    return el.get(w(name), default)


@dataclass
class Level:
    ilvl: int
    start: int = 1
    fmt: str = "decimal"
    text: str = ""
    is_lgl: bool = False
    restart: int = 0


@dataclass
class _AbstractNum:
    levels: dict = field(default_factory=dict)  # ilvl -> Level
    style_levels: dict = field(default_factory=dict)  # styleId -> ilvl


@dataclass
class _NumOverride:
    start_override: int | None = None
    level: Level | None = None


@dataclass
class _Num:
    abstract_num_id: int | None = None
    overrides: dict = field(default_factory=dict)  # ilvl -> _NumOverride


class Numbering:
    """Parsed view of ``numbering.xml`` (may be empty)."""

    def __init__(self, numbering_root):
        self._abstracts: dict[int, _AbstractNum] = {}
        self._nums: dict[int, _Num] = {}
        self._style_to_num: dict[str, tuple[int, int]] = {}
        if numbering_root is not None:
            self._parse(numbering_root)

    # ------------------------------------------------------------------ parse
    def _parse(self, root):
        for an in _children(root, "abstractNum"):
            an_id = int(an.get(w("abstractNumId"), "-1"))
            abstract = _AbstractNum()
            for lvl_el in _children(an, "lvl"):
                level = self._parse_level(lvl_el)
                abstract.levels[level.ilvl] = level
                pstyle = lvl_el.find(w("pStyle"))
                if pstyle is not None:
                    abstract.style_levels[pstyle.get(w("val"), "")] = level.ilvl
            self._abstracts[an_id] = abstract

        for num_el in _children(root, "num"):
            num_id = int(num_el.get(w("numId"), "-1"))
            num = _Num(abstract_num_id=None)
            an_ref = num_el.find(w("abstractNumId"))
            if an_ref is not None:
                num.abstract_num_id = int(an_ref.get(w("val"), "-1"))
            for ov_el in _children(num_el, "lvlOverride"):
                ilvl = int(ov_el.get(w("ilvl"), "0"))
                override = _NumOverride()
                start_el = ov_el.find(w("startOverride"))
                if start_el is not None:
                    override.start_override = int(start_el.get(w("val"), "1"))
                lvl_el = ov_el.find(w("lvl"))
                if lvl_el is not None:
                    override.level = self._parse_level(lvl_el, ilvl)
                num.overrides[ilvl] = override
            self._nums[num_id] = num

        for num_id, num in self._nums.items():
            abstract = self._abstracts.get(num.abstract_num_id)
            if abstract is None:
                continue
            for style_id, ilvl in abstract.style_levels.items():
                self._style_to_num[style_id] = (num_id, ilvl)

    @staticmethod
    def _parse_level(lvl_el, ilvl=None):
        if ilvl is None:
            ilvl = int(lvl_el.get(w("ilvl"), "0"))
        level = Level(ilvl=ilvl)
        start_el = lvl_el.find(w("start"))
        if start_el is not None:
            level.start = int(start_el.get(w("val"), "1"))
        fmt_el = lvl_el.find(w("numFmt"))
        if fmt_el is not None:
            level.fmt = fmt_el.get(w("val"), "decimal")
        text_el = lvl_el.find(w("lvlText"))
        if text_el is not None:
            level.text = text_el.get(w("val"), "")
        if lvl_el.find(w("isLgl")) is not None:
            level.is_lgl = _attr(lvl_el.find(w("isLgl")), "val", "1") in ("1", "true")
        restart_el = lvl_el.find(w("lvlRestart"))
        if restart_el is not None:
            level.restart = int(restart_el.get(w("val"), "0"))
        return level

    # --------------------------------------------------------------- lookups
    def style_link(self, style_id):
        """numId/ilvl implied by a ``w:pStyle`` reference inside a level."""
        return self._style_to_num.get(style_id)

    def _abstract(self, num_id):
        num = self._nums.get(num_id)
        if num is None:
            raise AuditError(
                "unknown_numbering_instance",
                f"段落引用了未定义的编号实例 w:numId={num_id}",
            )
        abstract = self._abstracts.get(num.abstract_num_id)
        if abstract is None:
            raise AuditError(
                "unknown_numbering_definition",
                f"编号实例 w:numId={num_id} 指向未定义的 abstractNumId="
                f"{num.abstract_num_id}",
            )
        return num, abstract

    def resolve(self, num_id, ilvl):
        """Effective :class:`Level` for an instance/level pair."""
        if not 0 <= ilvl < MAX_LEVELS:
            raise AuditError(
                "unsupported_numbering_level", f"编号级别 ilvl={ilvl} 超出 Word 支持范围"
            )
        num, abstract = self._abstract(num_id)
        override = num.overrides.get(ilvl)
        base = abstract.levels.get(ilvl)
        if override is None or override.level is None:
            if base is None:
                raise AuditError(
                    "undefined_numbering_level",
                    f"编号级别 abstractNumId={num.abstract_num_id} ilvl={ilvl} 未定义",
                )
            return base
        if base is None:
            return override.level
        # Merge only the properties the override level actually specified.
        ov = override.level
        return Level(
            ilvl=ilvl,
            start=ov.start if _specified(ov, "start") else base.start,
            fmt=ov.fmt if _specified(ov, "fmt") else base.fmt,
            text=ov.text if _specified(ov, "text") else base.text,
            is_lgl=ov.is_lgl or base.is_lgl,
            restart=ov.restart if _specified(ov, "restart") else base.restart,
        )

    def peek_start_override(self, num_id, ilvl):
        """Return the instance/level one-shot start override (or None)."""
        num, _ = self._abstract(num_id)
        override = num.overrides.get(ilvl)
        if override is not None and override.start_override is not None:
            return override.start_override
        return None


_SPECIFIED = {}


def _specified(level: Level, prop: str) -> bool:
    """Whether an override level carried a non-default property.

    Override levels are parsed from a dedicated element so defaults are
    detected against a fresh :class:`Level` instance.
    """
    defaults = Level(ilvl=level.ilvl)
    return getattr(level, prop) != getattr(defaults, prop)


class Counter:
    """Per-view running counters for all numbering instances."""

    def __init__(self, numbering: Numbering):
        self._numbering = numbering
        self._values: dict[int, dict[int, int]] = {}
        # One-shot start overrides consumed within this numbering pass
        # (accept/reject views each get an independent Counter).
        self._overrides_used: dict[int, set[int]] = {}

    def advance(self, num_id, ilvl):
        """Advance counters for a numbered paragraph; return value tuple.

        Returns a dict ``{level_index: value}`` containing the values of all
        started levels up to and including ``ilvl``.
        """
        level = self._numbering.resolve(num_id, ilvl)
        if level.fmt not in SUPPORTED_FORMATS:
            raise AuditError(
                "unsupported_number_format",
                f"编号格式 w:numFmt=\"{level.fmt}\" 暂不支持",
            )
        instance = self._values.setdefault(num_id, {})

        # Deeper levels are reset only when this level *continues* an active
        # list (increments); a level appearing for the first time (or after
        # being reset itself) merely restarts, it does not restart children.
        continued = ilvl in instance
        if continued:
            for deeper in range(ilvl + 1, MAX_LEVELS):
                deeper_level = self._numbering.resolve(num_id, deeper)
                policy = deeper_level.restart
                # policy 0: any more significant level restarts this one;
                # policy k: only the k nearest significant levels do.
                distance = deeper - ilvl
                if policy == 0 or distance <= policy:
                    instance.pop(deeper, None)
            instance[ilvl] += 1
        else:
            used = self._overrides_used.setdefault(num_id, set())
            if ilvl not in used:
                used.add(ilvl)
                override = self._numbering.peek_start_override(num_id, ilvl)
                instance[ilvl] = override if override is not None else level.start
            else:
                instance[ilvl] = level.start

        values = {lvl: value for lvl, value in instance.items() if lvl <= ilvl}

        # Validate the formats needed to render this level's full number.
        for used_lvl in values:
            used_level = self._numbering.resolve(num_id, used_lvl)
            if used_level.fmt not in SUPPORTED_FORMATS:
                raise AuditError(
                    "unsupported_number_format",
                    f"编号格式 w:numFmt=\"{used_level.fmt}\" 暂不支持",
                )
        return values


# --------------------------------------------------------------------- format
_ROMAN_PAIRS = [
    (1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"),
    (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"),
    (5, "v"), (4, "iv"), (1, "i"),
]


def _to_roman(value: int, upper: bool) -> str:
    out = []
    for number, numeral in _ROMAN_PAIRS:
        while value >= number:
            out.append(numeral)
            value -= number
    text = "".join(out)
    return text.upper() if upper else text


def _to_letters(value: int, upper: bool) -> str:
    # Word lettering: A..Z then AA, AB, ...
    letters = []
    while value > 0:
        value, rem = divmod(value - 1, 26)
        letters.append(chr(ord("a" if not upper else "A") + rem))
    return "".join(reversed(letters))


def _to_ordinal(value: int) -> str:
    if 10 <= value % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(value % 10, "th")
    return f"{value}{suffix}"


def format_value(level: Level, value: int | None) -> str:
    if value is None:
        return ""
    fmt = level.fmt
    if fmt == "decimal":
        return str(value)
    if fmt == "decimalZero":
        width = max(2, len(str(level.start)))
        return str(value).zfill(width)
    if fmt == "lowerRoman":
        return _to_roman(value, False)
    if fmt == "upperRoman":
        return _to_roman(value, True)
    if fmt == "lowerLetter":
        return _to_letters(value, False)
    if fmt == "upperLetter":
        return _to_letters(value, True)
    if fmt == "ordinal":
        return _to_ordinal(value)
    if fmt == "none":
        return ""
    raise AuditError(
        "unsupported_number_format", f"编号格式 w:numFmt=\"{fmt}\" 暂不支持"
    )


# --------------------------------------------------------------- rendering
_PLACEHOLDER = re.compile(r"%(?:%|(\d+))")


def render_full(numbering: Numbering, num_id: int, ilvl: int, values: dict) -> str:
    """Render the complete clause number (``w`` switch behaviour)."""
    level = numbering.resolve(num_id, ilvl)
    return _render(level.text, numbering, num_id, ilvl, values, range(0, ilvl + 1))


def render_own(numbering: Numbering, num_id: int, ilvl: int, values: dict) -> str:
    """Render just the current level token (``n`` switch behaviour)."""
    level = numbering.resolve(num_id, ilvl)
    return format_value(level, values.get(ilvl))


def render_relative(
    numbering: Numbering,
    num_id: int,
    ilvl: int,
    values: dict,
    context_num_id: int | None,
    context_values: dict | None,
) -> str:
    """Render the number tail relative to a context heading (``r`` switch)."""
    level = numbering.resolve(num_id, ilvl)

    common = 0
    if context_num_id == num_id and context_values is not None:
        while (
            common <= ilvl
            and values.get(common) is not None
            and values.get(common) == context_values.get(common)
        ):
            common += 1
    if common > ilvl:
        # Target and context share the exact same number: render the full one.
        common = 0

    return _render(
        level.text,
        numbering,
        num_id,
        ilvl,
        values,
        range(common, ilvl + 1),
    )


def _render(text, numbering, num_id, ilvl, values, selected_levels):
    """Replace ``%n`` placeholders for selected 0-based levels in ``text``.

    Literal text preceding the first selected placeholder is dropped; literal
    text after the last selected placeholder is kept; ``w:isLgl`` forces
    decimal rendering for every level more significant than the paragraph
    level.
    """
    selected = set(selected_levels)
    own_level = numbering.resolve(num_id, ilvl)

    pieces: list[str] = []
    cursor = 0
    started = False
    for match in _PLACEHOLDER.finditer(text):
        token = text[match.start():match.end()]
        literal = text[cursor:match.start()]
        cursor = match.end()
        if token == "%%":
            if started:
                pieces.append("%")
            continue
        idx = int(match.group(1)) - 1
        if idx in selected:
            if not started:
                started = True
                # Drop the separator before the first selected placeholder.
            else:
                pieces.append(literal)
            if own_level.is_lgl and idx < ilvl:
                pieces.append(str(values.get(idx, "")))
            else:
                token_level = numbering.resolve(num_id, idx)
                pieces.append(format_value(token_level, values.get(idx)))
    if started:
        pieces.append(text[cursor:])
    result = "".join(pieces)
    if not started:
        # lvlText does not mention a required placeholder; fall back to a
        # dotted concatenation of the selected values.
        result = ".".join(
            format_value(numbering.resolve(num_id, idx), values.get(idx))
            for idx in sorted(selected)
        )
    return result
