"""numbering.xml 解析与编号计算。

按 Word 规则处理：
- 列表实例：计数器按 w:num（numId）分别维护，互不影响；
- 级别重启：w:lvlRestart 指定在某一更高级别出现后重启；缺省时任何
  更高级别出现都会重启；w:val="0" 表示永不重启；
- 起始覆盖：w:num/w:lvlOverride/w:startOverride 覆盖级别的起始值，
  w:lvlOverride/w:lvl 整体替换级别定义；
- 数字格式：decimal / decimalZero / upperRoman / lowerRoman /
  upperLetter / lowerLetter / ordinal，其余按 decimal 处理；
  bullet / none 视为无编号。
"""

from __future__ import annotations

from typing import NamedTuple, Optional

from .oxml import W, wval

NO_NUMBER_FMTS = {"bullet", "none"}


class Level(NamedTuple):
    start: int = 1
    fmt: str = "decimal"
    text: str = ""
    restart: Optional[int] = None  # None=任一更高级别后重启；0=永不；k=第 k 级出现后
    pstyle: Optional[str] = None


def _parse_level(lvl_el, ilvl: int) -> Level:
    start_text = wval(lvl_el.find(W + "start"))
    restart_el = lvl_el.find(W + "lvlRestart")
    return Level(
        start=int(start_text) if start_text is not None else 1,
        fmt=wval(lvl_el.find(W + "numFmt")) or "decimal",
        text=wval(lvl_el.find(W + "lvlText")) or f"%{ilvl + 1}.",
        restart=int(wval(restart_el)) if restart_el is not None else None,
        pstyle=wval(lvl_el.find(W + "pStyle")),
    )


class Numbering:
    def __init__(self, root):
        self._abstracts: dict = {}  # abstractNumId -> {ilvl: Level}
        self._nums: dict = {}  # numId -> (abstractNumId, {ilvl: (startOverride, Level|None)})
        self.pstyle_map: dict = {}  # pStyle -> (numId, ilvl)，numId 最小者优先
        if root is None:
            return
        for an in root.findall(W + "abstractNum"):
            aid = an.get(W + "abstractNumId")
            levels = {}
            for lvl_el in an.findall(W + "lvl"):
                ilvl = int(lvl_el.get(W + "ilvl") or 0)
                levels[ilvl] = _parse_level(lvl_el, ilvl)
            self._abstracts[aid] = levels
        for num_el in root.findall(W + "num"):
            nid = num_el.get(W + "numId")
            aid = wval(num_el.find(W + "abstractNumId"))
            overrides: dict = {}
            for lo in num_el.findall(W + "lvlOverride"):
                ilvl = int(lo.get(W + "ilvl") or 0)
                so_text = wval(lo.find(W + "startOverride"))
                start_override = int(so_text) if so_text is not None else None
                lvl_el = lo.find(W + "lvl")
                lvl = _parse_level(lvl_el, ilvl) if lvl_el is not None else None
                overrides[ilvl] = (start_override, lvl)
            self._nums[nid] = (aid, overrides)
        # pStyle 链接：样式未显式编号时的兜底来源
        for nid in sorted(self._nums, key=lambda x: int(x) if x.isdigit() else 1 << 30):
            aid, _ = self._nums[nid]
            for ilvl, level in self._abstracts.get(aid, {}).items():
                if level.pstyle and level.pstyle not in self.pstyle_map:
                    self.pstyle_map[level.pstyle] = (nid, ilvl)

    def has_num(self, num_id: str) -> bool:
        return num_id in self._nums

    def level(self, num_id: str, ilvl: int) -> Level:
        aid, overrides = self._nums[num_id]
        base = self._abstracts.get(aid, {}).get(ilvl)
        if base is None:
            base = Level(text=f"%{ilvl + 1}.")
        ov = overrides.get(ilvl)
        if ov is not None:
            start_override, lvl = ov
            if lvl is not None:
                base = lvl
            if start_override is not None:
                base = base._replace(start=start_override)
        return base

    def start(self, num_id: str, ilvl: int) -> int:
        return self.level(num_id, ilvl).start

    def fmt(self, num_id: str, ilvl: int) -> str:
        return self.level(num_id, ilvl).fmt

    def restart(self, num_id: str, ilvl: int):
        return self.level(num_id, ilvl).restart


# ---------------------------------------------------------------- 数字格式化


def _to_roman(n: int) -> str:
    if n <= 0:
        return str(n)
    out = []
    for value, symbol in (
        (1000, "M"),
        (900, "CM"),
        (500, "D"),
        (400, "CD"),
        (100, "C"),
        (90, "XC"),
        (50, "L"),
        (40, "XL"),
        (10, "X"),
        (9, "IX"),
        (5, "V"),
        (4, "IV"),
        (1, "I"),
    ):
        while n >= value:
            out.append(symbol)
            n -= value
    return "".join(out)


def _to_letter(n: int) -> str:
    if n <= 0:
        return str(n)
    out = []
    while n > 0:
        n, r = divmod(n - 1, 26)
        out.append(chr(65 + r))
    return "".join(reversed(out))


def _to_ordinal(n: int) -> str:
    if 10 <= n % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def format_number(n: int, fmt: str) -> str:
    if fmt == "decimal":
        return str(n)
    if fmt == "decimalZero":
        return f"{n:02d}"
    if fmt == "upperRoman":
        return _to_roman(n)
    if fmt == "lowerRoman":
        return _to_roman(n).lower()
    if fmt == "upperLetter":
        return _to_letter(n)
    if fmt == "lowerLetter":
        return _to_letter(n).lower()
    if fmt == "ordinal":
        return _to_ordinal(n)
    if fmt in NO_NUMBER_FMTS:
        return ""
    return str(n)


# ---------------------------------------------------------------- 计数引擎


class NumberingEngine:
    """按文档顺序为每个编号段落推进计数器（计数器按 numId 隔离）。"""

    def __init__(self, numbering: Numbering):
        self._numbering = numbering
        self._counters: dict = {}  # numId -> [c0..c8]，None 表示未启用

    def advance(self, num_id: str, ilvl: int):
        """推进指定列表实例的级别计数，返回 (原始值列表, 格式化列表)（0..ilvl 级）。"""
        counters = self._counters.setdefault(num_id, [None] * 9)
        start = self._numbering.start(num_id, ilvl)
        counters[ilvl] = start if counters[ilvl] is None else counters[ilvl] + 1
        # 更深层级按重启规则复位
        for j in range(ilvl + 1, 9):
            rv = self._numbering.restart(num_id, j)
            if rv is None or (rv != 0 and ilvl + 1 <= rv):
                counters[j] = None
        raw = []
        fmt = []
        for j in range(ilvl + 1):
            value = counters[j]
            if value is None:
                value = self._numbering.start(num_id, j)
            raw.append(value)
            fmt.append(format_number(value, self._numbering.fmt(num_id, j)))
        return raw, fmt
