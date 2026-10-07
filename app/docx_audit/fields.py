"""REF field instruction parsing."""

from __future__ import annotations

import re
import dataclasses

from .errors import AuditError

# Switches that may accompany a numbered REF.
_NUMBER_SWITCHES = {"n", "r", "w"}
_HYPERLINK_SWITCH = "h"
_GENERAL_FORMAT_PATTERN = re.compile(r"^MERGEFORMAT$|^@|^#")

_TOKEN = re.compile(r'"[^"]*"|\S+')


@dataclasses.dataclass
class RefInstruction:
    field_type: str  # e.g. "REF" / "HYPERLINK" / "PAGEREF"
    bookmark: str | None = None
    number_switch: str | None = None  # "n" | "r" | "w"
    hyperlink: bool = False


def parse_instruction(instr: str) -> RefInstruction | None:
    """Parse a field instruction.

    Returns ``None`` for non-REF fields.  Raises on malformed REF fields or
    on REF switches this service cannot audit.
    """
    collapsed = " ".join(instr.split())
    tokens = _TOKEN.findall(collapsed)
    tokens = [t[1:-1] if t.startswith('"') and t.endswith('"') else t for t in tokens]
    if not tokens:
        return None

    field_type = tokens[0]
    if field_type.upper() != "REF":
        return RefInstruction(field_type=field_type.upper())

    bookmark = None
    number_switch = None
    hyperlink = False
    expect_format_value = False

    for token in tokens[1:]:
        if expect_format_value:
            expect_format_value = False
            continue
        if token == "\\*":
            expect_format_value = True
            continue
        if token.startswith("\\"):
            switch = token[1:]
            if switch in _NUMBER_SWITCHES:
                if number_switch is not None and number_switch != switch:
                    raise AuditError(
                        "malformed_field",
                        f"REF 域同时指定了多个编号开关 \\{number_switch} 和 \\{switch}",
                    )
                number_switch = switch
            elif switch == _HYPERLINK_SWITCH:
                hyperlink = True
            elif switch in ("p", "f", "d", "t"):
                raise AuditError(
                    "unsupported_ref_switch",
                    f"REF 域包含本服务不支持的开关 \\{switch}",
                )
            else:
                raise AuditError(
                    "unsupported_ref_switch",
                    f"REF 域包含未知开关 \\{switch}",
                )
        elif bookmark is None:
            bookmark = token
        else:
            raise AuditError(
                "malformed_field", f"REF 域存在多余的指令片段: {token}"
            )

    if bookmark is None:
        raise AuditError("malformed_field", "REF 域缺少书签名")
    if number_switch is None:
        # Unnumbered REF (plain text / page / hyperlink) -- not audited.
        return RefInstruction(
            field_type="REF", bookmark=bookmark, hyperlink=hyperlink
        )
    return RefInstruction(
        field_type="REF",
        bookmark=bookmark,
        number_switch=number_switch,
        hyperlink=hyperlink,
    )
