"""OOXML (WordprocessingML) namespace map used throughout the auditor."""

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

NSMAP = {"w": W_NS, "r": R_NS}


def w(tag: str) -> str:
    """Clark-notation name for a ``w:`` element."""
    return f"{{{W_NS}}}{tag}"


def r(tag: str) -> str:
    """Clark-notation name for an ``r:`` element."""
    return f"{{{R_NS}}}{tag}"
