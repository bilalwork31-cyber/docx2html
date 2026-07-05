"""
Core helpers: OOXML namespaces, element utilities, and unit conversions.

Word stores geometry in several units. Getting these conversions exactly
right is the whole ballgame for 1:1 layout parity:

    * twips      = 1/20 point = 1/1440 inch   (margins, indents, table widths)
    * EMU        = 1/914400 inch              (DrawingML: image/shape geometry)
    * half-point = 1/2 point                  (font sizes: <w:sz w:val="24"/> = 12pt)
    * eighth-pt  = 1/8 point                  (border widths)
    * pct50      = 1/50 percent               (some width/shade values)

CSS reference DPI is 96, so 1in = 96px and 1pt = 96/72 px = 4/3 px.
We emit lengths in px (predictable box math) and font metrics in pt
(matches Word's typographic intent). Both resolve to the same physical
scale, so mixing them is safe.
"""
from __future__ import annotations

# ---------------------------------------------------------------------------
# Namespaces
# ---------------------------------------------------------------------------
NS = {
    "w":    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r":    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a":    "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp":   "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "pic":  "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "wps":  "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
    "wpg":  "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpc":  "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "mc":   "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "v":    "urn:schemas-microsoft-com:vml",
    "o":    "urn:schemas-microsoft-com:office:office",
    "w10":  "urn:schemas-microsoft-com:office:word",
    "w14":  "http://schemas.microsoft.com/office/word/2010/wordml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "rel":  "http://schemas.openxmlformats.org/package/2006/relationships",
    "ct":   "http://schemas.openxmlformats.org/package/2006/content-types",
}


def qn(name: str) -> str:
    """'w:val' -> '{http://.../main}val'  (Clark notation for ElementTree)."""
    prefix, local_name = name.split(":", 1)
    return "{%s}%s" % (NS[prefix], local_name)


def local(tag) -> str:
    """Strip the namespace from a Clark-notation tag: '{ns}p' -> 'p'."""
    if tag is None or not isinstance(tag, str):
        return ""
    return tag.rsplit("}", 1)[-1]


# ---------------------------------------------------------------------------
# Element access helpers (all namespace-aware, all None-safe)
# ---------------------------------------------------------------------------
def get(el, name: str, default=None):
    """Attribute lookup by prefixed name, e.g. get(el, 'w:val')."""
    if el is None:
        return default
    v = el.get(qn(name))
    return default if v is None else v


def find(el, path: str):
    """First descendant matching an ElementTree path using our NS map."""
    if el is None:
        return None
    return el.find(path, NS)


def findall(el, path: str):
    if el is None:
        return []
    return el.findall(path, NS)


def child(el, name: str):
    """First *direct* child with the given prefixed tag."""
    if el is None:
        return None
    want = qn(name)
    for c in el:
        if c.tag == want:
            return c
    return None


def children(el, name: str | None = None):
    """Direct children, optionally filtered by prefixed tag."""
    if el is None:
        return []
    if name is None:
        return list(el)
    want = qn(name)
    return [c for c in el if c.tag == want]


def bool_attr(val, default=False):
    """Interpret an OOXML on/off value ('1','0','true','false','on','off')."""
    if val is None:
        return default
    return str(val).lower() not in ("0", "false", "off", "no")


def toggle(el, name: str):
    """
    Toggle property such as <w:b/> or <w:b w:val="0"/>.
    Returns True (on), False (explicitly off), or None (absent).
    """
    if el is None:
        return None
    sub = child(el, name)
    if sub is None:
        return None
    return bool_attr(sub.get(qn("w:val")), default=True)


def int_val(el, name: str, attr: str = "w:val", default=None):
    sub = child(el, name)
    if sub is None:
        return default
    raw = sub.get(qn(attr))
    if raw is None:
        return default
    try:
        return int(round(float(raw)))
    except (TypeError, ValueError):
        return default


def str_val(el, name: str, attr: str = "w:val", default=None):
    sub = child(el, name)
    if sub is None:
        return default
    v = sub.get(qn(attr))
    return default if v is None else v


# ---------------------------------------------------------------------------
# Unit conversions -> CSS
# ---------------------------------------------------------------------------
PX_PER_INCH = 96.0
PT_PER_INCH = 72.0
EMU_PER_INCH = 914400.0
TWIPS_PER_INCH = 1440.0
PT_TO_PX = PX_PER_INCH / PT_PER_INCH  # 4/3


def _num(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return None


def twips_to_px(v):
    v = _num(v)
    return None if v is None else v / TWIPS_PER_INCH * PX_PER_INCH  # v/15


def emu_to_px(v):
    v = _num(v)
    return None if v is None else v / EMU_PER_INCH * PX_PER_INCH  # v/9525


def halfpt_to_pt(v):
    v = _num(v)
    return None if v is None else v / 2.0


def eighthpt_to_px(v):
    """Border widths are in 1/8 pt."""
    v = _num(v)
    return None if v is None else (v / 8.0) * PT_TO_PX


def pt_to_px(v):
    v = _num(v)
    return None if v is None else v * PT_TO_PX


def px(v, nd: int = 2):
    """Format a px number compactly ('12.0' -> '12', '12.500' -> '12.5')."""
    if v is None:
        return None
    v = round(float(v), nd)
    if v == int(v):
        return "%dpx" % int(v)
    return ("%.*f" % (nd, v)).rstrip("0").rstrip(".") + "px"


def pt(v, nd: int = 2):
    if v is None:
        return None
    v = round(float(v), nd)
    if v == int(v):
        return "%dpt" % int(v)
    return ("%.*f" % (nd, v)).rstrip("0").rstrip(".") + "pt"
