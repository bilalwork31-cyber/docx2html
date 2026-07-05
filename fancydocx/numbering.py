"""
List numbering. Resolves <w:numPr> (numId + ilvl) to a concrete marker:
the bullet glyph or the formatted ordinal, plus the level's own indent and
run properties (bullet font/size/color).

We render markers ourselves rather than emitting <ul>/<ol> so the glyph,
color, font and hanging indent match Word exactly.
"""
from __future__ import annotations

from .core import NS, qn, child, children, find, str_val, int_val
from .styles import parse_ppr, parse_rpr

# Common Symbol/Wingdings private-use bullet code points -> Unicode.
_BULLET_MAP = {
    "": "•",  # Symbol bullet        -> •
    "·": "•",  # middle dot           -> •
    "": "▪",  # Wingdings sq bullet   -> ▪
    "": "▪",
    "": "◦",  # -> ◦
    "o": "◦",       # courier 'o' sub-bullet -> ◦
    "": "➢",  # -> ➢
    "": "✔",  # Wingdings check       -> ✔
    "": "▪",
    "–": "–",  # en dash bullet
    "-": "-",
    "": "❖",
    "": "⇨",
}


def _to_roman(n, upper=False):
    if n <= 0:
        return str(n)
    vals = [(1000, "m"), (900, "cm"), (500, "d"), (400, "cd"), (100, "c"),
            (90, "xc"), (50, "l"), (40, "xl"), (10, "x"), (9, "ix"),
            (5, "v"), (4, "iv"), (1, "i")]
    out = []
    for v, s in vals:
        while n >= v:
            out.append(s)
            n -= v
    r = "".join(out)
    return r.upper() if upper else r


def _to_letter(n, upper=False):
    # 1->a, 26->z, 27->aa (spreadsheet style)
    if n <= 0:
        return str(n)
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(ord("a") + rem) + s
    return s.upper() if upper else s


def format_number(n, fmt):
    if fmt == "decimal":
        return str(n)
    if fmt == "decimalZero":
        return "%02d" % n
    if fmt == "lowerLetter":
        return _to_letter(n, upper=False)
    if fmt == "upperLetter":
        return _to_letter(n, upper=True)
    if fmt == "lowerRoman":
        return _to_roman(n, upper=False)
    if fmt == "upperRoman":
        return _to_roman(n, upper=True)
    if fmt in ("none",):
        return ""
    return str(n)


def bullet_glyph(lvl_text):
    if not lvl_text:
        return "•"
    ch = lvl_text[0]
    if ch in _BULLET_MAP:
        return _BULLET_MAP[ch]
    if ch.isprintable() and ord(ch) < 0xF000:
        return ch
    return "•"


class Numbering:
    def __init__(self, pkg, theme):
        self.theme = theme
        self.abstract = {}   # abstractNumId -> {ilvl: leveldict}
        self.nums = {}       # numId -> {'abstract':id, 'overrides':{ilvl:start}}
        self._load(pkg)

    def _load(self, pkg):
        root = pkg.xml("word/numbering.xml")
        if root is None:
            return
        for ab in children(root, "w:abstractNum"):
            aid = ab.get(qn("w:abstractNumId"))
            levels = {}
            for lvl in children(ab, "w:lvl"):
                try:
                    il = int(lvl.get(qn("w:ilvl")))
                except (TypeError, ValueError):
                    continue
                levels[il] = {
                    "numFmt": str_val(lvl, "w:numFmt", default="decimal"),
                    "lvlText": str_val(lvl, "w:lvlText", default=""),
                    "start": int_val(lvl, "w:start", default=1),
                    "suff": str_val(lvl, "w:suff", default="tab"),
                    "ppr": parse_ppr(child(lvl, "w:pPr"), self.theme),
                    "rpr": parse_rpr(child(lvl, "w:rPr"), self.theme),
                }
            self.abstract[aid] = levels
        for num in children(root, "w:num"):
            nid = num.get(qn("w:numId"))
            ab = child(num, "w:abstractNumId")
            aid = ab.get(qn("w:val")) if ab is not None else None
            overrides = {}
            for ovr in children(num, "w:lvlOverride"):
                il = ovr.get(qn("w:ilvl"))
                so = child(ovr, "w:startOverride")
                if il is not None and so is not None:
                    try:
                        overrides[int(il)] = int(so.get(qn("w:val")))
                    except (TypeError, ValueError):
                        pass
            self.nums[nid] = {"abstract": aid, "overrides": overrides}

    def level(self, num_id, ilvl):
        info = self.nums.get(num_id)
        if not info:
            return None
        levels = self.abstract.get(info["abstract"])
        if not levels:
            return None
        return levels.get(ilvl)

    def start_value(self, num_id, ilvl):
        info = self.nums.get(num_id)
        if info and ilvl in info["overrides"]:
            return info["overrides"][ilvl]
        lvl = self.level(num_id, ilvl)
        return lvl["start"] if lvl else 1

    def marker(self, num_id, ilvl, count_of):
        """
        Build marker text for the level. `count_of(ilvl)` returns the current
        1-based counter for a level (used to expand %1..%9 in lvlText).
        Returns (text, rpr_dict) or None.
        """
        lvl = self.level(num_id, ilvl)
        if lvl is None:
            return None
        fmt = lvl["numFmt"]
        if fmt == "bullet":
            return bullet_glyph(lvl["lvlText"]), lvl["rpr"]
        text = lvl["lvlText"] or ""
        # Expand %1..%9 placeholders using each referenced level's own format.
        for n in range(1, 10):
            token = "%%%d" % n
            if token in text:
                ref = n - 1
                ref_lvl = self.level(num_id, ref)
                ref_fmt = ref_lvl["numFmt"] if ref_lvl else "decimal"
                text = text.replace(token, format_number(count_of(ref), ref_fmt))
        return text, lvl["rpr"]
