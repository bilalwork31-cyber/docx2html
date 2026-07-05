"""
Property model + style cascade + CSS generation.

Word resolves formatting through a cascade:

    docDefaults  ->  paragraph-style chain (basedOn)  ->  direct pPr/rPr

Character formatting adds a character-style chain and the run's own rPr on
top. We parse each rPr/pPr XML fragment into a normalized dict, deep-merge
the dicts in cascade order, and only then emit CSS. Keeping properties as
data until the end is what makes inheritance correct instead of guesswork.
"""
from __future__ import annotations

from .core import (
    NS, qn, child, children, toggle, get, find,
    twips_to_px, halfpt_to_pt, eighthpt_to_px, pt_to_px, px, pt, PT_TO_PX,
)
from .color import color_descriptor, resolve, HIGHLIGHT
from .fontmetrics import line_factor, substitute

# Serif families (for a sensible CSS generic fallback).
_SERIF_HINTS = (
    "times", "serif", "georgia", "cambria", "garamond", "minion", "book antiqua",
    "playfair", "merriweather", "goudy", "baskerville", "palatino", "constantia",
    "rockwell", "bookman", "caslon", "didot", "sylfaen", "cochin",
)
_MONO_HINTS = ("mono", "consolas", "courier", "menlo")

BORDER_STYLE = {
    "single": "solid", "thick": "solid", "double": "double", "triple": "double",
    "dotted": "dotted", "dashed": "dashed", "dotDash": "dashed",
    "dotDotDash": "dashed", "dashDotStroked": "dashed", "dashSmallGap": "dashed",
    "wave": "solid", "doubleWave": "double", "inset": "inset", "outset": "outset",
    "threeDEngrave": "groove", "threeDEmboss": "ridge",
    "thinThickSmallGap": "double", "thickThinSmallGap": "double",
    "thinThickThinSmallGap": "double", "none": None, "nil": None,
}
UNDERLINE_STYLE = {
    "single": "solid", "double": "double", "thick": "solid", "dotted": "dotted",
    "dottedHeavy": "dotted", "dash": "dashed", "dashedHeavy": "dashed",
    "dashLong": "dashed", "dotDash": "dashed", "dotDotDash": "dashed",
    "wave": "wavy", "wavyHeavy": "wavy", "wavyDouble": "wavy",
}
ALIGN = {
    "left": "left", "start": "left", "right": "right", "end": "right",
    "center": "center", "both": "justify", "distribute": "justify",
}


def quote_font(name):
    if not name:
        return None
    low = name.strip().lower()
    if any(h in low for h in _MONO_HINTS):
        generic = "monospace"
    elif any(h in low for h in _SERIF_HINTS):
        generic = "serif"
    else:
        generic = "sans-serif"
    # Single-quote the family: these strings land inside a double-quoted
    # style="..." attribute, so double quotes there would truncate the attribute.
    parts = ["'%s'" % name.replace("'", "")]
    # Word-style substitution when the font is missing (Aptos->Calibri,
    # Futura->Century Gothic, ...) so the fallback matches what Word shows
    # on a machine without the font, not the browser's default sans.
    sub = substitute(name)
    if sub and sub.lower() != low:
        parts.append("'%s'" % sub.replace("'", ""))
    parts.append(generic)
    return ", ".join(parts)


def _attr_color(el):
    """Descriptor from an element's color *attributes* (u/bdr carry these)."""
    if el is None:
        return None
    return {
        "val": el.get(qn("w:color")),
        "theme": el.get(qn("w:themeColor")),
        "tint": el.get(qn("w:themeTint")),
        "shade": el.get(qn("w:themeShade")),
    }


# ---------------------------------------------------------------------------
# Shading + borders
# ---------------------------------------------------------------------------
def parse_shd(shd):
    if shd is None:
        return None
    return {
        "pattern": shd.get(qn("w:val")),
        "fill": {
            "val": shd.get(qn("w:fill")), "theme": shd.get(qn("w:themeFill")),
            "tint": shd.get(qn("w:themeFillTint")), "shade": shd.get(qn("w:themeFillShade")),
        },
        "fg": {
            "val": shd.get(qn("w:color")), "theme": shd.get(qn("w:themeColor")),
            "tint": shd.get(qn("w:themeTint")), "shade": shd.get(qn("w:themeShade")),
        },
    }


def resolve_shd(shd, theme):
    if not shd:
        return None
    pat = shd.get("pattern")
    fillhex = resolve(shd.get("fill"), theme)
    fghex = resolve(shd.get("fg"), theme)
    if pat == "solid":
        return fghex or fillhex
    if pat in (None, "clear", "nil"):
        return fillhex
    return fillhex or fghex  # percentage patterns -> approximate with fill


def parse_border_side(el):
    if el is None:
        return None
    val = el.get(qn("w:val"))
    if val in ("nil", "none", None):
        return {"val": "none"}
    return {
        "val": val,
        "sz": el.get(qn("w:sz")),          # eighths of a point
        "space": el.get(qn("w:space")),    # points
        "color": _attr_color(el),
    }


def parse_borders(bd):
    """Parse a *Borders container (<w:pBdr>/<w:tblBorders>/<w:tcBorders>)."""
    if bd is None:
        return None
    out = {}
    for side in ("top", "left", "bottom", "right", "insideH", "insideV", "between", "bar", "start", "end"):
        s = parse_border_side(child(bd, "w:" + side))
        if s is not None:
            out[side] = s
    return out or None


def border_css_value(side, theme):
    """A single border side descriptor -> ('1px solid #000', space_px)."""
    if side is None or side.get("val") == "none":
        return None, 0
    style = BORDER_STYLE.get(side.get("val"), "solid")
    if style is None:
        return None, 0
    w = eighthpt_to_px(side.get("sz"))
    if not w or w < 1:
        w = 1.0
    if side.get("val") == "thick" and w < 2:
        w = 2.0
    color = resolve(side.get("color"), theme, default="#000000")
    space = pt_to_px(side.get("space")) or 0
    return "%s %s %s" % (px(w), style, color), space


# ---------------------------------------------------------------------------
# Run properties
# ---------------------------------------------------------------------------
def _pick_font(rf, theme):
    if rf is None:
        return None
    name = rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi"))
    if not name:
        t = rf.get(qn("w:asciiTheme")) or rf.get(qn("w:hAnsiTheme"))
        if t and theme:
            name = theme.font(t)
    # Deliberately NOT falling back to w:cs / w:eastAsia: those are the
    # complex-script / East-Asian fonts and must not override the Latin font.
    # (Heading/Title styles often carry only cs="Times New Roman (Headings CS)"
    # while their Latin text should inherit the ascii/theme font, e.g. Aptos.)
    return name


def parse_rpr(rpr, theme):
    d = {}
    if rpr is None:
        return d
    rstyle = child(rpr, "w:rStyle")
    if rstyle is not None:
        d["rStyle"] = rstyle.get(qn("w:val"))
    font = _pick_font(child(rpr, "w:rFonts"), theme)
    if font:
        d["font"] = font
    sz = child(rpr, "w:sz")
    if sz is not None:
        d["sz"] = halfpt_to_pt(sz.get(qn("w:val")))
    col = child(rpr, "w:color")
    if col is not None:
        d["color"] = color_descriptor(col)
    for key, tag in (("bold", "w:b"), ("italic", "w:i"), ("strike", "w:strike"),
                     ("dstrike", "w:dstrike"), ("caps", "w:caps"),
                     ("smallCaps", "w:smallCaps"), ("vanish", "w:vanish"),
                     ("outline", "w:outline"), ("emboss", "w:emboss")):
        v = toggle(rpr, tag)
        if v is not None:
            d[key] = v
    if toggle(rpr, "w:webHidden"):
        d["vanish"] = True
    u = child(rpr, "w:u")
    if u is not None:
        uval = u.get(qn("w:val"))
        if uval and uval != "none":
            d["underline"] = {"val": uval, "color": _attr_color(u)}
        elif uval == "none":
            d["underline"] = None
    va = child(rpr, "w:vertAlign")
    if va is not None:
        d["vertAlign"] = va.get(qn("w:val"))
    hl = child(rpr, "w:highlight")
    if hl is not None:
        d["highlight"] = hl.get(qn("w:val"))
    shd = parse_shd(child(rpr, "w:shd"))
    if shd:
        d["shd"] = shd
    sp = child(rpr, "w:spacing")
    if sp is not None and sp.get(qn("w:val")) is not None:
        d["spacing_pt"] = float(sp.get(qn("w:val"))) / 20.0  # twips -> pt
    pos = child(rpr, "w:position")
    if pos is not None and pos.get(qn("w:val")) is not None:
        d["position_pt"] = float(pos.get(qn("w:val"))) / 2.0  # half-pt -> pt
    scale = child(rpr, "w:w")
    if scale is not None and scale.get(qn("w:val")) is not None:
        try:
            d["scale"] = float(scale.get(qn("w:val")))
        except ValueError:
            pass
    return d


def rpr_to_css(p, theme):
    d = {}
    if p.get("vanish"):
        return {"display": "none"}
    if "font" in p:
        d["font-family"] = quote_font(p["font"])
    base_sz = p.get("sz")
    va = p.get("vertAlign")
    if va in ("superscript", "subscript"):
        d["vertical-align"] = "super" if va == "superscript" else "sub"
        d["font-size"] = pt(base_sz * 0.66) if base_sz else "0.66em"
    elif base_sz:
        d["font-size"] = pt(base_sz)
    c = resolve(p.get("color"), theme)
    if c:
        d["color"] = c
    b = p.get("bold")
    if b is True:
        d["font-weight"] = "700"
    elif b is False:
        d["font-weight"] = "400"
    it = p.get("italic")
    if it is True:
        d["font-style"] = "italic"
    elif it is False:
        d["font-style"] = "normal"
    lines = []
    u = p.get("underline")
    if u:
        lines.append("underline")
        d["text-decoration-style"] = UNDERLINE_STYLE.get(u.get("val"), "solid")
        uc = resolve(u.get("color"), theme)
        if uc:
            d["text-decoration-color"] = uc
    if p.get("strike") or p.get("dstrike"):
        lines.append("line-through")
    if lines:
        d["text-decoration-line"] = " ".join(lines)
    if p.get("caps"):
        d["text-transform"] = "uppercase"
    if p.get("smallCaps"):
        d["font-variant"] = "small-caps"
    hl = p.get("highlight")
    if hl and hl != "none" and hl in HIGHLIGHT:
        d["background-color"] = "#" + HIGHLIGHT[hl]
    sh = resolve_shd(p.get("shd"), theme)
    if sh:
        d["background-color"] = sh
    if p.get("spacing_pt"):
        d["letter-spacing"] = pt(p["spacing_pt"])
    if p.get("position_pt"):
        d["position"] = "relative"
        d["top"] = pt(-p["position_pt"])
    if p.get("scale"):
        d["display"] = "inline-block"
        d["transform"] = "scaleX(%g)" % (p["scale"] / 100.0)
    return d


# ---------------------------------------------------------------------------
# Paragraph properties
# ---------------------------------------------------------------------------
def parse_ppr(ppr, theme):
    d = {}
    if ppr is None:
        return d
    ps = child(ppr, "w:pStyle")
    if ps is not None:
        d["pStyle"] = ps.get(qn("w:val"))
    jc = child(ppr, "w:jc")
    if jc is not None:
        d["jc"] = jc.get(qn("w:val"))
    sp = child(ppr, "w:spacing")
    if sp is not None:
        s = {}
        for k, a in (("before", "w:before"), ("after", "w:after"),
                     ("line", "w:line")):
            v = sp.get(qn(a))
            if v is not None:
                s[k] = float(v)
        lr = sp.get(qn("w:lineRule"))
        if lr:
            s["lineRule"] = lr
        if sp.get(qn("w:beforeAutospacing")) in ("1", "true"):
            s["beforeAuto"] = True
        if sp.get(qn("w:afterAutospacing")) in ("1", "true"):
            s["afterAuto"] = True
        if s:
            d["spacing"] = s
    ind = child(ppr, "w:ind")
    if ind is not None:
        i = {}
        for k, a in (("left", "w:left"), ("left", "w:start"), ("right", "w:right"),
                     ("right", "w:end"), ("firstLine", "w:firstLine"),
                     ("hanging", "w:hanging")):
            v = ind.get(qn(a))
            if v is not None:
                i[k] = float(v)
        if i:
            d["ind"] = i
    shd = parse_shd(child(ppr, "w:shd"))
    if shd:
        d["shd"] = shd
    bd = parse_borders(child(ppr, "w:pBdr"))
    if bd:
        d["borders"] = bd
    numpr = child(ppr, "w:numPr")
    if numpr is not None:
        ilvl = child(numpr, "w:ilvl")
        numid = child(numpr, "w:numId")
        d["numPr"] = {
            "ilvl": int(ilvl.get(qn("w:val"))) if ilvl is not None and ilvl.get(qn("w:val")) else 0,
            "numId": numid.get(qn("w:val")) if numid is not None else None,
        }
    if toggle(ppr, "w:contextualSpacing"):
        d["contextualSpacing"] = True
    tabs = child(ppr, "w:tabs")
    if tabs is not None:
        tl = []
        for t in children(tabs, "w:tab"):
            tl.append({
                "val": t.get(qn("w:val")),
                "pos": float(t.get(qn("w:pos"))) if t.get(qn("w:pos")) else 0,
                "leader": t.get(qn("w:leader")),
            })
        if tl:
            d["tabs"] = tl
    # Paragraph-mark run properties (affect the whole line's default look).
    mark = child(ppr, "w:rPr")
    if mark is not None:
        d["markRpr"] = parse_rpr(mark, theme)
    return d


def line_height_css(sp, font_family=None, font_size_pt=None):
    """
    Word line spacing -> CSS line-height.

    * lineRule auto: w:line is 240ths of a *single line*, and a single line is
      the font's metric height (winAscent+winDescent [+ext leading]), NOT 1em.
      We emit the numeric multiplier m*factor(font) so geometry stays at
      Word's scale even if the viewer substitutes the font.
    * exact: fixed px box.
    * atLeast: max(px, natural). CSS can't express that directly; pick the
      larger of the two when we know the font size, else trust the px value.
    * absent: single spacing = factor(font).
    """
    factor = line_factor(font_family)
    sp = sp or {}
    if "line" not in sp:
        return "%g" % factor
    rule = sp.get("lineRule", "auto")
    if rule == "auto":
        return "%g" % round((sp["line"] / 240.0) * factor, 3)
    line_px = twips_to_px(sp["line"])
    if rule == "atLeast" and font_size_pt:
        natural = font_size_pt * PT_TO_PX * factor
        if natural > line_px:
            return "%g" % factor
    return px(line_px)


def ppr_to_css(p, theme, font_family=None, font_size_pt=None):
    d = {}
    jc = p.get("jc")
    if jc in ALIGN:
        d["text-align"] = ALIGN[jc]
    sp = p.get("spacing") or {}
    sh = resolve_shd(p.get("shd"), theme)
    borders = p.get("borders") or {}
    has_box = bool(sh) or any(borders.get(s) for s in ("top", "bottom", "left", "right"))
    # Word does NOT collapse paragraph spacing: gap = after(prev) + before(next).
    # CSS margins collapse (max), so emit `before` as padding-top -- padding
    # never collapses with the previous margin-bottom, giving Word's sum.
    # Exception: shaded/bordered paragraphs, where padding would wrongly
    # extend the decoration into the gap -> keep margin-top there.
    if "before" in sp and not sp.get("beforeAuto"):
        v = px(twips_to_px(sp["before"]))
        if has_box:
            d["margin-top"] = v
        else:
            d["padding-top"] = v
    if "after" in sp and not sp.get("afterAuto"):
        d["margin-bottom"] = px(twips_to_px(sp["after"]))
    if "line" in sp:
        d["line-height"] = line_height_css(sp, font_family, font_size_pt)
    ind = p.get("ind") or {}
    if "left" in ind:
        d["margin-left"] = px(twips_to_px(ind["left"]))
    if "right" in ind:
        d["margin-right"] = px(twips_to_px(ind["right"]))
    if "hanging" in ind:
        d["text-indent"] = px(-twips_to_px(ind["hanging"]))
    elif "firstLine" in ind:
        d["text-indent"] = px(twips_to_px(ind["firstLine"]))
    if sh:
        d["background-color"] = sh
    for side in ("top", "bottom", "left", "right"):
        val, space = border_css_value(borders.get(side), theme)
        if val:
            d["border-%s" % side] = val
            if space:
                d["padding-%s" % side] = px(space)
    return d


# ---------------------------------------------------------------------------
# Deep merge of property dicts (cascade)
# ---------------------------------------------------------------------------
_DEEP_KEYS = ("spacing", "ind", "borders")


def merge_props(base, add):
    """Return base updated by add, deep-merging the nested dict keys."""
    if not add:
        return base
    for k, v in add.items():
        if k in _DEEP_KEYS and isinstance(v, dict) and isinstance(base.get(k), dict):
            merged = dict(base[k])
            merged.update(v)
            base[k] = merged
        else:
            base[k] = v
    return base


def merge_chain(dicts):
    out = {}
    for d in dicts:
        merge_props(out, d)
    return out


# ---------------------------------------------------------------------------
# Styles registry
# ---------------------------------------------------------------------------
class _Style:
    __slots__ = ("id", "type", "based_on", "ppr", "rpr", "name", "default", "el")

    def __init__(self, sid, stype):
        self.id = sid
        self.type = stype
        self.based_on = None
        self.ppr = {}
        self.rpr = {}
        self.name = None
        self.default = False
        self.el = None


class Styles:
    def __init__(self, pkg, theme):
        self.theme = theme
        self.styles = {}          # id -> _Style
        self.default_para = None
        self.default_char = None
        self.default_table = None
        self.doc_ppr = {}
        self.doc_rpr = {}
        self._ppr_cache = {}
        self._rpr_cache = {}
        self._load(pkg)

    def _load(self, pkg):
        root = pkg.xml("word/styles.xml")
        if root is None:
            return
        dd = find(root, "w:docDefaults")
        if dd is not None:
            rprd = find(dd, "w:rPrDefault")
            pprd = find(dd, "w:pPrDefault")
            if rprd is not None:
                self.doc_rpr = parse_rpr(child(rprd, "w:rPr"), self.theme)
            if pprd is not None:
                self.doc_ppr = parse_ppr(child(pprd, "w:pPr"), self.theme)
        for st in children(root, "w:style"):
            sid = st.get(qn("w:styleId"))
            if not sid:
                continue
            stype = st.get(qn("w:type"), "paragraph")
            s = _Style(sid, stype)
            s.el = st
            bo = child(st, "w:basedOn")
            s.based_on = bo.get(qn("w:val")) if bo is not None else None
            nm = child(st, "w:name")
            s.name = nm.get(qn("w:val")) if nm is not None else sid
            s.default = st.get(qn("w:default")) in ("1", "true")
            s.ppr = parse_ppr(child(st, "w:pPr"), self.theme)
            s.rpr = parse_rpr(child(st, "w:rPr"), self.theme)
            self.styles[sid] = s
            if s.default:
                if stype == "paragraph" and self.default_para is None:
                    self.default_para = sid
                elif stype == "character" and self.default_char is None:
                    self.default_char = sid
                elif stype == "table" and self.default_table is None:
                    self.default_table = sid

    def _chain(self, sid):
        """Style ids from root ancestor down to `sid` (basedOn order)."""
        seen = set()
        order = []
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            order.append(sid)
            sid = self.styles[sid].based_on
        return list(reversed(order))

    def style_ppr(self, sid):
        if sid not in self._ppr_cache:
            self._ppr_cache[sid] = merge_chain(
                [self.styles[i].ppr for i in self._chain(sid)])
        return self._ppr_cache[sid]

    def style_rpr(self, sid):
        if sid not in self._rpr_cache:
            self._rpr_cache[sid] = merge_chain(
                [self.styles[i].rpr for i in self._chain(sid)])
        return self._rpr_cache[sid]

    # -- effective (fully cascaded) properties -----------------------------
    def effective_ppr(self, p_style, direct):
        base = dict(self.doc_ppr)
        sid = p_style or self.default_para
        if sid:
            merge_props(base, self.style_ppr(sid))
        merge_props(base, direct or {})
        return base

    def effective_rpr(self, p_style, r_style, direct):
        base = dict(self.doc_rpr)
        psid = p_style or self.default_para
        if psid:
            merge_props(base, self.style_rpr(psid))
        if r_style:
            merge_props(base, self.style_rpr(r_style))
        merge_props(base, direct or {})
        return base
