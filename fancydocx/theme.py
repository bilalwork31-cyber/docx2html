"""
Theme parsing: the color scheme (dk1/lt1/dk2/lt2/accent1..6/hlink) and the
font scheme (major/minor). Runs and styles reference these indirectly
(w:themeColor="accent1", w:rFonts w:asciiTheme="minorHAnsi"), so we resolve
them up front.

The document's <w:clrSchemeMapping> (settings.xml) may remap the
text1/background1/text2/background2 slots onto scheme entries; we honor it.
"""
from __future__ import annotations
from .core import NS, qn, local, find

# Default themeColor-slot -> clrScheme-entry mapping.
DEFAULT_ALIAS = {
    "dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2",
    "text1": "dk1", "background1": "lt1", "text2": "dk2", "background2": "lt2",
    "accent1": "accent1", "accent2": "accent2", "accent3": "accent3",
    "accent4": "accent4", "accent5": "accent5", "accent6": "accent6",
    "hyperlink": "hlink", "followedHyperlink": "folHlink",
}
# <w:clrSchemeMapping w:tx1="dark1" .../> attribute -> slot it controls.
MAPPING_ATTRS = {
    "bg1": "background1", "t1": "text1", "bg2": "background2", "t2": "text2",
    "accent1": "accent1", "accent2": "accent2", "accent3": "accent3",
    "accent4": "accent4", "accent5": "accent5", "accent6": "accent6",
    "hlink": "hyperlink", "folHlink": "followedHyperlink",
}


class Theme:
    def __init__(self, pkg):
        self.scheme = {}   # 'dk1' -> 'RRGGBB'
        self.alias = dict(DEFAULT_ALIAS)
        self.fonts = {"major": None, "minor": None}
        self._load(pkg)
        self._load_color_map(pkg)

    def _theme_part(self, pkg):
        for rel in pkg.doc_rels.values():
            if rel.type.endswith("theme"):
                return pkg.resolve_target(rel.target)
        return "word/theme/theme1.xml"

    def _load(self, pkg):
        root = pkg.xml(self._theme_part(pkg))
        if root is None:
            return
        elements = find(root, "a:themeElements")
        clr = find(elements, "a:clrScheme")
        if clr is not None:
            for entry in clr:
                key = local(entry.tag)          # dk1, lt1, accent1, ...
                self.scheme[key] = _color_of(entry)
        fs = find(elements, "a:fontScheme")
        if fs is not None:
            major = find(fs, "a:majorFont")
            minor = find(fs, "a:minorFont")
            self.fonts["major"] = _latin_of(major)
            self.fonts["minor"] = _latin_of(minor)

    def _load_color_map(self, pkg):
        settings = pkg.xml("word/settings.xml")
        if settings is None:
            return
        mapping = find(settings, "w:clrSchemeMapping")
        if mapping is None:
            return
        for attr, slot in MAPPING_ATTRS.items():
            val = mapping.get(qn("w:" + attr))
            if not val:
                continue
            # val is like 'dark1'/'light2'/'accent3' -> resolve to a scheme key.
            self.alias[slot] = DEFAULT_ALIAS.get(val, val)

    # -- public lookups ----------------------------------------------------
    def color(self, name):
        """Accept a scheme key ('accent1') or a slot alias ('text1')."""
        if name in self.scheme:
            return self.scheme[name]
        key = self.alias.get(name)
        if key and key in self.scheme:
            return self.scheme[key]
        return None

    def font(self, which):
        """which in {'major','minor'} (or 'majorHAnsi'/'minorAscii'/...)."""
        if which.startswith("major"):
            return self.fonts.get("major")
        if which.startswith("minor"):
            return self.fonts.get("minor")
        return None


def _color_of(entry):
    """<a:dk1><a:sysClr lastClr='000000'/></a:dk1> or <a:srgbClr val='...'/>."""
    srgb = find(entry, "a:srgbClr")
    if srgb is not None:
        v = srgb.get("val")
        if v:
            return v.upper()
    sysclr = find(entry, "a:sysClr")
    if sysclr is not None:
        v = sysclr.get("lastClr") or sysclr.get("val")
        if v and len(v) == 6:
            return v.upper()
    return None


def _latin_of(fontnode):
    latin = find(fontnode, "a:latin")
    if latin is not None:
        return latin.get("typeface")
    return None
