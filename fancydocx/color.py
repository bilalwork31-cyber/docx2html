"""
Color resolution: hex parsing, theme-color lookup, and the tint/shade
math Office applies to themed colors.

Word colors come in three flavors:
    * explicit sRGB      <w:color w:val="1F4E79"/>
    * "auto"             <w:color w:val="auto"/>  (context default)
    * theme reference    <w:color w:themeColor="accent1" w:themeShade="BF"/>

For theme references, `themeTint`/`themeShade` are a hex fraction of 255
applied to the *luminance* of the resolved theme color (HSL space) -- this
is what Office actually does, not a naive per-channel scale, so the
accent-bar shades come out matching.
"""
from __future__ import annotations
import colorsys

# Named highlight colors (<w:highlight w:val="yellow"/>).
HIGHLIGHT = {
    "black": "000000", "blue": "0000FF", "cyan": "00FFFF", "darkBlue": "00008B",
    "darkCyan": "008B8B", "darkGray": "A9A9A9", "darkGreen": "006400",
    "darkMagenta": "8B008B", "darkRed": "8B0000", "darkYellow": "808000",
    "green": "00FF00", "lightGray": "D3D3D3", "magenta": "FF00FF", "red": "FF0000",
    "white": "FFFFFF", "yellow": "FFFF00",
}

# themeColor attribute value -> clrScheme key.  The <w:clrSchemeMapping> in
# settings.xml can remap tx1/bg1/tx2/bg2, handled in theme.py; this is the
# default identity mapping.
THEME_ALIAS = {
    "dark1": "dk1", "light1": "lt1", "dark2": "dk2", "light2": "lt2",
    "text1": "dk1", "background1": "lt1", "text2": "dk2", "background2": "lt2",
    "accent1": "accent1", "accent2": "accent2", "accent3": "accent3",
    "accent4": "accent4", "accent5": "accent5", "accent6": "accent6",
    "hyperlink": "hlink", "followedHyperlink": "folHlink",
}


def normalize_hex(val):
    """Return a 6-digit uppercase hex string, or None for auto/blank/invalid."""
    if not val:
        return None
    v = val.strip().lstrip("#")
    if v.lower() == "auto":
        return None
    if len(v) == 3:  # rare shorthand
        v = "".join(c * 2 for c in v)
    if len(v) != 6:
        return None
    try:
        int(v, 16)
    except ValueError:
        return None
    return v.upper()


def hex_to_rgb(h):
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def rgb_to_hex(rgb):
    return "".join("%02X" % max(0, min(255, int(round(c)))) for c in rgb)


def apply_tint_shade(hex6, tint=None, shade=None):
    """
    Apply themeTint / themeShade (hex byte, fraction of 255) to a base color,
    operating on HSL luminance the way Office does.
    """
    if not hex6:
        return hex6
    r, g, b = (c / 255.0 for c in hex_to_rgb(hex6))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    if shade is not None:
        try:
            f = int(shade, 16) / 255.0
            l = l * f
        except ValueError:
            pass
    if tint is not None:
        try:
            f = int(tint, 16) / 255.0
            l = l * f + (1.0 - f)
        except ValueError:
            pass
    l = max(0.0, min(1.0, l))
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return rgb_to_hex((r * 255, g * 255, b * 255))


def color_descriptor(el):
    """
    Build a color descriptor from any element carrying w:val / w:themeColor
    (+ themeTint/themeShade).  Returns None if the element is absent.
    """
    if el is None:
        return None
    from .core import qn
    return {
        "val": el.get(qn("w:val")),
        "theme": el.get(qn("w:themeColor")),
        "tint": el.get(qn("w:themeTint")),
        "shade": el.get(qn("w:themeShade")),
    }


def resolve(desc, theme, default=None):
    """
    Descriptor -> '#RRGGBB' (or `default` when it resolves to auto/none).

    Precedence: when Word saves a theme-referenced color it ALSO bakes the
    resolved sRGB into w:val (e.g. w:color w:val="9A92BF"
    w:themeColor="accent5" w:themeTint="99"). That cached value is Word's own
    integer-HSL math -- bit-exact by definition -- so prefer it and only
    recompute from the theme when no explicit value exists (or it is 'auto').
    """
    if desc is None:
        return default
    hexv = normalize_hex(desc.get("val"))
    if hexv:
        return "#" + hexv
    tname = desc.get("theme")
    if tname and theme is not None:
        base = theme.color(tname) or theme.color(THEME_ALIAS.get(tname, tname))
        if base:
            base = apply_tint_shade(base, desc.get("tint"), desc.get("shade"))
            return "#" + base
    return default
