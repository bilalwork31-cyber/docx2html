"""
Local font metrics + (opt-in) font embedding. Pure standard library.

Why this exists: Word's "single" line spacing is not 1.0x the font size and
not a universal 1.15 -- it is the font's own design line height:

    single = (usWinAscent + usWinDescent) / unitsPerEm            [OS/2, head]
           + max(0, lineGap - ((winAsc+winDesc) - (hheaAsc - hheaDesc)))
                                                                   [external leading]

(the GDI TEXTMETRIC formula Word inherits). Calibri = 1.2207, Segoe UI =
1.3301, Times New Roman = 1.1074... If we emitted CSS `line-height:1.0` for
single spacing we would be ~20-30% too tight; `normal` would use whatever
FALLBACK font the viewer has, drifting from Word's geometry whenever the real
font is missing. So we probe the actual font file when it is installed
locally, fall back to a table of known Office/Windows fonts, then to 1.2.
Emitting the *number* keeps the layout at Word's geometry even when the
browser substitutes the family.

The same TTF/TTC name-table parser powers the opt-in local font embedding
pass (--embed-fonts): match referenced families against installed fonts and
inline them as @font-face so text renders with the exact intended metrics.
"""
from __future__ import annotations
import base64
import os
import struct

# Measured (usWinAscent+usWinDescent+extLeading)/upm for fonts we cannot
# always probe. Sources: the fonts' own OS/2 tables.
KNOWN_FACTORS = {
    "calibri": 1.2207, "calibri light": 1.2207,
    "cambria": 1.1729,
    "segoe ui": 1.3301, "segoe ui semibold": 1.3301, "segoe ui light": 1.3301,
    "segoe ui semilight": 1.3301, "segoe ui black": 1.3301,
    "arial": 1.1499, "arial black": 1.4102, "arial narrow": 1.1367,
    "times new roman": 1.1499,
    "georgia": 1.1367,
    "verdana": 1.2158,
    "tahoma": 1.2070,
    "trebuchet ms": 1.1602,
    "courier new": 1.1328,
    "garamond": 1.1250,
    "book antiqua": 1.1699,
    # Century Gothic carries a large hhea.lineGap that Word adds as external
    # leading on EVERY line (the classic "Century Gothic is double spaced"
    # effect). Validated against Word's own render of a Futura->CG
    # substituted document: 14pt line = 29.7px -> factor 1.594.
    "century gothic": 1.5940,
    "candara": 1.2207, "constantia": 1.2168, "corbel": 1.2207, "consolas": 1.1719,
    "franklin gothic book": 1.1367, "franklin gothic medium": 1.1367,
    "gill sans mt": 1.1621,
    "rockwell": 1.1719,
    "comic sans ms": 1.3945,
    "impact": 1.2188,
    "sitka": 1.3242,
    # Aptos family (M365 default since 2023; a.k.a. Bierstadt).
    # 1.2847 measured from the actual OS/2+hhea tables of the released TTFs.
    "aptos": 1.2847, "aptos display": 1.2847, "aptos light": 1.2847,
    "aptos semibold": 1.2847, "aptos black": 1.2847, "aptos extrabold": 1.2847,
    "aptos narrow": 1.2847, "aptos serif": 1.2847, "aptos mono": 1.2847,
    # NOTE: deliberately no "futura" entry -- Futura is almost never
    # installed; Word substitutes Century Gothic, so line_factor() follows
    # the SUBSTITUTES chain to CG's metrics, matching what Word shows.
    "wingdings": 1.1000, "symbol": 1.2000, "webdings": 1.1000,
}
DEFAULT_FACTOR = 1.2

# Families Word commonly substitutes when the named font is missing.  Used to
# build the CSS fallback chain so a font-less viewer degrades the way Word
# would, rather than to the browser's default sans.
SUBSTITUTES = {
    "aptos": "Calibri", "aptos display": "Calibri Light", "aptos light": "Calibri Light",
    "aptos serif": "Cambria", "aptos mono": "Consolas",
    "futura": "Century Gothic", "futura md bt": "Century Gothic",
    "futura bk bt": "Century Gothic",
    "helvetica": "Arial", "helvetica neue": "Arial",
    "avenir": "Segoe UI", "avenir next": "Segoe UI",
    "gotham": "Montserrat", "proxima nova": "Segoe UI",
    "myriad pro": "Segoe UI", "minion pro": "Cambria",
}

_WEIGHT_WORDS = (
    ("thin", 100), ("hairline", 100), ("extralight", 200), ("ultralight", 200),
    ("semilight", 350), ("light", 300), ("medium", 500), ("demibold", 600),
    ("semibold", 600), ("extrabold", 800), ("ultrabold", 800), ("heavy", 900),
    ("black", 900), ("bold", 700),
)


class FontFace:
    __slots__ = ("path", "index", "family", "subfamily", "weight", "italic",
                 "factor", "full_name")

    def __init__(self, path, index, family, subfamily, weight, italic, factor, full_name):
        self.path = path
        self.index = index          # face index inside a .ttc, else None
        self.family = family
        self.subfamily = subfamily or ""
        self.weight = weight
        self.italic = italic
        self.factor = factor
        self.full_name = full_name or family


# ---------------------------------------------------------------------------
# sfnt parsing (TTF/OTF/TTC) - just the head/hhea/OS2/name tables.
# ---------------------------------------------------------------------------
def _u16(b, o):
    return struct.unpack_from(">H", b, o)[0]


def _s16(b, o):
    return struct.unpack_from(">h", b, o)[0]


def _u32(b, o):
    return struct.unpack_from(">I", b, o)[0]


def _parse_name_table(data, off):
    """Return {nameID: best string} preferring Windows/en records."""
    try:
        count = _u16(data, off + 2)
        str_off = off + _u16(data, off + 4)
        out = {}
        score = {}
        for i in range(count):
            rec = off + 6 + i * 12
            plat = _u16(data, rec)
            enc = _u16(data, rec + 2)
            lang = _u16(data, rec + 4)
            nid = _u16(data, rec + 6)
            ln = _u16(data, rec + 8)
            so = _u16(data, rec + 10)
            if nid not in (1, 2, 4, 16, 17):
                continue
            raw = data[str_off + so: str_off + so + ln]
            if plat == 3:  # Windows, UTF-16BE
                try:
                    s = raw.decode("utf-16-be")
                except UnicodeDecodeError:
                    continue
                sc = 3 if (lang & 0xFF) == 0x09 else 2  # prefer English
            elif plat == 1:  # Mac Roman
                s = raw.decode("mac_roman", "replace")
                sc = 1
            else:
                continue
            if sc >= score.get(nid, 0):
                score[nid] = sc
                out[nid] = s.strip("\x00").strip()
        return out
    except (struct.error, IndexError):
        return {}


def _parse_sfnt(data, base=0):
    """Parse one font (at offset `base` for TTC members). Returns dict or None."""
    try:
        numtables = _u16(data, base + 4)
        tables = {}
        for i in range(numtables):
            rec = base + 12 + 16 * i
            tag = data[rec:rec + 4].decode("latin-1")
            tables[tag] = (_u32(data, rec + 8), _u32(data, rec + 12))
        if "head" not in tables or "name" not in tables:
            return None
        head_off = tables["head"][0]
        upm = _u16(data, head_off + 18)
        mac_style = _u16(data, head_off + 44)
        if not upm:
            return None
        win_asc = win_desc = None
        weight = 400
        if "OS/2" in tables:
            os2 = tables["OS/2"][0]
            weight = _u16(data, os2 + 4) or 400
            win_asc = _u16(data, os2 + 74)
            win_desc = _u16(data, os2 + 76)
        hhea_asc = hhea_desc = line_gap = 0
        if "hhea" in tables:
            hh = tables["hhea"][0]
            hhea_asc = _s16(data, hh + 4)
            hhea_desc = _s16(data, hh + 6)
            line_gap = _s16(data, hh + 8)
        if win_asc is None:
            win_asc, win_desc = hhea_asc, -hhea_desc
        ext = max(0, line_gap - ((win_asc + win_desc) - (hhea_asc - hhea_desc)))
        factor = (win_asc + win_desc + ext) / float(upm)
        names = _parse_name_table(data, tables["name"][0])
        family = names.get(16) or names.get(1)
        sub = names.get(17) or names.get(2) or ""
        if not family:
            return None
        italic = bool(mac_style & 2) or "italic" in sub.lower() or "oblique" in sub.lower()
        subl = sub.lower().replace(" ", "")
        for word, wval in _WEIGHT_WORDS:
            if word in subl:
                weight = wval
                break
        return {"family": family, "sub": sub, "weight": weight, "italic": italic,
                "factor": factor, "full": names.get(4)}
    except (struct.error, IndexError, UnicodeDecodeError):
        return None


def parse_font_file(path):
    """Yield FontFace for each face in a TTF/OTF/TTC file (header-only read)."""
    try:
        with open(path, "rb") as f:
            data = f.read()
    except OSError:
        return
    if len(data) < 12:
        return
    tag = data[:4]
    if tag == b"ttcf":
        try:
            n = _u32(data, 8)
            offsets = [_u32(data, 12 + 4 * i) for i in range(min(n, 64))]
        except struct.error:
            return
        for idx, off in enumerate(offsets):
            info = _parse_sfnt(data, off)
            if info:
                yield FontFace(path, idx, info["family"], info["sub"], info["weight"],
                               info["italic"], info["factor"], info["full"])
    elif tag in (b"\x00\x01\x00\x00", b"OTTO", b"true"):
        info = _parse_sfnt(data, 0)
        if info:
            yield FontFace(path, None, info["family"], info["sub"], info["weight"],
                           info["italic"], info["factor"], info["full"])


# ---------------------------------------------------------------------------
# System font registry (lazy singleton)
# ---------------------------------------------------------------------------
_FONT_DIRS = None
_REGISTRY = None  # lowercase family -> [FontFace]


def _font_dirs():
    global _FONT_DIRS
    if _FONT_DIRS is None:
        dirs = []
        # Extra dirs (testing / portable fonts), highest priority.
        extra = os.environ.get("DOCX2HTML_FONT_DIRS")
        if extra:
            dirs.extend(p for p in extra.split(os.pathsep) if p)
        windir = os.environ.get("WINDIR", r"C:\Windows")
        dirs.append(os.path.join(windir, "Fonts"))
        lad = os.environ.get("LOCALAPPDATA")
        if lad:
            dirs.append(os.path.join(lad, "Microsoft", "Windows", "Fonts"))
        _FONT_DIRS = [d for d in dirs if os.path.isdir(d)]
    return _FONT_DIRS


def _registry():
    """Scan system font files once per process; index faces by family name."""
    global _REGISTRY
    if _REGISTRY is not None:
        return _REGISTRY
    reg = {}
    for d in _font_dirs():
        try:
            entries = os.listdir(d)
        except OSError:
            continue
        for fn in entries:
            if not fn.lower().endswith((".ttf", ".otf", ".ttc")):
                continue
            for face in parse_font_file(os.path.join(d, fn)):
                reg.setdefault(face.family.lower(), []).append(face)
                # Also index "Family Subfamily" for styled families that CSS
                # can reference directly (e.g. 'Segoe UI Semibold').
                if face.subfamily and face.subfamily.lower() not in ("regular", "normal",
                                                                     "book", "roman"):
                    combo = ("%s %s" % (face.family, face.subfamily)).lower()
                    reg.setdefault(combo, []).append(face)
    _REGISTRY = reg
    return reg


def find_faces(family):
    """All installed faces for a family name (case-insensitive)."""
    if not family:
        return []
    return _registry().get(family.strip().lower(), [])


_factor_cache = {}


def _probe_factor(key):
    faces = find_faces(key)
    if faces:
        regular = min(faces, key=lambda f: (f.italic, abs(f.weight - 400)))
        return regular.factor
    return None


def line_factor(family):
    """
    Word-single-spacing multiplier for a font family (1 line == factor em).
    Resolution order mirrors what Word would actually render with:
      1. the font's real metrics if installed locally,
      2. a table of known Office/Windows fonts,
      3. the metrics of Word's SUBSTITUTE for a missing font (chained),
      4. the base family with style suffixes stripped,
      5. 1.2.
    """
    if not family:
        return DEFAULT_FACTOR
    key = family.strip().lower()
    if key in _factor_cache:
        return _factor_cache[key]
    val = _probe_factor(key)
    if val is None:
        val = KNOWN_FACTORS.get(key)
    if val is None:
        # Word substitution chain (e.g. Futura -> Century Gothic).
        seen = {key}
        sub = SUBSTITUTES.get(key)
        while val is None and sub and sub.lower() not in seen:
            skey = sub.lower()
            seen.add(skey)
            val = _probe_factor(skey)
            if val is None:
                val = KNOWN_FACTORS.get(skey)
            sub = SUBSTITUTES.get(skey)
    if val is None:
        # strip style suffixes: 'Aptos Display' -> 'aptos'
        base = key
        for suffix in (" display", " light", " semilight", " semibold", " medium",
                       " black", " extrabold", " condensed", " narrow"):
            if base.endswith(suffix):
                base = base[: -len(suffix)]
                break
        if base != key:
            val = KNOWN_FACTORS.get(base) or _probe_factor(base)
    if val is None:
        val = DEFAULT_FACTOR
    val = round(val, 4)
    _factor_cache[key] = val
    return val


def substitute(family):
    """Word-style substitution target for a missing family (or None)."""
    if not family:
        return None
    return SUBSTITUTES.get(family.strip().lower())


# ---------------------------------------------------------------------------
# Opt-in embedding: referenced families -> @font-face CSS from local files
# ---------------------------------------------------------------------------
def _extract_face_bytes(face):
    """Font program bytes for one face. TTC members are re-assembled into a
    standalone sfnt so browsers can load them."""
    with open(face.path, "rb") as f:
        data = f.read()
    if face.index is None:
        return data
    # Rebuild a single-font sfnt from the TTC member's table directory.
    base = _u32(data, 12 + 4 * face.index)
    numtables = _u16(data, base + 4)
    records = []
    for i in range(numtables):
        rec = base + 12 + 16 * i
        tag = data[rec:rec + 4]
        checksum = _u32(data, rec + 4)
        off = _u32(data, rec + 8)
        ln = _u32(data, rec + 12)
        records.append((tag, checksum, off, ln))
    header = data[base:base + 12]
    out_tables = []
    running = 12 + 16 * numtables
    directory = b""
    for tag, checksum, off, ln in records:
        blob = data[off:off + ln]
        pad = (-len(blob)) % 4
        directory += tag + struct.pack(">III", checksum, running, ln)
        out_tables.append(blob + b"\x00" * pad)
        running += ln + pad
    return header + directory + b"".join(out_tables)


def _mime_for(data):
    head = data[:4]
    if head == b"OTTO":
        return "font/otf", "opentype"
    if head == b"wOFF":
        return "font/woff", "woff"
    if head == b"wOF2":
        return "font/woff2", "woff2"
    return "font/ttf", "truetype"


def embed_css_for_families(families, already_embedded=(), max_bytes_per_face=6_000_000):
    """
    @font-face rules for every referenced family found on this machine.
    `already_embedded` families (e.g. fonts recovered from the docx itself)
    are skipped. Whole-file embedding: correct but heavy -- opt-in by design.
    """
    skip = {f.strip().lower() for f in already_embedded if f}
    rules = []
    seen_keys = set()
    for fam in sorted({(f or "").strip() for f in families if f}):
        low = fam.lower()
        if not low or low in skip:
            continue
        faces = find_faces(low)
        if not faces:
            continue
        # Regular / bold / italic / bold-italic at most, dedup by (w, i).
        picked = {}
        for face in faces:
            wkey = 700 if face.weight >= 600 else (face.weight if face.weight != 400 else 400)
            k = (wkey, face.italic)
            cur = picked.get(k)
            if cur is None or abs(face.weight - wkey) < abs(cur.weight - wkey):
                picked[k] = face
        for (w, italic), face in sorted(picked.items()):
            dedup = (low, w, italic)
            if dedup in seen_keys:
                continue
            seen_keys.add(dedup)
            try:
                blob = _extract_face_bytes(face)
            except (OSError, struct.error):
                continue
            if len(blob) > max_bytes_per_face:
                continue
            mime, fmt = _mime_for(blob)
            b64 = base64.b64encode(blob).decode("ascii")
            rules.append(
                "@font-face{font-family:'%s';font-weight:%d;font-style:%s;"
                "src:url(data:%s;base64,%s) format('%s');font-display:block;}"
                % (fam.replace("'", ""), w, "italic" if italic else "normal",
                   mime, b64, fmt))
    return "\n".join(rules)
