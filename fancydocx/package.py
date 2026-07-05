"""
OPC package access: unzip the .docx, resolve relationships, inline media as
data URIs, and recover embedded (obfuscated) fonts as @font-face rules.

The embedded-font recovery is what keeps text from reflowing: many designed
templates embed their display fonts. Word obfuscates each font file by XOR-ing
its first 32 bytes with the 16-byte, byte-reversed GUID from `w:fontKey`.
We reverse that and re-emit the raw TTF/OTF as base64 @font-face, so the HTML
renders with the exact same metrics even on a machine that lacks the font.
"""
from __future__ import annotations
import base64
import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET

from .core import NS, qn, local

# Relationship type suffixes we care about.
RT_OFFICE_DOC = "officeDocument"
IMAGE_MIME = {
    "png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
    "gif": "image/gif", "bmp": "image/bmp", "tif": "image/tiff",
    "tiff": "image/tiff", "svg": "image/svg+xml", "webp": "image/webp",
    "emf": "image/emf", "wmf": "image/wmf", "ico": "image/x-icon",
}


class Relationship:
    __slots__ = ("id", "type", "target", "mode")

    def __init__(self, rid, rtype, target, mode):
        self.id = rid
        self.type = rtype
        self.target = target
        self.mode = mode

    @property
    def external(self):
        return self.mode == "External"


class DocxPackage:
    def __init__(self, path):
        self.path = str(path)
        self.zip = zipfile.ZipFile(self.path, "r")
        self.names = set(self.zip.namelist())
        self._xml_cache = {}
        self._rels_cache = {}
        self._datauri_cache = {}

        # Locate the main document part via the package root relationships.
        self.main_part = self._find_main_part()
        self.main_dir = posixpath.dirname(self.main_part)
        self.doc_rels = self.load_rels(self.main_part)

    # -- raw part access ---------------------------------------------------
    def has(self, name):
        return name in self.names

    def part(self, name):
        return self.zip.read(name)

    def xml(self, name):
        if name in self._xml_cache:
            return self._xml_cache[name]
        root = None
        if name in self.names:
            try:
                root = ET.fromstring(self.zip.read(name))
            except ET.ParseError:
                root = None
        self._xml_cache[name] = root
        return root

    # -- relationships -----------------------------------------------------
    def _find_main_part(self):
        root = self.xml("_rels/.rels")
        if root is not None:
            for rel in root:
                if local(rel.tag) != "Relationship":
                    continue
                if rel.get("Type", "").endswith(RT_OFFICE_DOC):
                    return rel.get("Target").lstrip("/")
        # Fallback to the conventional location.
        return "word/document.xml"

    def load_rels(self, part_name):
        """Return {rId: Relationship} for the given part (cached)."""
        if part_name in self._rels_cache:
            return self._rels_cache[part_name]
        base = posixpath.dirname(part_name)
        rels_path = posixpath.join(base, "_rels", posixpath.basename(part_name) + ".rels")
        result = {}
        root = self.xml(rels_path)
        if root is not None:
            for rel in root:
                if local(rel.tag) != "Relationship":
                    continue
                result[rel.get("Id")] = Relationship(
                    rel.get("Id"), rel.get("Type", ""), rel.get("Target", ""),
                    rel.get("TargetMode", "Internal"),
                )
        self._rels_cache[part_name] = result
        return result

    def resolve_target(self, target, base_dir=None):
        """Resolve a (possibly ../-relative) rel target to a zip part name."""
        if base_dir is None:
            base_dir = self.main_dir
        if target.startswith("/"):
            return target.lstrip("/")
        return posixpath.normpath(posixpath.join(base_dir, target))

    # -- media -------------------------------------------------------------
    def media_ext(self, rels, rid):
        """Lowercase file extension of an image relationship ('' if unknown)."""
        rel = rels.get(rid) if rid else None
        if rel is None or rel.external:
            return ""
        part = self.resolve_target(rel.target)
        return part.rsplit(".", 1)[-1].lower() if "." in part else ""

    def data_uri(self, rels, rid):
        """Resolve an r:embed/r:id image relationship to a base64 data URI."""
        if not rid:
            return None
        rel = rels.get(rid)
        if rel is None:
            return None
        if rel.external:
            return rel.target  # remote URL, leave as-is
        part = self.resolve_target(rel.target)
        if part in self._datauri_cache:
            return self._datauri_cache[part]
        if part not in self.names:
            self._datauri_cache[part] = None
            return None
        ext = part.rsplit(".", 1)[-1].lower()
        mime = IMAGE_MIME.get(ext, "application/octet-stream")
        b64 = base64.b64encode(self.zip.read(part)).decode("ascii")
        uri = "data:%s;base64,%s" % (mime, b64)
        self._datauri_cache[part] = uri
        return uri

    # -- embedded fonts ----------------------------------------------------
    def font_face_css(self):
        css, _families = self.font_face_css_and_families()
        return css

    def font_face_css_and_families(self):
        """Recover embedded fonts -> (@font-face CSS, {family names})."""
        families = set()
        font_table = None
        for name in self.names:
            if re.match(r"word/fontTable\d*\.xml$", name):
                font_table = name
                break
        if font_table is None:
            return "", families
        root = self.xml(font_table)
        if root is None:
            return "", families
        rels = self.load_rels(font_table)
        slots = [
            ("embedRegular", "normal", "normal"),
            ("embedBold", "bold", "normal"),
            ("embedItalic", "normal", "italic"),
            ("embedBoldItalic", "bold", "italic"),
        ]
        rules = []
        for font in root.findall("w:font", NS):
            name = font.get(qn("w:name"))
            if not name:
                continue
            for tag, weight, style in slots:
                node = font.find("w:" + tag, NS)
                if node is None:
                    continue
                rid = node.get(qn("r:id"))
                key = node.get(qn("w:fontKey"))
                rel = rels.get(rid)
                if rel is None:
                    continue
                part = self.resolve_target(rel.target, base_dir="word")
                if part not in self.names:
                    continue
                data = self._deobfuscate(self.zip.read(part), key)
                fmt, mime = self._font_format(data)
                b64 = base64.b64encode(data).decode("ascii")
                families.add(name)
                rules.append(
                    "@font-face{font-family:%s;font-weight:%s;font-style:%s;"
                    "src:url(data:%s;base64,%s) format('%s');font-display:block;}"
                    % (_css_font_name(name), weight, style, mime, b64, fmt)
                )
        return "\n".join(rules), families

    @staticmethod
    def _deobfuscate(data, font_key):
        """Undo Word's font obfuscation (XOR first 32 bytes w/ reversed GUID)."""
        if not font_key:
            return data
        hexs = re.sub(r"[^0-9a-fA-F]", "", font_key)
        if len(hexs) != 32:
            return data
        key = bytes.fromhex(hexs)[::-1]  # GUID bytes, reversed
        buf = bytearray(data)
        for i in range(min(32, len(buf))):
            buf[i] ^= key[i % 16]
        return bytes(buf)

    @staticmethod
    def _font_format(data):
        head = data[:4]
        if head == b"OTTO":
            return "opentype", "font/otf"
        if head == b"wOFF":
            return "woff", "font/woff"
        if head == b"wOF2":
            return "woff2", "font/woff2"
        return "truetype", "font/ttf"

    def close(self):
        try:
            self.zip.close()
        except Exception:
            pass


def _css_font_name(name):
    """Quote a font family name for CSS."""
    return '"%s"' % name.replace('"', "")
