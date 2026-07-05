"""
fancydocx - pure-Python DOCX -> single self-contained HTML converter.

    import fancydocx
    fancydocx.convert("resume.docx", "resume.html")   # write a file
    html = fancydocx.convert("resume.docx")            # or get the HTML string

No external engines, no LibreOffice, no network. Images are inlined as data
URIs and embedded fonts are recovered as @font-face, so the output is one
portable .html file.
"""
from __future__ import annotations
import html as _html
import pathlib

from .core import local
from .package import DocxPackage
from .theme import Theme
from .styles import Styles, rpr_to_css, ppr_to_css, line_height_css
from .numbering import Numbering
from .render import Converter
from .fontmetrics import embed_css_for_families

__version__ = "0.1.0"
__all__ = ["convert", "convert_docx", "convert_file", "DocxPackage", "__version__"]

BASE_CSS = """
*{box-sizing:border-box}
html,body{margin:0;padding:0}
body{background:#e9e9ee;color:#000;-webkit-print-color-adjust:exact;print-color-adjust:exact;
     text-rendering:geometricPrecision}
.docx-doc{padding:24px 12px}
/* isolation:isolate makes each page its own stacking context, so that
   z-index:-1 layers (header/footer art, behindDoc shapes) paint ABOVE the
   page's own background but BELOW in-flow content -- exactly Word's
   page-color / behind-text / text layering. Without it, negative z-index
   children fall behind the page background and vanish. */
.docx-page{position:relative;background:#fff;margin:0 auto 24px;
           box-shadow:0 2px 14px rgba(0,0,0,.28);overflow:hidden;isolation:isolate}
/* .docx-body is intentionally NOT positioned so absolutely-positioned floats
   (anchored images/shapes) resolve against the .docx-page box = true page
   coordinates, matching Word's page-relative anchoring. */
.docx-page p{margin:0}
.docx-page table{border-spacing:0;max-width:none;border-collapse:collapse}
.docx-page td,.docx-page th{vertical-align:top}
.docx-page img{max-width:none}
.docx-page a{color:inherit;text-decoration:inherit}
.leader{flex:1 1 auto;align-self:flex-end;border-bottom:1px dotted currentColor;margin:0 4px 3px}
.tab{display:inline-block;min-width:2em}
.docx-header,.docx-footer{pointer-events:none}
@media print{
 html,body{background:#fff}
 .docx-doc{padding:0}
 .docx-page{box-shadow:none;margin:0;page-break-after:always}
 @page{margin:0}
}
"""


def _title(pkg, path):
    core = pkg.xml("docProps/core.xml")
    if core is not None:
        for el in core.iter():
            if local(el.tag) == "title" and el.text:
                return el.text.strip()
    return pathlib.Path(str(path)).stem


def _body_rule(styles, theme):
    """Default inherited run/paragraph look, applied to .docx-body."""
    rpr = styles.effective_rpr(None, None, {})
    ppr = styles.effective_ppr(None, {})
    d = rpr_to_css(rpr, theme)
    out = {}
    for k in ("font-family", "font-size", "color"):
        if k in d:
            out[k] = d[k]
    # Word single spacing is font-metric based (see fontmetrics.py); the
    # numeric factor keeps the geometry even under font substitution.
    out["line-height"] = line_height_css(ppr.get("spacing"),
                                         rpr.get("font"), rpr.get("sz") or 11.0)
    out.setdefault("font-family", "'Calibri', 'Segoe UI', sans-serif")
    out.setdefault("font-size", "11pt")
    out["word-wrap"] = "break-word"
    return ".docx-body{%s}" % ";".join("%s:%s" % (k, v) for k, v in out.items())


def convert_docx(path, include_headers=True, embed_fonts=False):
    """
    Convert a .docx file to a single self-contained HTML string.

    embed_fonts: additionally inline every referenced font family found on
    THIS machine as base64 @font-face. This makes the HTML render with the
    exact intended glyph metrics on any viewer, at the cost of several MB
    per file -- off by default for batch conversions.
    """
    pkg = DocxPackage(path)
    try:
        theme = Theme(pkg)
        styles = Styles(pkg, theme)
        numbering = Numbering(pkg, theme)
        conv = Converter(pkg, theme, styles, numbering, include_headers=include_headers)
        body = conv.render_document()
        font_css, doc_families = pkg.font_face_css_and_families()
        if embed_fonts:
            local_css = embed_css_for_families(conv.used_fonts, already_embedded=doc_families)
            if local_css:
                font_css = font_css + "\n" + local_css if font_css else local_css
        body_rule = _body_rule(styles, theme)
        title = _title(pkg, path)
    finally:
        pkg.close()

    return (
        "<!doctype html>\n<html lang=\"en\">\n<head>\n<meta charset=\"utf-8\">\n"
        "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n"
        "<title>%s</title>\n<style>\n%s\n%s\n%s\n</style>\n</head>\n<body>\n"
        "<div class=\"docx-doc\">%s</div>\n</body>\n</html>\n"
        % (_html.escape(title), BASE_CSS, body_rule, font_css, body)
    )


def convert_file(in_path, out_path, include_headers=True, embed_fonts=False):
    """Convert one .docx to one .html on disk. Returns the output path."""
    result = convert_docx(in_path, include_headers=include_headers,
                          embed_fonts=embed_fonts)
    out = pathlib.Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(result, encoding="utf-8")
    return str(out)


def convert(source, output=None, *, embed_fonts=False, include_headers=True):
    """
    One-line entry point.

        import fancydocx
        fancydocx.convert("resume.docx", "resume.html")   # write the file, returns path
        html = fancydocx.convert("resume.docx")            # no output -> returns HTML str

    Parameters
    ----------
    source : str | os.PathLike
        Path to the input .docx file.
    output : str | os.PathLike | None
        Where to write the HTML. If None, the HTML is returned as a string.
    embed_fonts : bool
        Inline locally-installed referenced fonts as base64 @font-face
        (exact metrics on any viewer, at the cost of file size).
    include_headers : bool
        Render document headers/footers (default True).
    """
    if output is None:
        return convert_docx(source, include_headers=include_headers, embed_fonts=embed_fonts)
    return convert_file(source, output, include_headers=include_headers, embed_fonts=embed_fonts)
