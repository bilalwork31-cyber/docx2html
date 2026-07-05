"""
The document walker: OOXML body -> HTML.

Model: each section becomes a fixed-size ".docx-page" box (page size, padding =
margins) so the result looks like Word's print view. Normal-flow content sits
in the padding area; floating shapes/images are absolutely positioned relative
to the page box (coordinates computed from the page origin). Tables reproduce
the exact grid via <colgroup> + table-layout:fixed, with a proper occupancy
grid for gridSpan/vMerge. Every risky sub-render is wrapped so a single odd
element never kills a whole document -- important when batching thousands.
"""
from __future__ import annotations
import html
import re
import colorsys

from .core import (
    NS, qn, local, get, find, findall, child, children, int_val,
    twips_to_px, emu_to_px, pt_to_px, px, pt, PT_TO_PX,
)
from .color import hex_to_rgb, rgb_to_hex, normalize_hex
from .styles import (
    parse_ppr, parse_rpr, ppr_to_css, rpr_to_css, resolve_shd, parse_shd,
    parse_borders, border_css_value, merge_props, quote_font, line_height_css,
)

# DrawingML scheme-color name -> theme slot.
DML_SCHEME = {
    "bg1": "background1", "tx1": "text1", "bg2": "background2", "tx2": "text2",
    "dk1": "dk1", "lt1": "lt1", "dk2": "dk2", "lt2": "lt2",
    "accent1": "accent1", "accent2": "accent2", "accent3": "accent3",
    "accent4": "accent4", "accent5": "accent5", "accent6": "accent6",
    "hlink": "hyperlink", "folHlink": "followedHyperlink",
}
VALIGN = {"top": "top", "center": "middle", "both": "middle", "bottom": "bottom"}


def style_attr(d):
    if not d:
        return ""
    s = ";".join("%s:%s" % (k, v) for k, v in d.items() if v is not None and v != "")
    return ' style="%s"' % s if s else ""


def _len_to_px(tok):
    """Parse a CSS/VML length token ('10pt', '1in', '90px') -> px float."""
    if tok is None:
        return None
    tok = str(tok).strip().lower()
    m = re.match(r"^(-?[\d.]+)\s*(pt|px|in|pc|mm|cm|em|)$", tok)
    if not m:
        return None
    v = float(m.group(1))
    factor = {"pt": PT_TO_PX, "px": 1.0, "in": 96.0, "pc": 16.0,
              "mm": 96.0 / 25.4, "cm": 96.0 / 2.54, "em": 16.0, "": 1.0}[m.group(2)]
    return v * factor


class Converter:
    def __init__(self, pkg, theme, styles, numbering, include_headers=True):
        self.pkg = pkg
        self.theme = theme
        self.styles = styles
        self.numbering = numbering
        self.include_headers = include_headers
        self.rels_stack = [pkg.doc_rels]
        self.list_counters = {}
        self.bg_color = None
        self.page = self._default_page()
        self.document = pkg.xml(pkg.main_part)
        self.hf_ctx = None  # None | 'header' | 'footer' (changes anchor bases)
        self.used_fonts = set()  # families referenced by the document
        # Body-default run look, to avoid repeating it on every paragraph.
        body_rpr = styles.effective_rpr(None, None, {})
        self.body_font = body_rpr.get("font")
        self.body_sz = body_rpr.get("sz") or 11.0
        if self.body_font:
            self.used_fonts.add(self.body_font)
        body_ppr = styles.effective_ppr(None, {})
        self.body_line = line_height_css(body_ppr.get("spacing"),
                                         self.body_font, self.body_sz)

    @property
    def rels(self):
        return self.rels_stack[-1]

    # ==================================================================
    # Document / sections / pages
    # ==================================================================
    def _default_page(self):
        return {"w": twips_to_px(12240), "h": twips_to_px(15840),
                "ml": twips_to_px(1440), "mr": twips_to_px(1440),
                "mt": twips_to_px(1440), "mb": twips_to_px(1440),
                "header": twips_to_px(720), "footer": twips_to_px(720),
                "cols": 1, "colgap": twips_to_px(720)}

    def render_document(self):
        root = self.document
        if root is None:
            return "<p>Could not read document.xml</p>"
        body = find(root, "w:body")
        if body is None:
            return ""
        bg = find(root, "w:background")
        if bg is not None:
            self.bg_color = normalize_hex(get(bg, "w:color"))
        sections = self._split_sections(body)
        return "".join(self._render_section(blocks, sectPr) for blocks, sectPr in sections)

    def _split_sections(self, body):
        sections, current = [], []
        for node in body:
            lt = local(node.tag)
            if lt == "sectPr":                       # final body-level section
                sections.append((current, node))
                current = []
            elif lt == "p":
                sp = find(child(node, "w:pPr"), "w:sectPr") if child(node, "w:pPr") is not None else None
                current.append(node)
                if sp is not None:                   # this paragraph ends a section
                    sections.append((current, sp))
                    current = []
            else:
                current.append(node)
        if current or not sections:
            sections.append((current, None))
        return sections

    def _geometry(self, sectPr):
        g = self._default_page()
        if sectPr is None:
            return g
        pgsz = find(sectPr, "w:pgSz")
        if pgsz is not None:
            g["w"] = twips_to_px(get(pgsz, "w:w")) or g["w"]
            g["h"] = twips_to_px(get(pgsz, "w:h")) or g["h"]
        pgmar = find(sectPr, "w:pgMar")
        if pgmar is not None:
            for k, a in (("mt", "w:top"), ("mb", "w:bottom"), ("ml", "w:left"),
                         ("mr", "w:right"), ("header", "w:header"), ("footer", "w:footer")):
                v = twips_to_px(get(pgmar, a))
                if v is not None:
                    g[k] = abs(v)
        cols = find(sectPr, "w:cols")
        if cols is not None:
            try:
                g["cols"] = int(get(cols, "w:num") or 1)
            except ValueError:
                g["cols"] = 1
            g["colgap"] = twips_to_px(get(cols, "w:space")) or g["colgap"]
        return g

    def _render_section(self, blocks, sectPr):
        self.page = self._geometry(sectPr)
        pages, cur = [], []
        styles_seq = [self._para_style_of(n) for n in blocks]
        for i, node in enumerate(blocks):
            lt = local(node.tag)
            prev_sid = styles_seq[i - 1] if i > 0 else None
            next_sid = styles_seq[i + 1] if i + 1 < len(blocks) else None
            try:
                if lt == "p":
                    # Break BEFORE this paragraph when Word itself started a
                    # new page here: w:pageBreakBefore, or the
                    # w:lastRenderedPageBreak marker Word saves at the exact
                    # spot its layout engine paginated. Splitting on it
                    # reproduces Word's automatic pagination for free.
                    if cur and self._breaks_before(node):
                        pages.append("".join(cur))
                        cur = []
                    h, brk = self.render_paragraph(node, prev_style=prev_sid,
                                                   next_style=next_sid)
                    cur.append(h)
                    if brk:
                        pages.append("".join(cur))
                        cur = []
                elif lt == "tbl":
                    cur.append(self.render_table(node))
                elif lt == "sdt":
                    c = child(node, "w:sdtContent")
                    if c is not None:
                        cur.append(self.render_block_children(c))
                elif lt == "AlternateContent":
                    pick = find(node, "mc:Choice") or find(node, "mc:Fallback")
                    if pick is not None:
                        cur.append(self.render_block_children(pick))
            except Exception as e:  # never let one block kill the page
                cur.append("<!-- block error: %s -->" % html.escape(str(e)))
        pages.append("".join(cur))
        hf = self._render_headers_footers(sectPr) if self.include_headers else ""
        return "".join(self._page_box(body_html, hf) for body_html in pages)

    def _para_style_of(self, node):
        """Paragraph style id of a block node ('' when styleless, None for
        non-paragraphs). Used for contextualSpacing adjacency checks."""
        if local(node.tag) != "p":
            return None
        ppr = child(node, "w:pPr")
        ps = child(ppr, "w:pStyle") if ppr is not None else None
        return (ps.get(qn("w:val")) or "") if ps is not None else ""

    def _page_box(self, body_html, hf):
        g = self.page
        style = {
            "width": px(g["w"]), "min-height": px(g["h"]),
            "padding": "%s %s %s %s" % (px(g["mt"]), px(g["mr"]), px(g["mb"]), px(g["ml"])),
        }
        if self.bg_color:
            style["background"] = "#" + self.bg_color
        body_style = {}
        if g["cols"] and g["cols"] > 1:
            body_style["column-count"] = str(g["cols"])
            body_style["column-gap"] = px(g["colgap"])
        return ('<section class="docx-page"%s>%s<div class="docx-body"%s>%s</div></section>'
                % (style_attr(style), hf, style_attr(body_style), body_html))

    def _render_headers_footers(self, sectPr):
        if sectPr is None:
            return ""
        out = []
        specs = (("w:headerReference", "header", True), ("w:footerReference", "footer", False))
        for tag, kind, is_header in specs:
            for ref in children(sectPr, tag):
                if get(ref, "w:type") == "even":
                    continue
                rid = get(ref, "r:id")
                rel = self.rels.get(rid) if rid else None
                if rel is None:
                    continue
                part = self.pkg.resolve_target(rel.target)
                root = self.pkg.xml(part)
                if root is None:
                    continue
                try:
                    self.rels_stack.append(self.pkg.load_rels(part))
                    self.hf_ctx = kind
                    inner = self.render_block_children(root)
                finally:
                    self.hf_ctx = None
                    self.rels_stack.pop()
                g = self.page
                # Full-page layer (origin = page corner) so page-anchored
                # floats inside the header/footer resolve to true page
                # coordinates; padding places the in-flow header/footer text
                # at the header/footer distance like Word. Negative z-index
                # keeps header/footer design art behind the body content.
                pos = {"position": "absolute", "left": "0", "top": "0",
                       "width": px(g["w"]), "height": px(g["h"]), "z-index": "-1"}
                if is_header:
                    pos["padding"] = "%s %s 0 %s" % (px(g["header"]), px(g["mr"]), px(g["ml"]))
                else:
                    pos["padding"] = "0 %s %s %s" % (px(g["mr"]), px(g["footer"]), px(g["ml"]))
                    pos["display"] = "flex"
                    pos["flex-direction"] = "column"
                    pos["justify-content"] = "flex-end"
                out.append('<div class="docx-%s"%s>%s</div>'
                           % (kind, style_attr(pos), inner))
        return "".join(out)

    def render_block_children(self, parent):
        out = []
        nodes = list(parent)
        styles_seq = [self._para_style_of(n) for n in nodes]
        for i, node in enumerate(nodes):
            lt = local(node.tag)
            prev_sid = styles_seq[i - 1] if i > 0 else None
            next_sid = styles_seq[i + 1] if i + 1 < len(nodes) else None
            try:
                if lt == "p":
                    h, _ = self.render_paragraph(node, prev_style=prev_sid,
                                                 next_style=next_sid)
                    out.append(h)
                elif lt == "tbl":
                    out.append(self.render_table(node))
                elif lt == "sdt":
                    c = child(node, "w:sdtContent")
                    if c is not None:
                        out.append(self.render_block_children(c))
                elif lt == "AlternateContent":
                    pick = find(node, "mc:Choice") or find(node, "mc:Fallback")
                    if pick is not None:
                        out.append(self.render_block_children(pick))
            except Exception as e:
                out.append("<!-- block error: %s -->" % html.escape(str(e)))
        return "".join(out)

    # ==================================================================
    # Paragraphs
    # ==================================================================
    def render_paragraph(self, p, in_cell=False, prev_style=None, next_style=None):
        ppr_el = child(p, "w:pPr")
        direct = parse_ppr(ppr_el, self.theme)
        p_style = direct.get("pStyle")
        eff = self.styles.effective_ppr(p_style, direct)

        marker_html = ""
        numpr = eff.get("numPr")
        if numpr and numpr.get("numId") not in (None, "0"):
            num_id, ilvl = numpr["numId"], numpr.get("ilvl", 0)
            lvl = self.numbering.level(num_id, ilvl)
            if lvl:  # numbering level indent is a base beneath the paragraph
                base = {}
                merge_props(base, lvl["ppr"])
                merge_props(base, eff)
                eff = base
            self._bump_counter(num_id, ilvl)
            mk = self.numbering.marker(num_id, ilvl, lambda l: self._count_of(num_id, l))
            if mk:
                text, mrpr = mk
                suff = lvl["suff"] if lvl else "tab"
                marker_html = self._render_marker(text, mrpr, eff, p_style, suff)

        # The paragraph-mark run look defines the line strut: an empty
        # paragraph is exactly one line of the mark's font/size in Word, and
        # the mark's metrics bound every line's minimum height.
        mark_eff = self.styles.effective_rpr(p_style, None, eff.get("markRpr") or {})
        mark_font = mark_eff.get("font") or self.body_font
        mark_sz = mark_eff.get("sz") or self.body_sz
        if mark_font:
            self.used_fonts.add(mark_font)

        block = ppr_to_css(eff, self.theme, font_family=mark_font, font_size_pt=mark_sz)
        # Always pin the line-height: Word's single spacing is font-metric
        # based; leaving it to inheritance drifts as soon as fonts differ.
        lh = line_height_css(eff.get("spacing"), mark_font, mark_sz)
        if lh != self.body_line or "line-height" in block:
            block["line-height"] = lh
        else:
            block.pop("line-height", None)
        if mark_sz and mark_sz != self.body_sz:
            block.setdefault("font-size", pt(mark_sz))
        if mark_font and mark_font != self.body_font:
            block.setdefault("font-family", quote_font(mark_font))

        # contextualSpacing: suppress the gap between paragraphs of the same
        # style (Word's "don't add space between paragraphs of the same style").
        if eff.get("contextualSpacing"):
            mine = p_style or ""
            if prev_style is not None and prev_style == mine:
                block.pop("padding-top", None)
                block.pop("margin-top", None)
            if next_style is not None and next_style == mine:
                block.pop("margin-bottom", None)

        floats = []
        tokens = list(self._inline_tokens(p, p_style, floats))
        inner, layout = self._assemble_inline(tokens, eff, allow_grid=not marker_html)
        if layout:
            block.update(layout)
        if not inner.strip() and not layout and not marker_html:
            inner = "<br>"
        para = "<p%s>%s%s</p>" % (style_attr(block), marker_html, inner)
        return para + "".join(floats), self._has_page_break(p)

    def _render_marker(self, text, mrpr, eff_ppr, p_style, suff):
        eff = self.styles.effective_rpr(p_style, mrpr.get("rStyle"), mrpr)
        css = rpr_to_css(eff, self.theme)
        css.pop("display", None)
        hang = (eff_ppr.get("ind") or {}).get("hanging")
        css["display"] = "inline-block"
        css["min-width"] = px(twips_to_px(hang)) if hang else "1.4em"
        sfx = " " if suff == "space" else ""
        return "<span%s>%s%s</span>" % (style_attr(css), html.escape(text), sfx)

    def _assemble_inline(self, tokens, eff_ppr, allow_grid=True):
        """
        Join inline tokens, mapping tab characters onto tab-stop layout.
        Returns (html, layout_css_or_None). Two exact patterns:
          * one tab + a right/decimal/center stop -> flex space-between
            (+ optional dot leader), robust at any indent;
          * all-left explicit stops -> CSS grid whose column edges sit at the
            stop positions; minmax(col,max-content) reproduces Word pushing
            later stops when a segment overruns.
        Anything else degrades to a fixed-width tab spacer.
        """
        segs, ntab = [[]], 0
        for kind, val in tokens:
            if kind == "tab":
                segs.append([])
                ntab += 1
            elif kind == "html":
                segs[-1].append(val)
        seg_html = ["".join(s) for s in segs]
        if ntab == 0:
            return seg_html[0], None
        stops = [t for t in (eff_ppr.get("tabs") or []) if t.get("val") != "clear"]
        stops.sort(key=lambda t: t.get("pos", 0))

        if ntab == 1:
            special = next((t for t in stops
                            if t.get("val") in ("right", "end", "decimal", "center")), None)
            if special is not None or not stops:
                leader = (special or {}).get("leader")
                mid = ('<span class="leader"></span>'
                       if leader in ("dot", "middleDot", "underscore", "hyphen") else "")
                layout = {"display": "flex", "align-items": "baseline",
                          "justify-content": "space-between"}
                if special is not None and special.get("val") == "center":
                    layout["justify-content"] = "space-around"
                return ("<span>%s</span>%s<span>%s</span>"
                        % (seg_html[0], mid, seg_html[1]), layout)

        left_stops = [t for t in stops if t.get("val") in (None, "left", "start", "num")]
        if (allow_grid and len(left_stops) == len(stops) and len(stops) >= ntab
                and all(t.get("pos", 0) > 0 for t in stops)):
            # columns [0..s1][s1..s2]...[last..1fr]
            cols, prev = [], 0.0
            for t in stops[:ntab]:
                w = twips_to_px(t["pos"]) - prev
                if w <= 1:
                    w = 1
                cols.append("minmax(%s,max-content)" % px(w))
                prev = twips_to_px(t["pos"])
            cols.append("1fr")
            cells = "".join("<span>%s</span>" % s for s in seg_html)
            return cells, {"display": "grid",
                           "grid-template-columns": " ".join(cols),
                           "align-items": "baseline"}
        return '<span class="tab"></span>'.join(seg_html), None

    def _has_page_break(self, p):
        brk, tp = qn("w:br"), qn("w:type")
        for el in p.iter(brk):
            if el.get(tp) == "page":
                return True
        return False

    def _breaks_before(self, p):
        """True when Word starts a new page AT this paragraph."""
        ppr = child(p, "w:pPr")
        if ppr is not None and child(ppr, "w:pageBreakBefore") is not None:
            v = get(child(ppr, "w:pageBreakBefore"), "w:val")
            if v in (None, "1", "true", "on"):
                return True
        # lastRenderedPageBreak before any text content -> page starts here.
        lrpb = qn("w:lastRenderedPageBreak")
        text = qn("w:t")
        for el in p.iter():
            if el.tag == lrpb:
                return True
            if el.tag == text and (el.text or "").strip():
                return False
        return False

    # -- inline token stream -------------------------------------------
    def _inline_tokens(self, container, p_style, floats):
        for node in container:
            lt = local(node.tag)
            if lt == "r":
                yield from self._run_tokens(node, p_style, floats)
            elif lt == "hyperlink":
                href = None
                rid = get(node, "r:id")
                if rid:
                    rel = self.rels.get(rid)
                    if rel is not None:
                        href = rel.target
                anchor = get(node, "w:anchor")
                if not href and anchor:
                    href = "#" + anchor
                inner = "".join(v for k, v in self._inline_tokens(node, p_style, floats) if k == "html")
                if href:
                    yield "html", '<a href="%s">%s</a>' % (html.escape(href, quote=True), inner)
                else:
                    yield "html", inner
            elif lt in ("ins", "smartTag", "fldSimple"):
                yield from self._inline_tokens(node, p_style, floats)
            elif lt == "sdt":
                c = child(node, "w:sdtContent")
                if c is not None:
                    yield from self._inline_tokens(c, p_style, floats)
            elif lt == "AlternateContent":
                pick = find(node, "mc:Choice") or find(node, "mc:Fallback")
                if pick is not None:
                    yield from self._inline_tokens(pick, p_style, floats)
            # bookmarkStart/End, proofErr, del, commentRange -> ignored

    def _run_tokens(self, r, p_style, floats):
        direct = parse_rpr(child(r, "w:rPr"), self.theme)
        eff = self.styles.effective_rpr(p_style, direct.get("rStyle"), direct)
        if eff.get("font"):
            self.used_fonts.add(eff["font"])
        css = rpr_to_css(eff, self.theme)
        if css.get("display") == "none":
            return
        sattr = style_attr(css)
        buf = []

        def flush():
            if buf:
                text = "".join(buf)
                yield ("html", ("<span%s>%s</span>" % (sattr, text)) if sattr else text)
                buf.clear()

        for c in self._run_children(r):
            lt = local(c.tag)
            if lt == "t":
                buf.append(self._text(c))
            elif lt in ("br", "cr"):
                if lt == "br" and get(c, "w:type") == "page":
                    yield from flush()
                    yield ("pagebreak", None)
                else:
                    buf.append("<br>")
            elif lt == "tab":
                yield from flush()
                yield ("tab", None)
            elif lt == "sym":
                buf.append(self._sym(c))
            elif lt == "noBreakHyphen":
                buf.append("&#8209;")
            elif lt == "softHyphen":
                buf.append("&shy;")
            elif lt == "drawing":
                yield from flush()
                for tok in self._render_drawing(c):
                    if tok[0] == "float":
                        floats.append(tok[1])
                    else:
                        yield tok
            elif lt == "pict":
                yield from flush()
                for tok in self._render_vml(c):
                    if tok[0] == "float":
                        floats.append(tok[1])
                    else:
                        yield tok
        yield from flush()

    def _run_children(self, r):
        """Run children, transparently unwrapping AlternateContent choices."""
        for c in r:
            if local(c.tag) == "AlternateContent":
                pick = find(c, "mc:Choice") or find(c, "mc:Fallback")
                if pick is not None:
                    for sub in pick:
                        yield sub
            else:
                yield c

    def _text(self, t):
        s = t.text or ""
        if not s:
            return ""
        out = html.escape(s)
        # Preserve runs of spaces (HTML collapses them otherwise).
        out = out.replace("  ", "  ")
        # Preserve a significant leading space at a run boundary.
        if s[:1] == " ":
            out = " " + out[1:]
        return out

    def _sym(self, sym):
        font = get(sym, "w:font") or ""
        ch = ""
        code = get(sym, "w:char")
        if code:
            try:
                ch = chr(int(code, 16))
            except ValueError:
                ch = ""
        return "<span style=\"font-family:'%s'\">%s</span>" % (font.replace("'", ""), html.escape(ch))

    # ==================================================================
    # Numbering counters
    # ==================================================================
    def _bump_counter(self, num_id, ilvl):
        key = (num_id, ilvl)
        cur = self.list_counters.get(key)
        self.list_counters[key] = (cur + 1) if cur is not None else self.numbering.start_value(num_id, ilvl)
        for k in list(self.list_counters):
            if k[0] == num_id and k[1] > ilvl:
                del self.list_counters[k]

    def _count_of(self, num_id, ilvl):
        return self.list_counters.get((num_id, ilvl), self.numbering.start_value(num_id, ilvl))

    # ==================================================================
    # Tables
    # ==================================================================
    def render_table(self, tbl):
        tblPr = child(tbl, "w:tblPr")
        grid = child(tbl, "w:tblGrid")
        col_w = [twips_to_px(get(gc, "w:w")) or 0 for gc in children(grid, "w:gridCol")] if grid is not None else []

        style_id = None
        ts = child(tblPr, "w:tblStyle")
        if ts is not None:
            style_id = get(ts, "w:val")
        tstyle = self._table_style_props(style_id)

        tbl_borders = dict(tstyle.get("borders") or {})
        db = parse_borders(child(tblPr, "w:tblBorders"))
        if db:
            tbl_borders.update(db)
        default_cell_shd = tstyle.get("cell_shd")
        first_row = tstyle.get("first_row") or {}

        table_css = {"border-collapse": "collapse", "table-layout": "fixed"}
        total = sum(col_w)
        if total:
            table_css["width"] = px(total)
        tblw = child(tblPr, "w:tblW")
        if tblw is not None and get(tblw, "w:type") == "pct":
            try:
                table_css["width"] = "%g%%" % (float(get(tblw, "w:w")) / 50.0)
            except (TypeError, ValueError):
                pass
        jc = get(child(tblPr, "w:jc"), "w:val") if child(tblPr, "w:jc") is not None else None
        if jc == "center":
            table_css["margin-left"] = "auto"
            table_css["margin-right"] = "auto"
        elif jc in ("right", "end"):
            table_css["margin-left"] = "auto"
        ind = child(tblPr, "w:tblInd")
        if ind is not None:
            table_css["margin-left"] = px(twips_to_px(get(ind, "w:w")))
        tshd = resolve_shd(parse_shd(child(tblPr, "w:shd")), self.theme)
        if tshd:
            table_css["background-color"] = tshd

        colgroup = "".join("<col style=\"width:%s\">" % px(w) for w in col_w) if col_w else ""

        # Table-level default cell margins (fallback for sides a cell's own
        # tcMar doesn't specify). Word defaults: 0 top/bottom, 108tw sides.
        def_mar = {"top": 0.0, "left": twips_to_px(108), "bottom": 0.0,
                   "right": twips_to_px(108)}
        cellmar = child(tblPr, "w:tblCellMar")
        if cellmar is not None:
            for side_name in ("top", "left", "bottom", "right", "start", "end"):
                m = child(cellmar, "w:" + side_name)
                if m is not None and get(m, "w:type") != "nil":
                    v = twips_to_px(get(m, "w:w"))
                    if v is not None:
                        key = {"start": "left", "end": "right"}.get(side_name, side_name)
                        def_mar[key] = v

        rows = children(tbl, "w:tr")
        grid_rows = self._build_grid(rows)
        n_rows = len(grid_rows)

        html_rows = []
        for r_idx, (tr, cells) in enumerate(grid_rows):
            tr_css = {}
            trPr = child(tr, "w:trPr")
            hgt = child(trPr, "w:trHeight")
            if hgt is not None:
                hv = twips_to_px(get(hgt, "w:val"))
                if hv:
                    # CSS height on a table row behaves as a minimum (the row
                    # grows with content), which matches hRule=atLeast (the
                    # default). min-height is IGNORED on table rows, so it
                    # must not be used here. hRule=exact can still grow in
                    # HTML; accepted approximation.
                    tr_css["height"] = px(hv)
            tds = []
            for cell in cells:
                tds.append(self._render_cell(cell, r_idx, n_rows, len(col_w),
                                              tbl_borders, default_cell_shd, first_row,
                                              def_mar))
            html_rows.append("<tr%s>%s</tr>" % (style_attr(tr_css), "".join(tds)))

        return ("<table%s>%s<tbody>%s</tbody></table>"
                % (style_attr(table_css), colgroup, "".join(html_rows)))

    def _build_grid(self, rows):
        """Resolve gridSpan/vMerge into per-cell colspan/rowspan (occupancy)."""
        col_restart = {}
        grid_rows = []
        for tr in rows:
            gridcol = 0
            row_cells = []
            for tc in children(tr, "w:tc"):
                tcPr = child(tc, "w:tcPr")
                gridspan = int_val(tcPr, "w:gridSpan", default=1) or 1
                vmerge = child(tcPr, "w:vMerge")
                if vmerge is not None:
                    vval = get(vmerge, "w:val") or "continue"
                    if vval == "continue":
                        owner = col_restart.get(gridcol)
                        if owner is not None:
                            owner["rowspan"] += 1
                        gridcol += gridspan
                        continue
                cell = {"tc": tc, "tcPr": tcPr, "colspan": gridspan,
                        "rowspan": 1, "gridcol": gridcol}
                row_cells.append(cell)
                if vmerge is not None:  # restart
                    for c in range(gridcol, gridcol + gridspan):
                        col_restart[c] = cell
                else:
                    for c in range(gridcol, gridcol + gridspan):
                        col_restart.pop(c, None)
                gridcol += gridspan
            grid_rows.append((tr, row_cells))
        return grid_rows

    def _render_cell(self, cell, r_idx, n_rows, n_cols, tbl_borders, default_shd,
                     first_row, def_mar=None):
        tc, tcPr = cell["tc"], cell["tcPr"]
        css = {"overflow": "hidden"}
        last_col = cell["gridcol"] + cell["colspan"] >= n_cols

        cell_borders = parse_borders(child(tcPr, "w:tcBorders")) or {}
        for side, at_edge, edge_key, inside_key in (
                ("top", r_idx == 0, "top", "insideH"),
                ("bottom", r_idx == n_rows - 1, "bottom", "insideH"),
                ("left", cell["gridcol"] == 0, "left", "insideV"),
                ("right", last_col, "right", "insideV")):
            desc = cell_borders.get(side)
            if desc is None:
                desc = tbl_borders.get(edge_key) if at_edge else tbl_borders.get(inside_key)
            val, space = border_css_value(desc, self.theme)
            if val:
                css["border-%s" % side] = val

        shd = resolve_shd(parse_shd(child(tcPr, "w:shd")), self.theme)
        if shd is None and r_idx == 0 and first_row.get("shd"):
            shd = first_row["shd"]
        if shd is None and default_shd:
            shd = default_shd
        if shd:
            css["background-color"] = shd

        va = get(child(tcPr, "w:vAlign"), "w:val") if child(tcPr, "w:vAlign") is not None else None
        css["vertical-align"] = VALIGN.get(va, "top")

        # Cell padding: tcMar side -> table tblCellMar side -> Word default.
        base_mar = def_mar or {"top": 0.0, "left": twips_to_px(108),
                               "bottom": 0.0, "right": twips_to_px(108)}
        mar = child(tcPr, "w:tcMar")
        pads = []
        for side_key, tags in (("top", ("w:top",)), ("right", ("w:right", "w:end")),
                               ("bottom", ("w:bottom",)), ("left", ("w:left", "w:start"))):
            v = None
            if mar is not None:
                for a in tags:
                    m = child(mar, a)
                    if m is not None and get(m, "w:type") != "nil":
                        v = twips_to_px(get(m, "w:w"))
                        if v is not None:
                            break
            if v is None:
                v = base_mar.get(side_key, 0.0)
            pads.append(px(v) or "0")
        css["padding"] = " ".join(pads)

        if child(tcPr, "w:noWrap") is not None:
            css["white-space"] = "nowrap"
        td_dir = child(tcPr, "w:textDirection")
        if td_dir is not None and get(td_dir, "w:val") in ("btLr", "tbRlV"):
            css["writing-mode"] = "vertical-rl"

        attrs = ""
        if cell["colspan"] > 1:
            attrs += ' colspan="%d"' % cell["colspan"]
        if cell["rowspan"] > 1:
            attrs += ' rowspan="%d"' % cell["rowspan"]
        content = self.render_block_children(tc)
        return "<td%s%s>%s</td>" % (attrs, style_attr(css), content)

    def _table_style_props(self, style_id):
        out = {"borders": {}, "cell_shd": None, "first_row": {}}
        if not style_id:
            return out
        chain = self.styles._chain(style_id)
        for sid in chain:
            st = self.styles.styles.get(sid)
            if st is None or st.el is None:
                continue
            tblPr = child(st.el, "w:tblPr")
            bd = parse_borders(child(tblPr, "w:tblBorders"))
            if bd:
                out["borders"].update(bd)
            tcPr = child(st.el, "w:tcPr")
            shd = resolve_shd(parse_shd(child(tcPr, "w:shd")), self.theme)
            if shd:
                out["cell_shd"] = shd
            for cond in children(st.el, "w:tblStylePr"):
                if get(cond, "w:type") == "firstRow":
                    ctc = child(cond, "w:tcPr")
                    fs = resolve_shd(parse_shd(child(ctc, "w:shd")), self.theme)
                    if fs:
                        out["first_row"]["shd"] = fs
        return out

    # ==================================================================
    # DrawingML
    # ==================================================================
    def _render_drawing(self, drawing):
        try:
            node = find(drawing, "wp:inline")
            inline = node is not None
            if node is None:
                node = find(drawing, "wp:anchor")
            if node is None:
                return []
            extent = find(node, "wp:extent")
            w = emu_to_px(extent.get("cx")) if extent is not None else None
            h = emu_to_px(extent.get("cy")) if extent is not None else None
            graphic = find(node, "a:graphic")
            gdata = find(graphic, "a:graphicData")
            inner, extra = self._render_graphic(gdata, w, h)
            box = {}
            if w:
                box["width"] = px(w)
            if h:
                box["height"] = px(h)
            box.update(extra or {})
            if inline:
                box["display"] = "inline-block"
                # Default (baseline) alignment: Word rests an inline object's
                # bottom edge on the text baseline; a childless inline-block's
                # baseline is its bottom margin edge, which matches.
                return [("html", "<span%s>%s</span>" % (style_attr(box), inner))]
            # Text-wrapping floats: Word reflows text around wrapSquare/
            # wrapTight shapes. position:absolute cannot reflow, so for the
            # alignment-anchored cases emit a real CSS float that pushes text
            # like Word does. Offset-positioned or behind/in-front shapes stay
            # absolutely positioned (no reflow -- documented approximation).
            wrap_kind = None
            for wtag in ("wp:wrapSquare", "wp:wrapTight", "wp:wrapThrough",
                         "wp:wrapTopAndBottom"):
                if find(node, wtag) is not None:
                    wrap_kind = local(wtag.split(":")[1])
                    break
            behind = str(node.get("behindDoc")) in ("1", "true")
            if wrap_kind and not behind:
                posH = find(node, "wp:positionH")
                halign = None
                if posH is not None:
                    al = find(posH, "wp:align")
                    halign = al.text.strip() if al is not None and al.text else None
                dist = {}
                for a, cssk in (("distL", "margin-left"), ("distR", "margin-right"),
                                ("distT", "margin-top"), ("distB", "margin-bottom")):
                    v = emu_to_px(node.get(a))
                    if v:
                        dist[cssk] = px(v)
                if wrap_kind == "wrapTopAndBottom":
                    box["display"] = "block"
                    box.update(dist)
                    if halign == "center":
                        box["margin-left"] = box["margin-right"] = "auto"
                    elif halign in ("right", "end", "outside"):
                        box["margin-left"] = "auto"
                    return [("html", "<span%s>%s</span>" % (style_attr(box), inner))]
                if halign in ("left", "right", "start", "end", "inside", "outside"):
                    box["float"] = "left" if halign in ("left", "start", "inside") else "right"
                    dist.setdefault("margin-left" if box["float"] == "right" else "margin-right",
                                    px(emu_to_px(114300)))  # Word default 0.125in
                    box.update(dist)
                    return [("html", "<span%s>%s</span>" % (style_attr(box), inner))]
            left, top = self._anchor_pos(node, w or 0, h or 0)
            box["position"] = "absolute"
            box["left"] = px(left)
            box["top"] = px(top)
            if behind:
                # Must be NEGATIVE: a z-index:0 positioned box still paints on
                # top of static in-flow content per CSS painting order, which
                # would hide the whole document behind a background shape.
                box["z-index"] = "-1"
                box["pointer-events"] = "none"
            else:
                box["z-index"] = "5"
            return [("float", "<div%s>%s</div>" % (style_attr(box), inner))]
        except Exception as e:
            return [("html", "<!-- drawing error: %s -->" % html.escape(str(e)))]

    def _render_graphic(self, gdata, w, h):
        if gdata is None:
            return "", {}
        pic = find(gdata, "pic:pic")
        if pic is not None:
            return self._render_pic(pic)
        wsp = find(gdata, "wps:wsp")
        if wsp is not None:
            return self._render_wsp(wsp)
        grp = find(gdata, "wpg:wgp") or find(gdata, "wpg:grpSp")
        if grp is not None:
            return self._render_group(grp, w, h)
        return "", {}

    def _render_pic(self, pic):
        blip = find(pic, "pic:blipFill/a:blip")
        rid = get(blip, "r:embed") or get(blip, "r:link") if blip is not None else None
        uri = self.pkg.data_uri(self.rels, rid)
        extra = {}
        spPr = find(pic, "pic:spPr")
        xfrm = find(spPr, "a:xfrm")
        if xfrm is not None:
            self._apply_xfrm_transform(xfrm, extra)
        # Non-rect crop shape: Word masks the picture with the preset
        # geometry. Ellipse (the common "circle photo") maps to border-radius;
        # rounded rect likewise.
        geom = find(spPr, "a:prstGeom")
        prst = geom.get("prst") if geom is not None else None
        if prst == "ellipse":
            extra["border-radius"] = "50%"
            extra["overflow"] = "hidden"
        elif prst in ("roundRect", "round2SameRect"):
            extra["border-radius"] = "8%"
            extra["overflow"] = "hidden"
        if not uri:
            return "", extra
        rid_ext = self.pkg.media_ext(self.rels, rid)
        if rid_ext in ("emf", "wmf", "tif", "tiff"):
            # Browsers cannot decode EMF/WMF/TIFF. Emit an honest sized
            # placeholder instead of a broken image icon.
            ph = ("<div style=\"width:100%;height:100%;box-sizing:border-box;"
                  "border:1px dashed #b6b6c2;background:"
                  "repeating-linear-gradient(45deg,#f2f2f6 0 8px,#e8e8ee 8px 16px);"
                  "display:flex;align-items:center;justify-content:center;"
                  "color:#8a8a97;font:9px sans-serif;overflow:hidden\">%s</div>"
                  % rid_ext.upper())
            return ph, extra
        # srcRect crop: percentages (of 100000) trimmed off each edge.
        crop = find(pic, "pic:blipFill/a:srcRect")
        img_style = "display:block;width:100%;height:100%;object-fit:fill"
        if crop is not None and any(crop.get(k) for k in ("l", "t", "r", "b")):
            def frac(k):
                try:
                    return max(0.0, float(crop.get(k) or 0) / 100000.0)
                except ValueError:
                    return 0.0
            l, t, r, b = frac("l"), frac("t"), frac("r"), frac("b")
            wf, hf = max(1e-6, 1.0 - l - r), max(1e-6, 1.0 - t - b)
            img_style = ("display:block;position:absolute;width:%.4f%%;height:%.4f%%;"
                         "left:-%.4f%%;top:-%.4f%%" %
                         (100.0 / wf, 100.0 / hf, 100.0 * l / wf, 100.0 * t / hf))
            extra["position"] = extra.get("position", "relative")
            extra["overflow"] = "hidden"
        img = ('<img src="%s" alt="" style="%s">'
               % (html.escape(uri, quote=True), img_style))
        return img, extra

    def _render_wsp(self, wsp):
        spPr = find(wsp, "wps:spPr")
        geom = find(spPr, "a:prstGeom")
        prst = geom.get("prst") if geom is not None else None
        # Connector/line shapes: the "shape" is the stroke itself. A cx x 0
        # extent with a border would paint TWO hairlines (top+bottom border of
        # a zero-height box); render a filled bar of the stroke width instead.
        if prst in ("line", "straightConnector1", "bentConnector3", "curvedConnector3"):
            return self._render_line_shape(spPr)
        xfrm = find(spPr, "a:xfrm")
        ext = find(xfrm, "a:ext") if xfrm is not None else None
        size = None
        if ext is not None:
            size = (emu_to_px(ext.get("cx")), emu_to_px(ext.get("cy")))
        style, svg_clip = self._shape_style(spPr, size)
        extra = {}
        if xfrm is not None:
            self._apply_xfrm_transform(xfrm, extra)
        style.update(extra)
        txbx = find(wsp, "wps:txbx")
        content = find(txbx, "w:txbxContent") if txbx is not None else None
        inner = svg_clip or ""
        if content is not None:
            body = self.render_block_children(content)
            bodyPr = find(wsp, "wps:bodyPr")
            # Text sits above the SVG silhouette.
            tb = self._wrap_textbox(body, bodyPr)
            inner += ('<div style="position:relative;width:100%%;height:100%%">%s</div>' % tb
                      if svg_clip else tb)
        return inner, style

    def _render_line_shape(self, spPr):
        """A DrawingML line/connector -> horizontal/vertical rule or SVG."""
        ln = find(spPr, "a:ln")
        stroke_w = emu_to_px(ln.get("w")) if ln is not None and ln.get("w") else 1.0
        stroke_w = max(stroke_w or 1.0, 0.75)
        color = "#000000"
        if ln is not None:
            lc = find(ln, "a:solidFill")
            got = self._dml_color(list(lc)[0] if lc is not None and len(lc) else None)
            if got:
                color = got
        xfrm = find(spPr, "a:xfrm")
        ext = find(xfrm, "a:ext") if xfrm is not None else None
        cx = emu_to_px(ext.get("cx")) if ext is not None else None
        cy = emu_to_px(ext.get("cy")) if ext is not None else None
        flip_v = xfrm is not None and xfrm.get("flipV") in ("1", "true")
        flip_h = xfrm is not None and xfrm.get("flipH") in ("1", "true")
        if not cy:  # horizontal rule
            # Word centers the stroke on the geometric line, which sits on
            # the text baseline for inline drawings -> shift down half the
            # stroke so the bar straddles the baseline like Word draws it.
            return "", {"background-color": color, "height": px(stroke_w),
                        "vertical-align": px(-stroke_w / 2.0)}
        if not cx:  # vertical rule
            return "", {"background-color": color, "width": px(stroke_w)}
        # Diagonal: inline SVG keeps it self-contained and crisp.
        x1, y1, x2, y2 = 0, 0, cx, cy
        if flip_h:
            x1, x2 = x2, x1
        if flip_v:
            y1, y2 = y2, y1
        svg = ('<svg width="100%%" height="100%%" viewBox="0 0 %g %g" '
               'preserveAspectRatio="none"><line x1="%g" y1="%g" x2="%g" y2="%g" '
               'stroke="%s" stroke-width="%g"/></svg>'
               % (cx, cy, x1, y1, x2, y2, color, stroke_w))
        return svg, {}

    def _render_group(self, grp, w, h):
        # Child shapes live in the group's child coordinate space (chOff/
        # chExt); scale into the group's actual extent.
        gpr = find(grp, "wpg:grpSpPr")
        gx = find(gpr, "a:xfrm") if gpr is not None else None
        ch_off = find(gx, "a:chOff") if gx is not None else None
        ch_ext = find(gx, "a:chExt") if gx is not None else None
        ox = emu_to_px(ch_off.get("x")) if ch_off is not None else 0.0
        oy = emu_to_px(ch_off.get("y")) if ch_off is not None else 0.0
        cw = emu_to_px(ch_ext.get("cx")) if ch_ext is not None else None
        chh = emu_to_px(ch_ext.get("cy")) if ch_ext is not None else None
        sx = (w / cw) if (w and cw) else 1.0
        sy = (h / chh) if (h and chh) else 1.0
        parts = []
        for sp in children(grp):
            lt = local(sp.tag)
            try:
                if lt == "wsp":
                    spPr = find(sp, "wps:spPr")
                    xfrm = find(spPr, "a:xfrm")
                    inner, style = self._render_wsp(sp)
                elif lt == "pic":
                    spPr = find(sp, "pic:spPr")
                    xfrm = find(spPr, "a:xfrm")
                    inner, style = self._render_pic(sp)
                elif lt in ("grpSp", "wgp"):
                    spPr = find(sp, "wpg:grpSpPr")
                    xfrm = find(spPr, "a:xfrm") if spPr is not None else None
                    ext2 = find(xfrm, "a:ext") if xfrm is not None else None
                    w2 = emu_to_px(ext2.get("cx")) if ext2 is not None else None
                    h2 = emu_to_px(ext2.get("cy")) if ext2 is not None else None
                    inner, style = self._render_group(sp, w2, h2)
                else:
                    continue
            except Exception:
                continue
            off = find(xfrm, "a:off") if xfrm is not None else None
            ext = find(xfrm, "a:ext") if xfrm is not None else None
            cell = dict(style)
            cell["position"] = "absolute"
            if off is not None:
                cell["left"] = px(((emu_to_px(off.get("x")) or 0) - (ox or 0)) * sx)
                cell["top"] = px(((emu_to_px(off.get("y")) or 0) - (oy or 0)) * sy)
            if ext is not None:
                cx = (emu_to_px(ext.get("cx")) or 0) * sx
                cy = (emu_to_px(ext.get("cy")) or 0) * sy
                cell["width"] = px(cx)
                # Line shapes carry their own height (stroke width).
                if "height" not in cell or (cy and cy > 0):
                    cell["height"] = px(cy)
            parts.append("<div%s>%s</div>" % (style_attr(cell), inner))
        return '<div style="position:relative;width:100%%;height:100%%">%s</div>' % "".join(parts), {}

    def _apply_xfrm_transform(self, xfrm, extra):
        rot = xfrm.get("rot")
        flip_h = xfrm.get("flipH") in ("1", "true")
        flip_v = xfrm.get("flipV") in ("1", "true")
        tf = []
        if rot:
            try:
                tf.append("rotate(%gdeg)" % (int(rot) / 60000.0))
            except ValueError:
                pass
        if flip_h:
            tf.append("scaleX(-1)")
        if flip_v:
            tf.append("scaleY(-1)")
        if tf:
            extra["transform"] = " ".join(tf)

    def _shape_style(self, spPr, size=None):
        """
        CSS for a shape's fill/outline/geometry.
        `size` = (w_px, h_px) enables SVG rendering of non-box geometry
        (custGeom paths, ellipse, non-trivial presets) so blobs/waves/circles
        in modern templates render as their true silhouette, not a rectangle.
        Returns (style_dict, svg_html_or_None); svg is a clip layer to place
        as the shape's first child.
        """
        style = {}
        svg_clip = None
        if spPr is None:
            return style, None
        solid = find(spPr, "a:solidFill")
        grad = find(spPr, "a:gradFill")
        nofill = find(spPr, "a:noFill")
        blip = find(spPr, "a:blipFill")
        fill_color = fill_grad = fill_uri = None
        if solid is not None:
            fill_color = self._dml_color(list(solid)[0] if len(solid) else None)
        elif grad is not None:
            fill_grad = self._dml_gradient(grad)
        elif blip is not None:
            b = find(blip, "a:blip")
            fill_uri = self.pkg.data_uri(self.rels, get(b, "r:embed")) if b is not None else None
        # Outline
        ln = find(spPr, "a:ln")
        line_w = line_color = None
        if ln is not None and find(ln, "a:noFill") is None:
            lc = find(ln, "a:solidFill")
            line_color = self._dml_color(list(lc)[0] if lc is not None and len(lc) else None)
            if line_color or lc is None:
                line_w = emu_to_px(ln.get("w")) or 1
                line_color = line_color or "#000000"

        geom = find(spPr, "a:prstGeom")
        prst = get_local_attr(geom, "prst") if geom is not None else None
        custom = find(spPr, "a:custGeom")

        # Path-based silhouette (blobs, waves, ellipses, non-box presets).
        d_attr = vb = None
        if size and (custom is not None or prst == "ellipse"
                     or (prst and prst not in ("rect", "roundRect", "round2SameRect"))):
            if custom is not None:
                d_attr, vb = self._custgeom_path(custom)
            elif prst == "ellipse":
                d_attr, vb = self._ellipse_path()
        if d_attr and size:
            svg_fill = fill_color or "none"
            defs = pattern_ref = ""
            if fill_uri:
                pid = "p%d" % (abs(hash(fill_uri)) % 100000)
                defs = ('<defs><pattern id="%s" width="1" height="1" '
                        'patternContentUnits="objectBoundingBox">'
                        '<image href="%s" width="1" height="1" '
                        'preserveAspectRatio="xMidYMid slice"/></pattern></defs>'
                        % (pid, html.escape(fill_uri, quote=True)))
                svg_fill = "url(#%s)" % pid
            elif fill_grad and fill_grad.startswith("linear"):
                svg_fill = fill_color or "#888"  # gradient in SVG omitted; approx
            stroke = (' stroke="%s" stroke-width="%g"' % (line_color, line_w)) if line_w else ""
            svg_clip = ('<svg width="100%%" height="100%%" viewBox="%s" '
                        'preserveAspectRatio="none" style="position:absolute;'
                        'inset:0;display:block">%s<path d="%s" fill="%s"%s/></svg>'
                        % (vb, defs, d_attr, svg_fill, stroke))
            style["position"] = style.get("position", "relative")
            return style, svg_clip

        # Box rendering (rect / roundRect / fallback).
        if fill_color:
            style["background-color"] = fill_color
        elif fill_grad:
            style["background"] = fill_grad
        elif fill_uri:
            style["background-image"] = "url(%s)" % fill_uri
            style["background-size"] = "cover"
            style["background-position"] = "center"
        elif nofill is not None:
            style["background-color"] = "transparent"
        if line_w:
            style["border"] = "%s solid %s" % (px(line_w), line_color)
        if prst in ("roundRect", "round2SameRect"):
            style["border-radius"] = "10px"
        elif prst == "ellipse":
            style["border-radius"] = "50%"
        return style, None

    def _custgeom_path(self, custom):
        """DrawingML <a:custGeom> -> (SVG path d, viewBox). Path coords are in
        the path's own w x h space, so the viewBox matches and preserveAspect
        none scales it onto the shape box."""
        path = find(custom, "a:pathLst/a:path")
        if path is None:
            return None, None
        pw = float(path.get("w") or 1)
        ph = float(path.get("h") or 1)
        cmds = []
        for node in path:
            t = local(node.tag)
            pts = [(float(p.get("x")), float(p.get("y")))
                   for p in children(node, "a:pt")]
            if t == "moveTo" and pts:
                cmds.append("M%g %g" % pts[0])
            elif t == "lnTo" and pts:
                cmds.append("L%g %g" % pts[0])
            elif t == "cubicBezTo" and len(pts) >= 3:
                cmds.append("C%g %g %g %g %g %g" % (pts[0] + pts[1] + pts[2]))
            elif t == "quadBezTo" and len(pts) >= 2:
                cmds.append("Q%g %g %g %g" % (pts[0] + pts[1]))
            elif t == "close":
                cmds.append("Z")
        if not cmds:
            return None, None
        return " ".join(cmds), "0 0 %g %g" % (pw, ph)

    def _ellipse_path(self):
        # Unit ellipse in a 0..100 box (two arcs).
        return ("M0 50 A50 50 0 1 1 100 50 A50 50 0 1 1 0 50 Z", "0 0 100 100")

    def _wrap_textbox(self, body, bodyPr):
        # Default DrawingML text insets (EMU): L/R 0.1in, T/B 0.05in.
        ins = {"lIns": 91440, "tIns": 45720, "rIns": 91440, "bIns": 45720}
        anchor = "top"
        if bodyPr is not None:
            for a in ins:
                v = bodyPr.get(a)
                if v is not None:
                    try:
                        ins[a] = int(v)
                    except ValueError:
                        pass
            anchor = {"t": "top", "ctr": "center", "b": "bottom"}.get(bodyPr.get("anchor"), "top")
        pad = "%s %s %s %s" % (px(emu_to_px(ins["tIns"])), px(emu_to_px(ins["rIns"])),
                               px(emu_to_px(ins["bIns"])), px(emu_to_px(ins["lIns"])))
        style = {"box-sizing": "border-box", "width": "100%", "height": "100%", "padding": pad}
        if anchor in ("center", "bottom"):
            style["display"] = "flex"
            style["flex-direction"] = "column"
            style["justify-content"] = "center" if anchor == "center" else "flex-end"
        return "<div%s>%s</div>" % (style_attr(style), body)

    def _dml_color(self, el):
        if el is None:
            return None
        lt = local(el.tag)
        base = None
        if lt == "srgbClr":
            base = el.get("val")
        elif lt == "sysClr":
            base = el.get("lastClr") or el.get("val")
        elif lt == "schemeClr":
            slot = DML_SCHEME.get(el.get("val"), el.get("val"))
            base = self.theme.color(slot) if self.theme else None
        base = normalize_hex(base)
        if base is None:
            return None
        r, g, b = (c / 255.0 for c in hex_to_rgb(base))
        hh, ll, ss = colorsys.rgb_to_hls(r, g, b)
        alpha = None
        for mod in el:
            m = local(mod.tag)
            raw = mod.get("val")
            f = (float(raw) / 100000.0) if raw not in (None, "") else None
            if f is None:
                continue
            if m == "lumMod":
                ll *= f
            elif m == "lumOff":
                ll += f
            elif m == "shade":
                ll *= f
            elif m == "tint":
                ll = ll * f + (1.0 - f)
            elif m == "satMod":
                ss *= f
            elif m == "alpha":
                alpha = f
        ll = max(0.0, min(1.0, ll))
        ss = max(0.0, min(1.0, ss))
        r, g, b = colorsys.hls_to_rgb(hh, ll, ss)
        if alpha is not None:
            return "rgba(%d,%d,%d,%g)" % (round(r * 255), round(g * 255), round(b * 255), alpha)
        return "#" + rgb_to_hex((r * 255, g * 255, b * 255))

    def _dml_gradient(self, grad):
        stops = []
        gs_lst = find(grad, "a:gsLst")
        for gs in children(gs_lst, "a:gs") if gs_lst is not None else []:
            pos = gs.get("pos")
            col = self._dml_color(list(gs)[0] if len(gs) else None)
            if col:
                p = "%g%%" % (float(pos) / 1000.0) if pos else None
                stops.append(col + (" " + p if p else ""))
        angle = "180deg"
        lin = find(grad, "a:lin")
        if lin is not None and lin.get("ang"):
            try:
                angle = "%gdeg" % ((int(lin.get("ang")) / 60000.0) + 90)
            except ValueError:
                pass
        if len(stops) >= 2:
            return "linear-gradient(%s,%s)" % (angle, ",".join(stops))
        return stops[0] if stops else "transparent"

    def _anchor_pos(self, anchor, w, h):
        return (self._axis(find(anchor, "wp:positionH"), "h", w),
                self._axis(find(anchor, "wp:positionV"), "v", h))

    def _axis(self, node, axis, size):
        g = self.page
        if node is None:
            return 0
        relfrom = node.get("relativeFrom") or ("column" if axis == "h" else "paragraph")
        off = find(node, "wp:posOffset")
        align = find(node, "wp:align")
        if axis == "h":
            base = {"page": 0, "margin": g["ml"], "column": g["ml"], "leftMargin": g["ml"],
                    "rightMargin": g["w"] - g["mr"], "character": g["ml"],
                    "insideMargin": g["ml"], "outsideMargin": 0}.get(relfrom, g["ml"])
            content = g["w"] - g["ml"] - g["mr"]
        else:
            # The vertical base of paragraph/line anchoring is where the
            # anchor paragraph sits. Inside a header that is the header
            # distance (not the body margin) -- header art like full-page
            # background groups is anchored to the header's first paragraph.
            para_base = g["mt"]
            if self.hf_ctx == "header":
                para_base = g["header"]
            elif self.hf_ctx == "footer":
                para_base = g["h"] - g["mb"]
            base = {"page": 0, "margin": g["mt"], "paragraph": para_base,
                    "line": para_base,
                    "topMargin": 0, "bottomMargin": g["h"] - g["mb"],
                    "insideMargin": 0, "outsideMargin": 0}.get(relfrom, para_base)
            content = g["h"] - g["mt"] - g["mb"]
        if off is not None and off.text:
            try:
                return base + emu_to_px(off.text.strip())
            except (ValueError, AttributeError):
                return base
        if align is not None and align.text:
            a = align.text.strip()
            if a in ("left", "start", "inside", "top"):
                return base
            if a in ("right", "end", "outside", "bottom"):
                if relfrom == "page":
                    return (g["w"] if axis == "h" else g["h"]) - size - (0)
                return base + content - size
            if a == "center":
                if relfrom == "page":
                    return ((g["w"] if axis == "h" else g["h"]) - size) / 2.0
                return base + (content - size) / 2.0
        return base

    # ==================================================================
    # VML (legacy shapes / textboxes)
    # ==================================================================
    def _render_vml(self, pict):
        out = []
        for shape in pict:
            lt = local(shape.tag)
            if lt not in ("rect", "roundrect", "oval", "shape", "line", "group"):
                continue
            try:
                out.extend(self._render_vml_shape(shape))
            except Exception as e:
                out.append(("html", "<!-- vml error: %s -->" % html.escape(str(e))))
        return out

    def _render_vml_shape(self, shape):
        vstyle = _parse_vml_style(shape.get("style") or "")
        css = {}
        absolute = vstyle.get("position") == "absolute"
        is_line = local(shape.tag) == "line"
        if is_line:
            # v:line geometry lives in from/to, not the style box.
            try:
                fx, fy = [_len_to_px(v) or 0 for v in (shape.get("from") or "0,0").split(",")[:2]]
                tx, ty = [_len_to_px(v) or 0 for v in (shape.get("to") or "0,0").split(",")[:2]]
            except (ValueError, AttributeError):
                fx = fy = tx = ty = 0
            sw = _len_to_px(shape.get("strokeweight")) or 1.0
            color = _vml_color(shape.get("strokecolor")) if shape.get("strokecolor") else "#000000"
            wpx, hpx = abs(tx - fx), abs(ty - fy)
            css["left"] = px(min(fx, tx))
            css["top"] = px(min(fy, ty))
            if hpx < 0.5:   # horizontal rule
                css["width"] = px(wpx)
                css["height"] = px(max(sw, 0.75))
                css["background-color"] = color
                css["vertical-align"] = px(-sw / 2.0)
            elif wpx < 0.5:  # vertical rule
                css["height"] = px(hpx)
                css["width"] = px(max(sw, 0.75))
                css["background-color"] = color
            else:
                css["width"] = px(wpx)
                css["height"] = px(hpx)
            if absolute:
                css["position"] = "absolute"
            else:
                css["display"] = "inline-block"
                css.pop("left", None)
                css.pop("top", None)
            return [("float" if absolute else "html",
                     "<div%s></div>" % style_attr(css))]
        for k in ("left", "top", "width", "height"):
            v = _len_to_px(vstyle.get(k))
            if v is not None:
                css[k] = px(v)
        # position relative to page vs margin
        if absolute:
            g = self.page
            hrel = vstyle.get("mso-position-horizontal-relative")
            vrel = vstyle.get("mso-position-vertical-relative")
            if hrel in ("margin", "left-margin-area", "text") and "left" in css:
                css["left"] = px(_len_to_px(vstyle.get("left")) + g["ml"])
            if vrel in ("margin", "top-margin-area", "text") and "top" in css:
                css["top"] = px(_len_to_px(vstyle.get("top")) + g["mt"])
            css["position"] = "absolute"
            zi = (vstyle.get("z-index") or "").strip()
            css["z-index"] = "-1" if zi.startswith("-") else "5"
            if zi.startswith("-"):
                css["pointer-events"] = "none"
        else:
            css["display"] = "inline-block"

        fill = shape.get("fillcolor")
        if shape.get("filled") == "f":
            css["background-color"] = "transparent"
        elif fill:
            css["background-color"] = _vml_color(fill)
        stroke = shape.get("strokecolor")
        stroked = shape.get("stroked")
        if stroked != "f":
            sw = shape.get("strokeweight")
            w = _len_to_px(sw) if sw else 1
            css["border"] = "%s solid %s" % (px(w or 1), _vml_color(stroke) if stroke else "#000000")
        if local(shape.tag) == "oval":
            css["border-radius"] = "50%"
        elif local(shape.tag) == "roundrect":
            css["border-radius"] = "8px"
        rot = vstyle.get("rotation")
        if rot:
            try:
                css["transform"] = "rotate(%gdeg)" % float(re.sub("[^0-9.\\-]", "", rot))
            except ValueError:
                pass

        imgdata = find(shape, "v:imagedata")
        inner = ""
        if imgdata is not None:
            rid = get(imgdata, "r:id") or get(imgdata, "o:relid")
            uri = self.pkg.data_uri(self.rels, rid)
            if uri:
                inner = ('<img src="%s" alt="" style="width:100%%;height:100%%;object-fit:fill">'
                         % html.escape(uri, quote=True))
        txbx = find(shape, "v:textbox")
        content = find(txbx, "w:txbxContent") if txbx is not None else None
        if content is not None:
            body = self.render_block_children(content)
            inner += self._wrap_textbox(body, None)
        css.setdefault("box-sizing", "border-box")
        tag_html = "<div%s>%s</div>" % (style_attr(css), inner)
        return [("float" if absolute else "html", tag_html)]


def get_local_attr(el, name):
    if el is None:
        return None
    return el.get(name)


def _parse_vml_style(s):
    out = {}
    for part in s.split(";"):
        if ":" in part:
            k, v = part.split(":", 1)
            out[k.strip().lower()] = v.strip()
    return out


def _vml_color(c):
    if not c:
        return "transparent"
    c = c.strip()
    if c.startswith("#"):
        return c
    # VML sometimes uses "#rrggbb" or named or "windowText"
    if re.match(r"^[0-9A-Fa-f]{6}$", c):
        return "#" + c
    return c
