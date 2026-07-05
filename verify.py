#!/usr/bin/env python3
"""
verify.py -- measurement harness for fancydocx. NOT part of the converter;
this file may use third-party deps (Pillow, numpy, pymupdf, pywin32) and
external tools (Word COM, the gstack headless browser). The converter package
(fancydocx/) must never import anything from here.

Two kinds of ground truth:

(A) OOXML conformance: parse the .docx ourselves and assert the emitted HTML
    encodes the exact intended geometry (page size, margins, column widths,
    font sizes, colors, border widths, image extents, float offsets).
    Catches unit-math and mapping bugs precisely, fully offline.

(B) Visual: compare a raster of Word's OWN rendering against a screenshot of
    our HTML at identical pixel dimensions.
      Source 1: Word COM automation -> PDF -> raster    (if Word installed)
      Source 2: docProps/thumbnail.emf inside the .docx (Word drew it when
                the file was saved; page 1 only)
    If neither exists we say so and skip (B).

Usage:
    py -3.12 verify.py <file.docx> [--html out.html] [--outdir out\\verify]
                       [--browser <path-to-gstack-browse>] [--skip-visual]
"""
from __future__ import annotations
import argparse
import io
import os
import re
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
DEFAULT_BROWSER = Path.home() / ".claude" / "skills" / "gstack" / "browse" / "dist" / "browse"


# ---------------------------------------------------------------------------
# (A) conformance checks
# ---------------------------------------------------------------------------
def _tw2px(v):
    return float(v) / 15.0


def conformance_checks(docx_path, html_text):
    """Yield (ok, label, expected, found) tuples."""
    z = zipfile.ZipFile(docx_path)
    doc = ET.fromstring(z.read("word/document.xml"))

    def styleval(pattern):
        m = re.search(pattern, html_text)
        return m.group(1) if m else None

    # page geometry ------------------------------------------------------
    sect = doc.find(".//%ssectPr" % W)
    if sect is not None:
        pgsz = sect.find("%spgSz" % W)
        pgmar = sect.find("%spgMar" % W)
        page_style = styleval(r'class="docx-page" style="([^"]*)"')
        if pgsz is not None and page_style:
            for attr, css in (("w", "width"), ("h", "min-height")):
                want = _tw2px(pgsz.get(W + attr))
                m = re.search(r"%s:([\d.]+)px" % css, page_style)
                got = float(m.group(1)) if m else None
                yield (got is not None and abs(got - want) < 0.06,
                       "page %s" % css, "%.2fpx" % want, "%s" % got)
        if pgmar is not None and page_style:
            m = re.search(r"padding:([\d.]+)px ([\d.]+)px ([\d.]+)px ([\d.]+)px", page_style)
            if m:
                got = tuple(float(x) for x in m.groups())
                want = tuple(_tw2px(pgmar.get(W + k)) for k in ("top", "right", "bottom", "left"))
                ok = all(abs(a - b) < 0.06 for a, b in zip(got, want))
                yield (ok, "page margins (t r b l)",
                       " ".join("%.2f" % x for x in want),
                       " ".join("%.2f" % x for x in got))
            else:
                yield (False, "page margins", "padding shorthand", "missing")

    # table grids ----------------------------------------------------------
    grids = doc.findall(".//%stbl/%stblGrid" % (W, W))
    html_colgroups = re.findall(r"<colgroup>?(?:<col[^>]*>)+", html_text)
    col_runs = re.findall(r"((?:<col style=\"width:[\d.]+px\">)+)", html_text)
    for t_idx, grid in enumerate(grids):
        want = [_tw2px(gc.get(W + "w")) for gc in grid.findall(W + "gridCol")]
        if t_idx < len(col_runs):
            got = [float(x) for x in re.findall(r"width:([\d.]+)px", col_runs[t_idx])]
            ok = len(got) == len(want) and all(abs(a - b) < 0.06 for a, b in zip(got, want))
            yield (ok, "table %d col widths" % t_idx,
                   ",".join("%.1f" % x for x in want),
                   ",".join("%.1f" % x for x in got))
        else:
            yield (False, "table %d col widths" % t_idx, "%d cols" % len(want), "table missing")

    # Font sizes / colors: only assert values that style VISIBLE text. An
    # rPr that decorates an empty paragraph mark or a vMerged-away spacer
    # never reaches the output (correctly), so checking it would be a false
    # alarm. We gather rPr from runs that contain non-empty <w:t>.
    vis_sizes, vis_colors = set(), set()

    def run_has_text(r):
        return any((t.text or "").strip() for t in r.findall(W + "t"))

    for r in doc.iter(W + "r"):
        if not run_has_text(r):
            continue
        rpr = r.find(W + "rPr")
        if rpr is None:
            continue
        sz = rpr.find(W + "sz")
        if sz is not None and (sz.get(W + "val") or "").isdigit():
            vis_sizes.add(int(sz.get(W + "val")))
        col = rpr.find(W + "color")
        if col is not None:
            v = col.get(W + "val")
            if v and re.fullmatch(r"[0-9A-Fa-f]{6}", v):
                vis_colors.add(v.upper())
    for hp in sorted(vis_sizes):
        token = "font-size:%gpt" % (hp / 2.0)
        yield (token in html_text, "font size %shp" % hp, token,
               "present" if token in html_text else "absent")
    for c in sorted(vis_colors):
        tok = "#" + c
        yield (tok in html_text.upper(), "color %s" % c, tok,
               "present" if tok in html_text.upper() else "absent")

    # image extents (EMU -> px)
    NS_WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    for idx, ext in enumerate(doc.iter(NS_WP + "extent")):
        cx, cy = float(ext.get("cx")) / 9525.0, float(ext.get("cy")) / 9525.0
        wtok = "width:%s" % _fmt_px(cx)
        ok = wtok in html_text
        if cy:  # zero-height (line) extents legitimately override height
            htok = "height:%s" % _fmt_px(cy)
            ok = ok and htok in html_text
        yield (ok, "drawing %d extent" % idx, "%.1fx%.1f" % (cx, cy),
               "present" if ok else "absent")
    z.close()


def _fmt_px(v):
    v = round(v, 2)
    if v == int(v):
        return "%dpx" % int(v)
    return ("%.2f" % v).rstrip("0").rstrip(".") + "px"


# ---------------------------------------------------------------------------
# (B) visual ground truth
# ---------------------------------------------------------------------------
def word_com_pdf(docx_path, out_pdf):
    """Export via Word COM. Returns True on success."""
    try:
        import win32com.client  # noqa
    except ImportError:
        return False
    try:
        word = None
        import win32com.client as wc
        word = wc.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(Path(docx_path).resolve()), ReadOnly=True)
        doc.ExportAsFixedFormat(str(Path(out_pdf).resolve()), 17)  # wdExportFormatPDF
        doc.Close(False)
        word.Quit()
        return Path(out_pdf).exists()
    except Exception as e:
        print("  Word COM unavailable: %s" % e)
        try:
            if word is not None:
                word.Quit()
        except Exception:
            pass
        return False


def pdf_page_png(pdf_path, page_idx, out_png, dpi=192):
    import fitz
    doc = fitz.open(pdf_path)
    page = doc[page_idx]
    pm = page.get_pixmap(dpi=dpi)
    pm.save(out_png)
    doc.close()


def thumbnail_png(docx_path, out_png, dpi=192):
    """Rasterize docProps/thumbnail.emf|wmf via Pillow (Windows GDI)."""
    from PIL import Image
    z = zipfile.ZipFile(docx_path)
    name = None
    for n in z.namelist():
        if n.lower().startswith("docprops/thumbnail."):
            name = n
            break
    if name is None:
        return False
    data = z.read(name)
    z.close()
    if name.lower().endswith((".emf", ".wmf")):
        with tempfile.NamedTemporaryFile(suffix=Path(name).suffix, delete=False) as f:
            f.write(data)
            tmp = f.name
        try:
            im = Image.open(tmp)
            im.load(dpi=dpi)
        finally:
            os.unlink(tmp)
    else:
        im = Image.open(io.BytesIO(data)).convert("RGB")
    im.convert("RGB").save(out_png)
    return True


def browser_page_png(browser, html_path, out_png, page_w, page_h, scale=2):
    """Screenshot the first .docx-page at page_w x page_h CSS px, x`scale`."""
    url = "file:///" + str(Path(html_path).resolve()).replace("\\", "/")
    cmds = [
        [str(browser), "viewport", "%dx%d" % (page_w + 80, page_h + 80), "--scale", str(scale)],
        [str(browser), "goto", url],
        [str(browser), "screenshot", "--selector", ".docx-page", str(out_png)],
    ]
    for c in cmds:
        r = subprocess.run(c, capture_output=True, text=True, timeout=120)
        if r.returncode != 0:
            print("  browser step failed: %s\n  %s" % (" ".join(c[1:3]), r.stderr.strip() or r.stdout.strip()))
            return False
    return Path(out_png).exists()


def compare_images(gt_png, ours_png, out_prefix):
    """Layout-level comparison, tolerant of anti-aliasing differences."""
    import numpy as np
    from PIL import Image, ImageChops

    a = Image.open(gt_png).convert("L")
    b = Image.open(ours_png).convert("L")
    if a.size != b.size:
        b = b.resize(a.size, Image.LANCZOS)
    # quarter-resolution comparison: structural, AA-insensitive
    q = (max(1, a.size[0] // 4), max(1, a.size[1] // 4))
    aq = np.asarray(a.resize(q, Image.BOX), dtype=np.int16)
    bq = np.asarray(b.resize(q, Image.BOX), dtype=np.int16)
    diff = np.abs(aq - bq)
    mean_abs = float(diff.mean())
    pct_off = float((diff > 48).mean() * 100.0)
    score = max(0.0, 100.0 - pct_off * 2 - mean_abs / 2)

    # global vertical drift via row-ink cross-correlation
    arow = 255.0 - aq.mean(axis=1)
    brow = 255.0 - bq.mean(axis=1)
    best_shift, best_err = 0, None
    for s in range(-24, 25):
        if s >= 0:
            err = float(np.abs(arow[s:] - brow[:len(brow) - s]).mean()) if s < len(arow) else 1e9
        else:
            err = float(np.abs(arow[:s] - brow[-s:]).mean())
        if best_err is None or err < best_err:
            best_err, best_shift = err, s
    v_drift_px = best_shift * 4 / 2.0  # quarter-res -> full-res -> CSS px (x2 scale)

    # worst 6 regions on an 8x10 grid
    gh, gw = 10, 8
    H, Wd = diff.shape
    regions = []
    for gy in range(gh):
        for gx in range(gw):
            sl = diff[gy * H // gh:(gy + 1) * H // gh, gx * Wd // gw:(gx + 1) * Wd // gw]
            regions.append((float(sl.mean()), gy, gx))
    regions.sort(reverse=True)

    # artifacts: side-by-side + heatmap
    rgb_a = Image.open(gt_png).convert("RGB")
    rgb_b = Image.open(ours_png).convert("RGB").resize(rgb_a.size, Image.LANCZOS)
    side = Image.new("RGB", (rgb_a.width * 2 + 8, rgb_a.height), (40, 40, 40))
    side.paste(rgb_a, (0, 0))
    side.paste(rgb_b, (rgb_a.width + 8, 0))
    side.save(out_prefix + "_side.png")
    heat = ImageChops.difference(rgb_a, rgb_b).convert("L").point(lambda v: min(255, v * 3))
    heat.save(out_prefix + "_heat.png")

    return {
        "mean_abs": mean_abs, "pct_off": pct_off, "score": score,
        "v_drift_css_px": v_drift_px,
        "worst": [(round(m, 1), "row %d/%d col %d/%d" % (gy + 1, gh, gx + 1, gw))
                  for m, gy, gx in regions[:6]],
        "side": out_prefix + "_side.png", "heat": out_prefix + "_heat.png",
    }


# ---------------------------------------------------------------------------
def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("docx")
    ap.add_argument("--html", help="existing HTML (else converts fresh)")
    ap.add_argument("--outdir", default=str(HERE / "out" / "verify"))
    ap.add_argument("--browser", default=str(DEFAULT_BROWSER))
    ap.add_argument("--skip-visual", action="store_true")
    ap.add_argument("--embed-fonts", action="store_true",
                    help="convert with --embed-fonts before verifying")
    args = ap.parse_args(argv)

    docx_path = Path(args.docx)
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    stem = docx_path.stem

    if args.html:
        html_path = Path(args.html)
        html_text = html_path.read_text(encoding="utf-8")
    else:
        from fancydocx import convert_docx
        html_text = convert_docx(str(docx_path), embed_fonts=args.embed_fonts)
        html_path = outdir / (stem + ".html")
        html_path.write_text(html_text, encoding="utf-8")

    print("== (A) OOXML conformance: %s" % docx_path.name)
    n_ok = n_bad = 0
    failures = []
    for ok, label, want, got in conformance_checks(docx_path, html_text):
        if ok:
            n_ok += 1
        else:
            n_bad += 1
            failures.append("   FAIL %-28s want=%s got=%s" % (label, want, got))
    print("   %d checks passed, %d failed" % (n_ok, n_bad))
    for f in failures[:20]:
        print(f)
    if len(failures) > 20:
        print("   ... %d more" % (len(failures) - 20))

    if args.skip_visual:
        return 0

    print("== (B) visual ground truth")
    m = re.search(r'class="docx-page" style="width:([\d.]+)px;min-height:([\d.]+)px', html_text)
    page_w, page_h = (int(float(m.group(1))), int(float(m.group(2)))) if m else (816, 1056)

    gt_png = outdir / (stem + "_gt.png")
    got_gt = False
    pdf = outdir / (stem + "_word.pdf")
    if word_com_pdf(docx_path, pdf):
        pdf_page_png(str(pdf), 0, str(gt_png), dpi=192)
        got_gt = True
        print("   ground truth: Word COM -> PDF page 1 @192dpi")
    elif thumbnail_png(docx_path, str(gt_png), dpi=192):
        got_gt = True
        print("   ground truth: docProps thumbnail (Word's own save-time render)")
    else:
        print("   no Word COM and no thumbnail in package -> visual check skipped")

    if not got_gt:
        return 0

    ours_png = outdir / (stem + "_ours.png")
    if not browser_page_png(args.browser, html_path, ours_png, page_w, page_h, scale=2):
        return 1
    res = compare_images(str(gt_png), str(ours_png), str(outdir / stem))
    print("   fidelity score : %.1f / 100" % res["score"])
    print("   mean |delta|   : %.2f (0-255, quarter-res grayscale)" % res["mean_abs"])
    print("   pixels off>48  : %.2f%%" % res["pct_off"])
    print("   vertical drift : %+.1f css px (row-ink correlation)" % res["v_drift_css_px"])
    print("   worst regions  :")
    for mval, where in res["worst"]:
        print("      %-18s mean|d|=%s" % (where, mval))
    print("   artifacts: %s | %s" % (res["side"], res["heat"]))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
