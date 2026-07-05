# fancydocx

Convert fancy, design-heavy **.docx** files (resumes, cover letters, posters, reports) into a **single self-contained .html** — CSS, images, and fonts all inlined. Built as a from-scratch OOXML parser in **pure Python standard library**: no LibreOffice, no Word, no headless browser, no network, no pip dependencies.

The goal is 1:1 visual parity with Microsoft Word's rendering, for batch conversion of large corpora.

---

## Install

```bash
pip install fancydocx
```

Requires **Python 3.8+**. The converter has **zero runtime dependencies** — standard library only. (The optional `verify.py` measurement harness may use `Pillow`/`numpy`/`pywin32`; the package never imports them. Install with `pip install "fancydocx[verify]"`.)

---

## Quick start

One line of Python:
```python
import fancydocx

fancydocx.convert("resume.docx", "resume.html")   # write the file
html = fancydocx.convert("resume.docx")            # or get the HTML string
```

From the command line (installed as the `fancydocx` command):
```bash
fancydocx resume.docx                     # -> resume.html
fancydocx resume.docx -o out/resume.html
fancydocx ./docs -o ./html --workers 8    # whole folder, recursively
```

Other API entry points: `convert_docx(path) -> str`, `convert_file(src, dst) -> path`, and the `DocxPackage` class for low-level access.

Running from a source checkout without installing: `py -3.12 convert.py resume.docx -o out.html`.

### Options
| Flag | Meaning |
|---|---|
| `-o, --output` | Output `.html` (single) or output folder (batch) |
| `--glob "*.docx"` | Which files to pick up in batch mode |
| `--workers N` | Parallel worker processes for batch (default 1) |
| `--embed-fonts` | Also inline every referenced font family found on this machine as `@font-face` (see **Fonts**). Opt-in — it makes files large. |
| `--no-headers` | Skip header/footer rendering |
| `--quiet` | Batch: print only the final summary |

A corrupt/unreadable file in a batch is isolated and logged; it never aborts the run.

### Python API
```python
from fancydocx import convert_docx, convert_file

html = convert_docx("resume.docx", embed_fonts=False)   # -> str
convert_file("resume.docx", "out/resume.html")           # -> path
```

---

## How it works

Each `<w:sectPr>` section becomes a fixed-size `.docx-page` box (page size, padding = margins), so the result looks like Word's print view and prints to correct pages. Normal-flow content sits in the padding area; floating shapes/images are absolutely positioned against the page box using coordinates computed from the page origin (matching Word's page/margin anchoring).

Formatting is resolved as **data, not CSS strings** — every `rPr`/`pPr` is parsed into a normalized dict, the full cascade is deep-merged (`docDefaults → paragraph-style basedOn chain → direct`, plus the character-style chain for runs), and only then converted to CSS. That is what makes inheritance correct.

| Module | Responsibility |
|---|---|
| `core.py` | Namespaces, element helpers, exact unit math (twips/EMU/half-pt/eighth-pt → px/pt) |
| `color.py` | Hex/RGB/HSL, Office luminance tint/shade, theme color resolution |
| `theme.py` | Color scheme + font scheme, honoring `clrSchemeMapping` |
| `package.py` | Zip/relationships, base64 media, **embedded-font de-obfuscation → @font-face** |
| `styles.py` | Property model, full style cascade, CSS generation |
| `numbering.py` | List markers, counters, bullet-glyph mapping |
| `fontmetrics.py` | Reads real `OS/2`/`hhea` tables for Word's true line height + Word-style font substitution + opt-in local-font embedding |
| `render.py` | The walker: sections→pages, tables (vMerge/gridSpan grid), DrawingML + VML shapes/text boxes, `custGeom`→SVG, images |
| `__init__.py` | `convert_docx`/`convert_file`, base CSS, document assembly |

### What is covered
Page geometry & margins; sections & columns; the full style cascade; theme + explicit colors with tint/shade; fonts with metric-correct line height; bold/italic/underline/strike/caps/smallcaps/super-sub/highlight/letter-spacing; paragraph alignment/indent/spacing/shading/borders; tables (column grid, `gridSpan`/`vMerge`, per-cell borders/shading/margins/valign, pct widths); bullet & numbered lists; inline and floating images (with ellipse/roundRect crops); DrawingML shapes & text boxes; VML shapes/text boxes/images; `custGeom` freeform shapes → SVG silhouettes (blobs/waves); rotation/flip; hyperlinks; dot-leader tab stops; headers/footers; automatic pagination via Word's `lastRenderedPageBreak`; document/page background.

---

## Fonts & fidelity (read this)

Font availability is the single biggest factor in true parity.

- Line height is computed from each font's **real design metrics** (`OS/2`/`hhea`), so paragraphs keep Word's geometry even when the browser substitutes a family. Word's "single" spacing is the font's design height (Calibri ≈ 1.22, Aptos ≈ 1.28, Century Gothic ≈ 1.59), not `1.0` or `1.15`.
- When a font isn't installed, the CSS fallback follows **Word's own substitution** (Aptos→Calibri, Futura→Century Gothic, …) rather than the browser's default sans.
- Fonts that are **embedded inside the .docx** are recovered automatically (de-obfuscated) and inlined — no flag needed.
- Fonts that are only **installed on your machine** are inlined only with `--embed-fonts`. This is opt-in because whole-face embedding is heavy (a resume using Segoe UI can grow to several MB). For a corpus, prefer having the fonts installed on the viewing machine over embedding into every file.

For exact rendering: view the HTML on a machine that has the document's fonts, or use `--embed-fonts` for portability.

---

## Verifying fidelity — `verify.py`

`verify.py` is separate tooling (not part of the converter) that measures output two ways:

```bash
py -3.12 verify.py resume.docx --html out/resume.html --browser <path-to-gstack-browse>
```

1. **OOXML conformance (always available, offline):** parses the `.docx` and asserts the HTML encodes the exact intended geometry — page size, margins, column widths, font sizes, colors, border widths, image extents, float offsets.
2. **Visual ground truth:** rasterizes Word's *own* render and diffs it against a screenshot of the HTML at identical dimensions. Ground truth comes from Word COM (if Word is installed) or, failing that, from **`docProps/thumbnail.emf`** — the page-1 preview Word bakes into the file on save. Outputs a fidelity score, a side-by-side, and a difference heatmap.

Note on the score: two different rasterizers (Word's GDI vs a browser) never produce pixel-identical text even for a perfect layout, and any font missing on the comparison machine inflates the difference. Read the heatmap, not just the number — diffuse noise over text is antialiasing/substitution; concentrated blocks or displaced edges are real layout bugs.

---

## Limitations

- **Fonts** must be present (installed or `--embed-fonts`) for exact metrics; otherwise Word-style substitutes are used.
- **EMF/WMF** vector images can't display in browsers; they render as sized placeholders.
- **Automatic pagination** is reproduced only where Word recorded a `lastRenderedPageBreak`; there is no independent line-breaking/pagination engine.
- **Text wrap around floats** is approximate; there is no full reflow engine.
- **SmartArt/WordArt/charts** render via their fallback image when present, else are skipped.
- Absolute pixel-identity to Word is not a goal a browser can reach; near-indistinguishable visual parity is.

---

## Publishing

Build the distribution artifacts (wheel + source dist):
```bash
pip install build
python -m build          # -> dist/fancydocx-<version>-py3-none-any.whl + .tar.gz
```

Upload to PyPI with [twine](https://twine.readthedocs.io/):
```bash
pip install twine
twine upload dist/*      # use TestPyPI first: twine upload -r testpypi dist/*
```

Before publishing, edit `pyproject.toml`: set the real author name and the
`YOUR-GITHUB-USERNAME` URLs, and bump `__version__` in `fancydocx/__init__.py`
(the version is single-sourced from there). Confirm the name `fancydocx` is
available on PyPI, or choose another `name` in `pyproject.toml`.

## Project layout
```
docxtohtml/
  pyproject.toml        packaging metadata (setuptools backend, PEP 621)
  LICENSE               MIT
  README.md
  fancydocx/            the pure-stdlib converter package (this is what ships)
    __init__.py         convert() / convert_docx() / convert_file()
    cli.py              the `fancydocx` command
    __main__.py         `python -m fancydocx`
    core, color, theme, package, styles, numbering, fontmetrics, render
  convert.py            dev shim to run the CLI from a source checkout
  verify.py             measurement harness (optional deps: Pillow/numpy/pywin32)
  make_sample_docx.py   synthetic feature-rich .docx generator
```
