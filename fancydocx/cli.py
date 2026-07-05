"""
Command-line interface for fancydocx, exposed as the ``fancydocx`` command
(and ``python -m fancydocx``).

  Single file:
      fancydocx resume.docx                 -> resume.html (next to input)
      fancydocx resume.docx -o out.html

  Whole folder (recursive), mirroring the tree into an output dir:
      fancydocx ./docs -o ./html
      fancydocx ./docs -o ./html --workers 8
"""
from __future__ import annotations
import argparse
import concurrent.futures as cf
import sys
import time
import traceback
from pathlib import Path

from . import __version__, convert_file


def _iter_docx(root, pattern):
    for p in sorted(Path(root).rglob(pattern)):
        # Skip Word lock/temp files like ~$name.docx
        if p.name.startswith("~$"):
            continue
        if p.is_file():
            yield p


def _one(in_path, out_path, include_headers, embed_fonts=False):
    t0 = time.perf_counter()
    try:
        convert_file(in_path, out_path, include_headers=include_headers,
                     embed_fonts=embed_fonts)
        return (in_path, out_path, None, time.perf_counter() - t0)
    except Exception as e:
        return (in_path, out_path,
                "".join(traceback.format_exception_only(type(e), e)).strip(),
                time.perf_counter() - t0)


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="fancydocx",
        description="Convert fancy .docx files to a single self-contained HTML file.")
    ap.add_argument("input", help="A .docx file or a folder containing .docx files")
    ap.add_argument("-o", "--output", help="Output .html file (single) or output folder (batch)")
    ap.add_argument("--glob", default="*.docx", help="Glob for batch mode (default: *.docx)")
    ap.add_argument("--workers", type=int, default=1,
                    help="Parallel worker processes for batch mode (default: 1)")
    ap.add_argument("--no-headers", action="store_true", help="Skip header/footer rendering")
    ap.add_argument("--embed-fonts", action="store_true",
                    help="Inline locally-installed referenced fonts as @font-face "
                         "(exact metrics everywhere, but several MB per file)")
    ap.add_argument("--quiet", action="store_true", help="Only print a final summary")
    ap.add_argument("--version", action="version", version="fancydocx %s" % __version__)
    args = ap.parse_args(argv)

    include_headers = not args.no_headers
    inp = Path(args.input)
    if not inp.exists():
        ap.error("input not found: %s" % inp)

    # ---- single file -------------------------------------------------
    if inp.is_file():
        out = Path(args.output) if args.output else inp.with_suffix(".html")
        in_p, out_p, err, dt = _one(inp, out, include_headers, args.embed_fonts)
        if err:
            print("FAILED %s\n  %s" % (in_p, err), file=sys.stderr)
            return 1
        print("OK  %s -> %s  (%.2fs)" % (in_p, out_p, dt))
        return 0

    # ---- batch folder ------------------------------------------------
    out_dir = Path(args.output) if args.output else inp / "_html"
    files = list(_iter_docx(inp, args.glob))
    if not files:
        print("No files matching %r under %s" % (args.glob, inp))
        return 0

    jobs = [(f, out_dir / f.relative_to(inp).with_suffix(".html")) for f in files]
    ok = fail = 0
    total = len(jobs)
    started = time.perf_counter()
    print("Converting %d file(s) -> %s  (workers=%d)" % (total, out_dir, args.workers))

    def report(res, i):
        nonlocal ok, fail
        in_p, out_p, err, dt = res
        if err:
            fail += 1
            print("[%d/%d] FAILED %s\n    %s" % (i, total, in_p, err), file=sys.stderr)
        else:
            ok += 1
            if not args.quiet:
                print("[%d/%d] %s -> %s  (%.2fs)" % (i, total, in_p.name, out_p, dt))

    if args.workers > 1:
        with cf.ProcessPoolExecutor(max_workers=args.workers) as ex:
            futs = {ex.submit(_one, f, o, include_headers, args.embed_fonts): idx
                    for idx, (f, o) in enumerate(jobs, 1)}
            for fut in cf.as_completed(futs):
                report(fut.result(), futs[fut])
    else:
        for idx, (f, o) in enumerate(jobs, 1):
            report(_one(f, o, include_headers, args.embed_fonts), idx)

    print("\nDone: %d ok, %d failed, %d total in %.1fs"
          % (ok, fail, total, time.perf_counter() - started))
    return 1 if fail else 0


if __name__ == "__main__":
    raise SystemExit(main())
