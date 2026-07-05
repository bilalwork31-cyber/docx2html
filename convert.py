#!/usr/bin/env python3
"""
Dev convenience shim: run the CLI from a source checkout without installing.

    py -3.12 convert.py resume.docx -o out.html

The real CLI lives in ``fancydocx/cli.py`` and is installed as the ``fancydocx``
command (and ``python -m fancydocx``) when you ``pip install`` the package.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fancydocx.cli import main  # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
