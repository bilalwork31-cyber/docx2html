fancydocx
=========

Convert fancy, design-heavy Word (.docx) files into a single, self-contained
HTML file, with all styling, images and fonts inlined. It is a from-scratch
reader for the Office Open XML format, written in pure Python. No LibreOffice,
no Word, and no third-party dependencies.

Available on PyPI: https://pypi.org/project/fancydocx/0.1.0/


Features
--------

- Tables, columns, shapes, text boxes and floating images
- Theme colours, fonts and the full paragraph/run style cascade
- Bullet and numbered lists, hyperlinks, headers and footers
- Embedded fonts are recovered and inlined automatically
- Batch conversion from a simple command-line tool


Installation
------------

    pip install fancydocx


Usage
-----

In Python:

    import fancydocx

    fancydocx.convert("resume.docx", "resume.html")   # write a file
    html = fancydocx.convert("resume.docx")           # or return a string

From the command line:

    fancydocx resume.docx -o resume.html
    fancydocx ./documents -o ./html --workers 8       # convert a folder


Requirements
------------

Python 3.8 or newer. The library uses only the standard library.


Notes
-----

- Text renders with the document's own fonts when they are installed on the
  viewer's machine; pass --embed-fonts to inline them for portability.
- EMF/WMF vector images are shown as placeholders, as browsers cannot
  display them.


License
-------

MIT. See the LICENSE file.
