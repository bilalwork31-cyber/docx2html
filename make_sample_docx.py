#!/usr/bin/env python3
"""
Generate a feature-rich sample .docx to exercise the converter:
theme colors, a two-column table with a shaded sidebar + vMerge rowspan,
a bullet list, a heading with a bottom border, an inline image, a
right-aligned dot-leader tab stop, and a floating VML banner.

    py make_sample_docx.py            -> sample.docx
"""
import struct
import sys
import zlib
import zipfile
from pathlib import Path


def make_png(w, h, rgb):
    def chunk(typ, data):
        c = typ + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xffffffff)
    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)  # 8-bit truecolor
    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
    return sig + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
 <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 <Default Extension="xml" ContentType="application/xml"/>
 <Default Extension="png" ContentType="image/png"/>
 <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
 <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
 <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
 <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
 <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
 <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
</Relationships>"""

THEME = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
 <a:themeElements>
  <a:clrScheme name="Office">
   <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
   <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
   <a:dk2><a:srgbClr val="44546A"/></a:dk2>
   <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
   <a:accent1><a:srgbClr val="4472C4"/></a:accent1>
   <a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
   <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
   <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
   <a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
   <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
   <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
   <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
  </a:clrScheme>
  <a:fontScheme name="Office">
   <a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>
   <a:minorFont><a:latin typeface="Calibri"/></a:minorFont>
  </a:fontScheme>
  <a:fmtScheme name="Office"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
 </a:themeElements>
</a:theme>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:docDefaults>
  <w:rPrDefault><w:rPr>
    <w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/>
    <w:sz w:val="22"/><w:color w:val="404040"/>
  </w:rPr></w:rPrDefault>
  <w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="264" w:lineRule="auto"/></w:pPr></w:pPrDefault>
 </w:docDefaults>
 <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
 <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/>
   <w:basedOn w:val="Normal"/>
   <w:pPr><w:spacing w:before="240" w:after="80"/>
     <w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="4472C4"/></w:pBdr></w:pPr>
   <w:rPr><w:b/><w:caps/><w:color w:val="44546A"/><w:sz w:val="30"/><w:spacing w:val="20"/></w:rPr>
 </w:style>
 <w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/>
   <w:basedOn w:val="Normal"/><w:pPr><w:ind w:left="600" w:hanging="300"/></w:pPr>
 </w:style>
 <w:style w:type="character" w:styleId="SidebarText"><w:name w:val="Sidebar Text"/>
   <w:rPr><w:color w:val="FFFFFF"/></w:rPr>
 </w:style>
</w:styles>"""

NUMBERING = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
 <w:abstractNum w:abstractNumId="0">
  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#61623;"/>
    <w:suff w:val="tab"/>
    <w:pPr><w:ind w:left="600" w:hanging="300"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/><w:color w:val="4472C4"/></w:rPr></w:lvl>
 </w:abstractNum>
 <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>"""

# The document body.
DOCUMENT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:v="urn:schemas-microsoft-com:vml"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:w10="urn:schemas-microsoft-com:office:word">
 <w:body>

  <w:p><w:r><w:pict>
    <v:rect id="banner" style="position:absolute;left:0;top:0;width:612pt;height:64pt;z-index:-1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page" fillcolor="#44546A" stroked="f">
      <v:textbox inset="36pt,14pt,18pt,14pt">
        <w:txbxContent>
          <w:p><w:r><w:rPr><w:color w:val="FFFFFF"/><w:b/><w:sz w:val="44"/><w:spacing w:val="30"/></w:rPr><w:t>KAI CARTER</w:t></w:r></w:p>
          <w:p><w:r><w:rPr><w:color w:val="FFC000"/><w:caps/><w:sz w:val="22"/></w:rPr><w:t>General Practitioner</w:t></w:r></w:p>
        </w:txbxContent>
      </v:textbox>
    </v:rect>
  </w:pict></w:r></w:p>

  <w:p><w:pPr><w:spacing w:before="1100"/></w:pPr></w:p>

  <w:tbl>
   <w:tblPr><w:tblW w:w="10800" w:type="dxa"/><w:tblLayout w:type="fixed"/>
     <w:tblCellMar><w:top w:w="120" w:type="dxa"/><w:left w:w="180" w:type="dxa"/>
       <w:bottom w:w="120" w:type="dxa"/><w:right w:w="180" w:type="dxa"/></w:tblCellMar>
   </w:tblPr>
   <w:tblGrid><w:gridCol w:w="3600"/><w:gridCol w:w="7200"/></w:tblGrid>
   <w:tr>
    <w:tc>
     <w:tcPr><w:tcW w:w="3600" w:type="dxa"/><w:vMerge w:val="restart"/>
       <w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr>
     <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing>
        <wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1143000" cy="1143000"/>
          <wp:docPr id="1" name="Photo"/>
          <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>
            <pic:nvPicPr><pic:cNvPr id="1" name="Photo"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
            <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1143000" cy="1143000"/></a:xfrm>
              <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></pic:spPr>
          </pic:pic></a:graphicData></a:graphic>
        </wp:inline></w:drawing></w:r></w:p>
     <w:p><w:pPr><w:spacing w:before="160" w:after="40"/></w:pPr><w:r><w:rPr><w:b/><w:caps/><w:color w:val="FFFFFF"/><w:sz w:val="24"/></w:rPr><w:t>Contact</w:t></w:r></w:p>
     <w:p><w:r><w:rPr><w:rStyle w:val="SidebarText"/></w:rPr><w:t>kai.carter@example.com</w:t></w:r></w:p>
     <w:p><w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t>+1 (415) 555-0199</w:t></w:r></w:p>
     <w:p><w:pPr><w:spacing w:before="160" w:after="40"/></w:pPr><w:r><w:rPr><w:b/><w:caps/><w:color w:val="FFFFFF"/><w:sz w:val="24"/></w:rPr><w:t>Skills</w:t></w:r></w:p>
     <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t>Clinical diagnosis</w:t></w:r></w:p>
     <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t>Patient care</w:t></w:r></w:p>
    </w:tc>
    <w:tc>
     <w:tcPr><w:tcW w:w="7200" w:type="dxa"/></w:tcPr>
     <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Profile</w:t></w:r></w:p>
     <w:p><w:r><w:t xml:space="preserve">Experienced and compassionate GP dedicated to delivering </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>excellent patient care</w:t></w:r><w:r><w:t xml:space="preserve">. Known for a </w:t></w:r><w:r><w:rPr><w:i/><w:color w:val="ED7D31"/></w:rPr><w:t>patient-centered approach</w:t></w:r><w:r><w:t>.</w:t></w:r></w:p>
    </w:tc>
   </w:tr>
   <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="3600" w:type="dxa"/><w:vMerge/><w:shd w:val="clear" w:color="auto" w:fill="4472C4"/></w:tcPr><w:p/></w:tc>
    <w:tc>
     <w:tcPr><w:tcW w:w="7200" w:type="dxa"/></w:tcPr>
     <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Experience</w:t></w:r></w:p>
     <w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="7000"/></w:tabs></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Larna Healthcare</w:t></w:r><w:r><w:tab/></w:r><w:r><w:rPr><w:color w:val="808080"/></w:rPr><w:t>2020 - Present</w:t></w:r></w:p>
     <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Evidence-based medicine for accurate diagnosis.</w:t></w:r></w:p>
     <w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Reduced screenings to over 200 residents.</w:t></w:r></w:p>
    </w:tc>
   </w:tr>
  </w:tbl>

  <w:sectPr>
   <w:pgSz w:w="12240" w:h="15840"/>
   <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0" w:gutter="0"/>
   <w:cols w:space="720"/>
  </w:sectPr>
 </w:body>
</w:document>"""


def main():
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("sample.docx")
    png = make_png(64, 64, (91, 155, 213))  # accent5 blue square "photo"
    parts = {
        "[Content_Types].xml": CONTENT_TYPES,
        "_rels/.rels": RELS,
        "word/document.xml": DOCUMENT,
        "word/_rels/document.xml.rels": DOC_RELS,
        "word/styles.xml": STYLES,
        "word/numbering.xml": NUMBERING,
        "word/theme/theme1.xml": THEME,
    }
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for name, text in parts.items():
            z.writestr(name, text)
        z.writestr("word/media/image1.png", png)
    print("wrote", out, "(%d bytes)" % out.stat().st_size)


if __name__ == "__main__":
    main()
