#!/usr/bin/env python3
"""
build_kw1c_report.py — Bouwt een KW1C-huisstijl Word-document.

Gebruik:
    python build_kw1c_report.py content.json output.docx

content.json structuur:
    {
      "doc_title": "Tekst die in de voettekst komt (kort, max ~50 tekens)",
      "title_page": {                           // optioneel; weglaten = geen titelpagina
        "title": "HOOFDTITEL",
        "subtitle": "Ondertitel (optioneel)",
        "author": "Door wie",
        "version": "Versie/datum"
      },
      "toc": true,                              // optioneel; true = inhoudsopgave invoegen
      "sections": [
        {
          "heading": "Hoofdstuk 1: De titel",   // wordt H1 (Impact 24pt KW1C-blauw)
          "level": 1,
          "blocks": [
            {"type": "paragraph", "text": "Body tekst."},
            {"type": "heading", "level": 2, "text": "1.1 Een paragraaftitel"},
            {"type": "paragraph", "text": "Meer tekst, eventueel met **vet** of *cursief*."},
            {"type": "bullets", "items": ["Punt een", "Punt twee", "Punt drie"]},
            {"type": "numbered", "items": ["Eerste", "Tweede"]},
            {"type": "table", "rows": [
                ["Label", "Inhoud"],
                ["Tweede", "Nog een rij"]
              ],
              "label_column": true               // optioneel; true = eerste kolom als blauw label
            },
            {"type": "callout", "text": "Belangrijke notitie in lichtblauw kader."}
          ]
        }
      ]
    }

De skill SKILL.md beschrijft hoe Claude dit script aanroept.
"""

import json
import re
import shutil
import sys
import zipfile
from pathlib import Path

# Kleuren (KW1C huisstijl)
RED = "EF2B2D"
BLUE = "0085CA"
GREY = "7F7F7F"
LIGHTBLUE_FILL = "DEEBF7"

# ---------------------------------------------------------------------------
# XML helpers
# ---------------------------------------------------------------------------

def escape_xml(text: str) -> str:
    """XML-escape inclusief slimme aanhalingstekens."""
    if text is None:
        return ""
    text = (text
            .replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
            .replace('"', "&quot;"))
    # Slimme aanhalingstekens en apostrof voor professionele uitstraling
    text = text.replace("'", "&#x2019;")
    text = re.sub(r'(?<=\s)"(?=\S)', "&#x201C;", text)  # opening
    text = text.replace('"', "&#x201D;")                 # closing fallback
    return text


def runs_from_inline_markdown(text: str, extra_rpr: str = "") -> str:
    """Verwerk **vet** en *cursief* binnen een paragraaf. Retourneert <w:r>...</w:r> blokken."""
    # Split op markers maar behoud delimiters
    pattern = r'(\*\*[^*]+\*\*|\*[^*]+\*)'
    parts = re.split(pattern, text)
    out = []
    for part in parts:
        if not part:
            continue
        bold = italic = False
        if part.startswith("**") and part.endswith("**"):
            bold = True
            part = part[2:-2]
        elif part.startswith("*") and part.endswith("*"):
            italic = True
            part = part[1:-1]
        rpr_inner = extra_rpr
        if bold:
            rpr_inner += "<w:b/>"
        if italic:
            rpr_inner += "<w:i/>"
        rpr = f"<w:rPr>{rpr_inner}</w:rPr>" if rpr_inner else ""
        out.append(f'<w:r>{rpr}<w:t xml:space="preserve">{escape_xml(part)}</w:t></w:r>')
    return "".join(out)


# ---------------------------------------------------------------------------
# Bouwstenen
# ---------------------------------------------------------------------------

def build_title_page(tp: dict) -> str:
    """Genereert de titelpagina-XML."""
    title = escape_xml(tp.get("title", "TITEL"))
    subtitle = escape_xml(tp.get("subtitle", ""))
    author = escape_xml(tp.get("author", ""))
    version = escape_xml(tp.get("version", ""))

    parts = [
        # Witruimte bovenaan
        '<w:p><w:pPr><w:spacing w:before="3600" w:after="0"/></w:pPr></w:p>',
        # Hoofdtitel: Impact, KW1C-rood, 36pt, hoofdletters
        f'''<w:p>
            <w:pPr><w:spacing w:before="0" w:after="80"/></w:pPr>
            <w:r>
              <w:rPr>
                <w:rFonts w:ascii="Impact" w:hAnsi="Impact"/>
                <w:color w:val="{RED}"/>
                <w:sz w:val="72"/>
                <w:szCs w:val="72"/>
                <w:caps/>
              </w:rPr>
              <w:t>{title}</w:t>
            </w:r>
        </w:p>''',
    ]
    if subtitle:
        parts.append(f'''<w:p>
            <w:pPr><w:pStyle w:val="Ondertitel"/></w:pPr>
            <w:r>
              <w:rPr><w:color w:val="{GREY}"/></w:rPr>
              <w:t>{subtitle}</w:t>
            </w:r>
        </w:p>''')

    if author or version:
        parts.append('<w:p><w:pPr><w:spacing w:before="4000" w:after="0"/></w:pPr></w:p>')
        if author:
            parts.append(f'''<w:p>
                <w:pPr><w:spacing w:after="80"/></w:pPr>
                <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Door: </w:t></w:r>
                <w:r><w:t>{author}</w:t></w:r>
            </w:p>''')
        if version:
            parts.append(f'''<w:p>
                <w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Versie: </w:t></w:r>
                <w:r><w:t>{version}</w:t></w:r>
            </w:p>''')

    # Pagina-einde
    parts.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    return "".join(parts)


def build_toc() -> str:
    """Genereert een automatische inhoudsopgave."""
    return '''<w:p>
        <w:pPr>
          <w:spacing w:before="240" w:after="240"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Impact" w:hAnsi="Impact"/>
            <w:color w:val="''' + BLUE + '''"/>
            <w:sz w:val="48"/>
          </w:rPr>
          <w:t>INHOUDSOPGAVE</w:t>
        </w:r>
      </w:p>
      <w:p>
        <w:pPr><w:pStyle w:val="Standaard"/></w:pPr>
        <w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r>
        <w:r><w:instrText xml:space="preserve">TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>
        <w:r><w:fldChar w:fldCharType="separate"/></w:r>
        <w:r><w:t>(Werk de inhoudsopgave bij door in Word op F9 te drukken)</w:t></w:r>
        <w:r><w:fldChar w:fldCharType="end"/></w:r>
      </w:p>
      <w:p><w:r><w:br w:type="page"/></w:r></w:p>'''


def build_heading(text: str, level: int, first: bool = False) -> str:
    """Bouwt een kop. Level 1 = Impact KW1C-blauw groot, level 2 = Aptos vet, level 3 = Aptos vet kleiner.

    Voor level 1: elke hoofdstuk-kop start op een nieuwe pagina, behalve als first=True
    (eerste hoofdstuk na titelpagina/TOC, daar zorgt een page break al voor de afsluiting).
    """
    safe = escape_xml(text)
    if level == 1:
        page_break = "" if first else '<w:pageBreakBefore/>'
        return f'''<w:p>
            <w:pPr>
              {page_break}
              <w:spacing w:before="360" w:after="240"/>
              <w:outlineLvl w:val="0"/>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:rFonts w:ascii="Impact" w:hAnsi="Impact"/>
                <w:color w:val="{BLUE}"/>
                <w:sz w:val="48"/>
                <w:caps/>
              </w:rPr>
              <w:t>{safe}</w:t>
            </w:r>
          </w:p>'''
    if level == 2:
        return f'''<w:p>
            <w:pPr>
              <w:spacing w:before="240" w:after="60"/>
              <w:outlineLvl w:val="1"/>
            </w:pPr>
            <w:r>
              <w:rPr>
                <w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/>
                <w:b/>
                <w:sz w:val="28"/>
              </w:rPr>
              <w:t>{safe}</w:t>
            </w:r>
          </w:p>'''
    # level 3 of dieper
    return f'''<w:p>
        <w:pPr>
          <w:spacing w:before="200" w:after="60"/>
          <w:outlineLvl w:val="2"/>
        </w:pPr>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/>
            <w:b/>
            <w:sz w:val="24"/>
          </w:rPr>
          <w:t>{safe}</w:t>
        </w:r>
      </w:p>'''


def build_paragraph(text: str) -> str:
    runs = runs_from_inline_markdown(text)
    return f'<w:p><w:pPr><w:pStyle w:val="Standaard"/></w:pPr>{runs}</w:p>'


def build_bullets(items: list) -> str:
    """Ongeordende lijst via echte Word-numbering (verwijst naar numId 200 in numbering.xml)."""
    out = []
    for item in items:
        runs = runs_from_inline_markdown(item)
        out.append(f'''<w:p>
            <w:pPr>
              <w:pStyle w:val="Standaard"/>
              <w:numPr>
                <w:ilvl w:val="0"/>
                <w:numId w:val="200"/>
              </w:numPr>
              <w:spacing w:after="80"/>
              <w:contextualSpacing/>
            </w:pPr>
            {runs}
          </w:p>''')
    return "".join(out)


def build_numbered(items: list) -> str:
    """Genummerde lijst via echte Word-numbering (verwijst naar numId 201 in numbering.xml)."""
    out = []
    for item in items:
        runs = runs_from_inline_markdown(item)
        out.append(f'''<w:p>
            <w:pPr>
              <w:pStyle w:val="Standaard"/>
              <w:numPr>
                <w:ilvl w:val="0"/>
                <w:numId w:val="201"/>
              </w:numPr>
              <w:spacing w:after="80"/>
              <w:contextualSpacing/>
            </w:pPr>
            {runs}
          </w:p>''')
    return "".join(out)


def build_table(rows: list, label_column: bool = False) -> str:
    """Bouwt een tabel met blauwe randen, optioneel met lichtblauwe label-kolom."""
    if not rows:
        return ""
    num_cols = len(rows[0])
    total_width = 9070  # content width binnen marges
    if label_column and num_cols >= 2:
        col_widths = [2500] + [(total_width - 2500) // (num_cols - 1)] * (num_cols - 1)
    else:
        col_widths = [total_width // num_cols] * num_cols
    # Stel resterende breedte bij
    diff = total_width - sum(col_widths)
    col_widths[-1] += diff

    gridcols = "".join(f'<w:gridCol w:w="{w}"/>' for w in col_widths)
    borders = f'''<w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="{BLUE}"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="{BLUE}"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="{BLUE}"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="{BLUE}"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{BLUE}"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="{BLUE}"/>
      </w:tblBorders>'''

    tr_xml = []
    for row in rows:
        cells = []
        for i, cell in enumerate(row):
            is_label = label_column and i == 0
            shading = f'<w:shd w:val="clear" w:color="auto" w:fill="{LIGHTBLUE_FILL}"/>' if is_label else ""
            rpr = f'<w:rPr><w:b/><w:color w:val="{BLUE}"/></w:rPr>' if is_label else ""
            content_runs = (
                f'<w:r>{rpr}<w:t xml:space="preserve">{escape_xml(str(cell))}</w:t></w:r>'
                if is_label else
                runs_from_inline_markdown(str(cell))
            )
            cells.append(f'''<w:tc>
                <w:tcPr>
                  <w:tcW w:w="{col_widths[i]}" w:type="dxa"/>
                  {shading}
                  <w:tcMar><w:top w:w="80" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="80" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tcMar>
                </w:tcPr>
                <w:p><w:pPr><w:pStyle w:val="Standaard"/><w:spacing w:after="0"/></w:pPr>{content_runs}</w:p>
              </w:tc>''')
        tr_xml.append(f"<w:tr>{''.join(cells)}</w:tr>")

    return f'''<w:tbl>
        <w:tblPr>
          <w:tblW w:w="{total_width}" w:type="dxa"/>
          {borders}
        </w:tblPr>
        <w:tblGrid>{gridcols}</w:tblGrid>
        {''.join(tr_xml)}
      </w:tbl>
      <w:p><w:pPr><w:pStyle w:val="Standaard"/><w:spacing w:after="120"/></w:pPr></w:p>'''


def build_callout(text: str) -> str:
    """Een 1-cel tabel met lichtblauwe achtergrond als notitie/callout."""
    runs = runs_from_inline_markdown(text)
    return f'''<w:tbl>
        <w:tblPr>
          <w:tblW w:w="9070" w:type="dxa"/>
          <w:tblBorders>
            <w:top w:val="single" w:sz="12" w:space="0" w:color="{BLUE}"/>
            <w:left w:val="single" w:sz="12" w:space="0" w:color="{BLUE}"/>
            <w:bottom w:val="single" w:sz="12" w:space="0" w:color="{BLUE}"/>
            <w:right w:val="single" w:sz="12" w:space="0" w:color="{BLUE}"/>
          </w:tblBorders>
        </w:tblPr>
        <w:tblGrid><w:gridCol w:w="9070"/></w:tblGrid>
        <w:tr>
          <w:tc>
            <w:tcPr>
              <w:tcW w:w="9070" w:type="dxa"/>
              <w:shd w:val="clear" w:color="auto" w:fill="{LIGHTBLUE_FILL}"/>
              <w:tcMar><w:top w:w="120" w:type="dxa"/><w:left w:w="200" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="200" w:type="dxa"/></w:tcMar>
            </w:tcPr>
            <w:p><w:pPr><w:pStyle w:val="Standaard"/><w:spacing w:after="0"/></w:pPr>{runs}</w:p>
          </w:tc>
        </w:tr>
      </w:tbl>
      <w:p><w:pPr><w:pStyle w:val="Standaard"/><w:spacing w:after="120"/></w:pPr></w:p>'''


def build_block(block: dict, first_heading: bool = False) -> str:
    btype = block.get("type", "paragraph")
    if btype == "paragraph":
        return build_paragraph(block.get("text", ""))
    if btype == "heading":
        return build_heading(block.get("text", ""), block.get("level", 2), first=first_heading)
    if btype == "bullets":
        return build_bullets(block.get("items", []))
    if btype == "numbered":
        return build_numbered(block.get("items", []))
    if btype == "table":
        return build_table(block.get("rows", []), block.get("label_column", False))
    if btype == "callout":
        return build_callout(block.get("text", ""))
    raise ValueError(f"Onbekend block-type: {btype}")


def build_section(section: dict, is_first: bool = False) -> str:
    parts = []
    if "heading" in section:
        parts.append(build_heading(section["heading"], section.get("level", 1), first=is_first))
    for block in section.get("blocks", []):
        parts.append(build_block(block))
    return "".join(parts)


# ---------------------------------------------------------------------------
# Hoofdroutine
# ---------------------------------------------------------------------------

SECTPR = '''<w:sectPr>
    <w:headerReference w:type="default" r:id="rIdHeader"/>
    <w:footerReference w:type="default" r:id="rIdFooter"/>
    <w:headerReference w:type="first" r:id="rIdHeaderFirst"/>
    <w:footerReference w:type="first" r:id="rIdFooterFirst"/>
    <w:pgSz w:w="11906" w:h="16838" w:code="9"/>
    <w:pgMar w:top="1701" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/>
    <w:cols w:space="708"/>
    <w:titlePg/>
    <w:docGrid w:linePitch="360"/>
  </w:sectPr>'''


def build_report(content: dict, template_path: Path, output_path: Path) -> None:
    # Schrijf de template uit naar een tijdelijke werkmap
    workdir = output_path.parent / f".{output_path.stem}_work"
    if workdir.exists():
        shutil.rmtree(workdir)
    workdir.mkdir(parents=True)

    with zipfile.ZipFile(template_path, "r") as zf:
        zf.extractall(workdir)

    # Bouw body
    body_parts = []
    if "title_page" in content and content["title_page"]:
        body_parts.append(build_title_page(content["title_page"]))
    if content.get("toc"):
        body_parts.append(build_toc())

    # Bepaal of de eerste sectie geen pageBreakBefore moet krijgen
    # Als er geen titelpagina en geen TOC is, start de eerste H1 gewoon op pagina 1.
    # Als er wel een titelpagina of TOC is, is er al een pagebreak ingevoegd; dan moet de
    # eerste H1 NIET ook nog een pageBreakBefore krijgen (dat zou een lege pagina geven).
    sections = content.get("sections", [])
    for idx, section in enumerate(sections):
        is_first = (idx == 0)
        body_parts.append(build_section(section, is_first=is_first))

    # Als er geen titelpagina is, moet de eerste pagina ook de "first"-header en footer
    # niet leeg laten. Word gebruikt dan automatisch de default header/footer.
    # Door <w:titlePg/> alleen in te zetten als er een titelpagina is, regelt Word dit netjes.
    if "title_page" not in content or not content["title_page"]:
        sectpr_no_titlepg = SECTPR.replace("<w:titlePg/>", "")
        body_xml = f'<w:body>{"".join(body_parts)}{sectpr_no_titlepg}</w:body>'
    else:
        body_xml = f'<w:body>{"".join(body_parts)}{SECTPR}</w:body>'

    # Schrijf naar document.xml
    doc_path = workdir / "word" / "document.xml"
    doc_xml = doc_path.read_text(encoding="utf-8")
    # Gebruik een lambda om back-slash escaping in de replacement string te vermijden
    doc_xml = re.sub(r'<w:body>.*?</w:body>', lambda m: body_xml, doc_xml, flags=re.DOTALL)
    doc_path.write_text(doc_xml, encoding="utf-8")

    # Voettekst-titel invullen
    footer_path = workdir / "word" / "footer1.xml"
    footer_xml = footer_path.read_text(encoding="utf-8")
    doc_title = escape_xml(content.get("doc_title", ""))
    footer_xml = footer_xml.replace("{{DOC_TITLE}}", doc_title)
    footer_path.write_text(footer_xml, encoding="utf-8")

    # Pak in als .docx (zip)
    if output_path.exists():
        output_path.unlink()
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in workdir.rglob("*"):
            if file.is_file():
                arcname = str(file.relative_to(workdir))
                zf.write(file, arcname)

    # Werkmap opruimen
    shutil.rmtree(workdir)
    print(f"Rapport opgeslagen: {output_path}")


def main():
    if len(sys.argv) != 3:
        print("Gebruik: python build_kw1c_report.py content.json output.docx")
        sys.exit(1)
    content_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2])
    content = json.loads(content_path.read_text(encoding="utf-8"))

    # Template moet naast de skill-map staan
    script_dir = Path(__file__).resolve().parent
    template_path = script_dir.parent / "assets" / "kw1c_template.docx"
    if not template_path.exists():
        print(f"Template niet gevonden: {template_path}")
        sys.exit(1)

    build_report(content, template_path, output_path)


if __name__ == "__main__":
    main()
