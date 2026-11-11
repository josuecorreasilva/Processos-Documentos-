"""Template oficial SMS Pelotas — cabeçalho com brasão, caixas e rodapé."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / "docs" / "assets"
BRASAO = ASSETS / "brasao_pelotas.png"
RODAPE = ASSETS / "logo_prefeitura_rodape.jpeg"


def _set_run_font(run, *, name="Arial", size=12, bold=False, color=None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def _set_cell_border(cell, **kwargs):
    """kwargs: top/left/bottom/right = {sz, val, color}"""
    tc = cell._tePr if hasattr(cell, "_tePr") else cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = OxmlElement("w:tcBorders")
    for edge in ("top", "left", "bottom", "right"):
        element = OxmlElement(f"w:{edge}")
        opts = kwargs.get(edge, {"sz": "8", "val": "single", "color": "000000"})
        element.set(qn("w:val"), opts.get("val", "single"))
        element.set(qn("w:sz"), opts.get("sz", "8"))
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), opts.get("color", "000000"))
        tcBorders.append(element)
    tcPr.append(tcBorders)


def _shade_cell(cell, fill="F2F2F2"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def new_official_doc() -> Document:
    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)
    section.header_distance = Cm(0.6)
    section.footer_distance = Cm(0.5)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    style.font.size = Pt(11)
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    pf.space_after = Pt(4)
    pf.space_before = Pt(0)

    # Header
    header = section.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.space_after = Pt(0)
    run = hp.add_run()
    if BRASAO.exists():
        run.add_picture(str(BRASAO), width=Cm(1.7))

    for text, size, bold in [
        ("ESTADO DO RIO GRANDE DO SUL", 10, True),
        ("MUNICÍPIO DE PELOTAS", 10, True),
        ("SECRETARIA MUNICIPAL DE SAÚDE", 10, True),
    ]:
        p = header.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(text)
        _set_run_font(r, size=size, bold=bold)

    # thin line under header
    line = header.add_paragraph()
    line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    line.paragraph_format.space_before = Pt(2)
    line.paragraph_format.space_after = Pt(0)
    pPr = line._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)

    # Footer
    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(0)
    fp.paragraph_format.space_after = Pt(0)
    # top border
    pPr = fp._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    top = OxmlElement("w:top")
    top.set(qn("w:val"), "single")
    top.set(qn("w:sz"), "6")
    top.set(qn("w:space"), "4")
    top.set(qn("w:color"), "000000")
    pBdr.append(top)
    pPr.append(pBdr)
    if RODAPE.exists():
        run = fp.add_run()
        run.add_picture(str(RODAPE), width=Cm(4.2))
    else:
        r = fp.add_run("Prefeitura de Pelotas — Secretaria Municipal de Saúde")
        _set_run_font(r, size=8)

    return doc


def title(doc: Document, text: str, subtitle: str | None = None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text.upper())
    _set_run_font(r, size=13, bold=True)
    if subtitle:
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after = Pt(10)
        r2 = p2.add_run(subtitle)
        _set_run_font(r2, size=10, bold=False)
    return p


def meta_line(doc: Document, label: str, value: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r1 = p.add_run(f"{label}: ")
    _set_run_font(r1, size=11, bold=True)
    r2 = p.add_run(value)
    _set_run_font(r2, size=11, bold=False)
    return p


def heading(doc: Document, text: str, level: int = 1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10 if level == 1 else 6)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    _set_run_font(r, size=11 if level == 1 else 10, bold=True)
    return p


def para(doc: Document, text: str, *, bold: bool = False, justify: bool = True, first_indent: bool = True):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT
    if first_indent:
        p.paragraph_format.first_line_indent = Cm(1.0)
    r = p.add_run(text)
    _set_run_font(r, size=11, bold=bold)
    return p


def bullet(doc: Document, text: str):
    p = doc.add_paragraph(style="List Bullet")
    p.clear()
    r = p.add_run(text)
    _set_run_font(r, size=11)
    return p


def boxed(doc: Document, title_text: str | None, body_lines: list[str], *, shaded_title: bool = True):
    """Caixa oficial no estilo do DFD da SMS (título + corpo em até 2 linhas de tabela)."""
    n_rows = 1 + (1 if title_text else 0)
    # if only title, 1 row; if title+body, 2 rows; if only body, 1 row
    if title_text and body_lines:
        n_rows = 2
    elif title_text or body_lines:
        n_rows = 1
    else:
        return None

    table = doc.add_table(rows=n_rows, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    def _pad(cell):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcMar = OxmlElement("w:tcMar")
        for edge, val in (("top", "80"), ("left", "100"), ("bottom", "80"), ("right", "100")):
            node = OxmlElement(f"w:{edge}")
            node.set(qn("w:w"), val)
            node.set(qn("w:type"), "dxa")
            tcMar.append(node)
        tcPr.append(tcMar)
        _set_cell_border(cell)

    row_idx = 0
    if title_text:
        cell = table.rows[row_idx].cells[0]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r = p.add_run(title_text)
        _set_run_font(r, size=11, bold=True)
        if shaded_title:
            _shade_cell(cell, "EDEDED")
        _pad(cell)
        row_idx += 1

    if body_lines:
        cell = table.rows[row_idx if title_text else 0].cells[0]
        cell.text = ""
        first = True
        for line in body_lines:
            p = cell.paragraphs[0] if first else cell.add_paragraph()
            first = False
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if line else WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(2)
            r = p.add_run(line)
            _set_run_font(r, size=11, bold=False)
        _pad(cell)

    doc.add_paragraph()
    return table


def simple_table(doc: Document, headers: list[str], rows: list[list[str]]):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        r = p.add_run(h)
        _set_run_font(r, size=9, bold=True)
        _shade_cell(cell, "EDEDED")
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row):
            cell = table.rows[r_idx].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            r = p.add_run(value)
            _set_run_font(r, size=9)
    doc.add_paragraph()
    return table


def signatures(doc: Document, blocks: list[list[str]]):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Pelotas/RS, _____ de ______________ de 2026.")
    _set_run_font(r, size=11)
    for block in blocks:
        doc.add_paragraph()
        line = doc.add_paragraph()
        line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        rr = line.add_run("_" * 42)
        _set_run_font(rr, size=11)
        for text in block:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(text)
            _set_run_font(r, size=10, bold=False)


def checkbox_line(doc: Document, text: str, checked: bool = False):
    mark = "x" if checked else " "
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(1)
    r = p.add_run(f"( {mark} ) {text}")
    _set_run_font(r, size=11)
    return p


def save(doc: Document, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    print(f"OK {path}")
