"""Helpers de formatação institucional SMS Pelotas para documentos .docx."""

from pathlib import Path
from typing import Iterable, Sequence

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor

from .dados import META

COR_TITULO = RGBColor(0x1B, 0x36, 0x5D)
COR_TEXTO = RGBColor(0x1A, 0x1A, 0x1A)
COR_SUAVE = RGBColor(0x4A, 0x55, 0x68)
HEX_HEADER_BG = "E6ECF2"
HEX_TOTAL_BG = "EEF3F8"
HEX_CALLOUT_BG = "F4F7FA"
HEX_CALLOUT_BORDER = "1B365D"
HEX_BORDER = "BFBFBF"

BRASAO_PATH = Path(__file__).resolve().parents[1] / "assets" / "brasao_pelotas.png"


def set_cell_shading(cell, fill_hex: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    for existing in tc_pr.findall(qn("w:shd")):
        tc_pr.remove(existing)
    tc_pr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))


def set_cell_margins(cell, top=60, bottom=60, left=100, right=100) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    tc_pr.append(tc_mar)


def set_table_borders(table, color: str = HEX_BORDER, sz: str = "4") -> None:
    tbl_pr = table._tbl.tblPr
    for existing in tbl_pr.findall(qn("w:tblBorders")):
        tbl_pr.remove(existing)
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f"</w:tblBorders>"
    )
    tbl_pr.append(borders)


def set_row_repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:tblHeader"))


def set_row_cant_split(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:cantSplit"))


def add_xml_field(paragraph, field_name: str, font_size: float = 9.0) -> None:
    run = paragraph.add_run()
    run.font.name = "Arial"
    run.font.size = Pt(font_size)
    run.font.color.rgb = COR_SUAVE
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {field_name} "
    fld_sep = OxmlElement("w:fldChar")
    fld_sep.set(qn("w:fldCharType"), "separate")
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_sep)
    run._r.append(fld_end)


def create_institutional_doc(titulo_principal: str, subtitulo: str | None = None) -> Document:
    """Cria documento A4 com cabeçalho oficial (brasão de Pelotas) e rodapé Página X de Y."""
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)
    sec.top_margin = Cm(2.8)
    sec.bottom_margin = Cm(2.0)
    sec.header_distance = Cm(0.8)
    sec.footer_distance = Cm(1.0)

    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = COR_TEXTO
    normal.paragraph_format.space_after = Pt(4)
    normal.paragraph_format.line_spacing = 1.15

    # Cabeçalho institucional com brasão de Pelotas
    hdr = sec.header
    p_img = hdr.paragraphs[0]
    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img.paragraph_format.space_before = Pt(0)
    p_img.paragraph_format.space_after = Pt(2)
    if BRASAO_PATH.exists():
        p_img.add_run().add_picture(str(BRASAO_PATH), height=Cm(1.45))

    for idx, linha in enumerate(
        (META["orgao_linha1"], META["orgao_linha2"], META["orgao_linha3"])
    ):
        p = hdr.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0 if idx < 2 else 4)
        p.paragraph_format.line_spacing = 1.05
        r = p.add_run(linha)
        r.font.name = "Arial"
        r.font.size = Pt(9.5 if idx < 2 else 10.0)
        r.bold = True
        r.font.color.rgb = COR_TITULO

    # Rodapé com paginação nativa Página X de Y
    ftr = sec.footer
    p_f = ftr.paragraphs[0]
    p_f.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_f.paragraph_format.space_before = Pt(0)
    p_f.paragraph_format.space_after = Pt(0)
    for text_part, field in (("Página ", None), (None, "PAGE"), (" de ", None), (None, "NUMPAGES")):
        if text_part is not None:
            r = p_f.add_run(text_part)
            r.font.name = "Arial"
            r.font.size = Pt(9.0)
            r.font.color.rgb = COR_SUAVE
        else:
            add_xml_field(p_f, field, font_size=9.0)

    # Título do documento
    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_before = Pt(4)
    p_t.paragraph_format.space_after = Pt(2 if subtitulo else 8)
    rt = p_t.add_run(titulo_principal)
    rt.font.name = "Arial"
    rt.font.size = Pt(12.5)
    rt.bold = True
    rt.font.color.rgb = COR_TITULO

    if subtitulo:
        p_s = doc.add_paragraph()
        p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_s.paragraph_format.space_before = Pt(0)
        p_s.paragraph_format.space_after = Pt(8)
        rs = p_s.add_run(subtitulo)
        rs.font.name = "Arial"
        rs.font.size = Pt(9.5)
        rs.font.color.rgb = COR_SUAVE

    return doc


def add_section_heading(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    p_pr = p._p.get_or_add_pPr()
    p_bdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="6" w:space="2" w:color="1B365D"/>'
        f"</w:pBdr>"
    )
    p_pr.append(p_bdr)
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(11.0)
    r.bold = True
    r.font.color.rgb = COR_TITULO


def add_subheading(doc: Document, text: str) -> None:
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(10.5)
    r.bold = True
    r.font.color.rgb = COR_TITULO


def add_body_p(
    doc: Document,
    text: str = "",
    *,
    bold_prefix: str | None = None,
    justify: bool = True,
    space_after: float = 4.0,
    font_size: float = 10.5,
):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        rb = p.add_run(bold_prefix)
        rb.font.name = "Arial"
        rb.font.size = Pt(font_size)
        rb.bold = True
        rb.font.color.rgb = COR_TEXTO
    if text:
        r = p.add_run(text)
        r.font.name = "Arial"
        r.font.size = Pt(font_size)
        r.font.color.rgb = COR_TEXTO
    return p


def add_bullet(doc: Document, text: str, *, bold_prefix: str | None = None, font_size: float = 10.0) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(0.75)
    p.paragraph_format.first_line_indent = Cm(-0.35)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.12
    rb0 = p.add_run("• ")
    rb0.font.name = "Arial"
    rb0.font.size = Pt(font_size)
    rb0.bold = True
    rb0.font.color.rgb = COR_TITULO
    if bold_prefix:
        rbp = p.add_run(bold_prefix)
        rbp.font.name = "Arial"
        rbp.font.size = Pt(font_size)
        rbp.bold = True
        rbp.font.color.rgb = COR_TEXTO
    r = p.add_run(text)
    r.font.name = "Arial"
    r.font.size = Pt(font_size)
    r.font.color.rgb = COR_TEXTO


def add_callout(doc: Document, titulo: str, texto: str) -> None:
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    cell.width = Cm(17.0)
    set_cell_shading(cell, HEX_CALLOUT_BG)
    set_cell_margins(cell, top=80, bottom=80, left=140, right=120)
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_pr.append(
        parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'<w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E0"/>'
            f'<w:left w:val="single" w:sz="18" w:space="0" w:color="{HEX_CALLOUT_BORDER}"/>'
            f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="CBD5E0"/>'
            f'<w:right w:val="single" w:sz="4" w:space="0" w:color="CBD5E0"/>'
            f"</w:tcBorders>"
        )
    )
    set_row_cant_split(tbl.rows[0])
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = 1.12
    rt = p.add_run(f"{titulo}: ")
    rt.font.name = "Arial"
    rt.font.size = Pt(9.5)
    rt.bold = True
    rt.font.color.rgb = COR_TITULO
    rb = p.add_run(texto)
    rb.font.name = "Arial"
    rb.font.size = Pt(9.5)
    rb.font.color.rgb = COR_TEXTO
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(0)
    p_sp.paragraph_format.space_after = Pt(3)


def add_styled_table(
    doc: Document,
    headers: Sequence[str],
    rows: Iterable[Sequence[str]],
    col_widths_cm: Sequence[float],
    alignments: Sequence[WD_ALIGN_PARAGRAPH],
    *,
    font_size: float = 8.5,
    total_row: Sequence[str] | None = None,
    merge_total_cols: int | None = None,
):
    rows_list = list(rows)
    n_cols = len(headers)
    total_rows_count = 1 + len(rows_list) + (1 if total_row else 0)
    table = doc.add_table(rows=total_rows_count, cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)

    def _fill_cell(cell, text: str, width_cm: float, align, bold: bool = False, fill: str | None = None):
        cell.width = Cm(width_cm)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_margins(cell, top=55, bottom=55, left=75, right=75)
        if fill:
            set_cell_shading(cell, fill)
        p = cell.paragraphs[0]
        p.alignment = align
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.08
        r = p.add_run(str(text))
        r.font.name = "Arial"
        r.font.size = Pt(font_size)
        r.bold = bold
        r.font.color.rgb = COR_TITULO if (bold and fill == HEX_HEADER_BG) else COR_TEXTO

    # Cabeçalho
    hdr_row = table.rows[0]
    set_row_repeat_header(hdr_row)
    set_row_cant_split(hdr_row)
    for c_idx, h_text in enumerate(headers):
        _fill_cell(
            hdr_row.cells[c_idx],
            h_text,
            col_widths_cm[c_idx],
            WD_ALIGN_PARAGRAPH.CENTER,
            bold=True,
            fill=HEX_HEADER_BG,
        )

    # Linhas de dados
    for r_idx, r_data in enumerate(rows_list, start=1):
        row = table.rows[r_idx]
        set_row_cant_split(row)
        for c_idx, val in enumerate(r_data):
            _fill_cell(row.cells[c_idx], val, col_widths_cm[c_idx], alignments[c_idx])

    # Linha de total
    if total_row:
        t_row = table.rows[-1]
        set_row_cant_split(t_row)
        if merge_total_cols and merge_total_cols > 1:
            merged = t_row.cells[0]
            for c_idx in range(1, merge_total_cols):
                merged = merged.merge(t_row.cells[c_idx])
            merged_w = sum(col_widths_cm[:merge_total_cols])
            _fill_cell(
                t_row.cells[0],
                total_row[0],
                merged_w,
                WD_ALIGN_PARAGRAPH.RIGHT,
                bold=True,
                fill=HEX_TOTAL_BG,
            )
            for c_idx in range(merge_total_cols, n_cols):
                _fill_cell(
                    t_row.cells[c_idx],
                    total_row[c_idx],
                    col_widths_cm[c_idx],
                    alignments[c_idx],
                    bold=True,
                    fill=HEX_TOTAL_BG,
                )
        else:
            for c_idx, val in enumerate(total_row):
                _fill_cell(
                    t_row.cells[c_idx],
                    val,
                    col_widths_cm[c_idx],
                    alignments[c_idx],
                    bold=True,
                    fill=HEX_TOTAL_BG,
                )

    p_after = doc.add_paragraph()
    p_after.paragraph_format.space_before = Pt(0)
    p_after.paragraph_format.space_after = Pt(3)
    return table


def add_signatures_block(
    doc: Document,
    *,
    include_team_intro: str | None = None,
    autorizo_suffix: bool = True,
) -> None:
    if include_team_intro:
        add_body_p(doc, include_team_intro, space_after=4)

    add_body_p(
        doc,
        "Certificamos que somos responsáveis pela elaboração do presente instrumento e que o mesmo "
        "está em consonância com a legislação vigente.",
        space_after=8,
    )

    # Tabela sem bordas com duas colunas para Gestor e Fiscal
    tbl = doc.add_table(rows=1, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    set_row_cant_split(tbl.rows[0])
    for cell in tbl.rows[0].cells:
        tc_pr = cell._tc.get_or_add_tcPr()
        tc_pr.append(
            parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                '<w:top w:val="none"/><w:left w:val="none"/>'
                '<w:bottom w:val="none"/><w:right w:val="none"/>'
                "</w:tcBorders>"
            )
        )
        cell.width = Cm(8.5)

    for cell, nome, mat, papel in (
        (tbl.cell(0, 0), META["gestor_nome"], META["gestor_matricula"], "Gestor"),
        (tbl.cell(0, 1), META["fiscal_nome"], META["fiscal_matricula"], "Fiscal"),
    ):
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        for line, is_bold in (
            ("__________________________________", False),
            (nome, True),
            (f"Matrícula: {mat}", False),
            (papel, False),
        ):
            r = p.add_run(line + "\n")
            r.font.name = "Arial"
            r.font.size = Pt(10.0)
            r.bold = is_bold
            r.font.color.rgb = COR_TEXTO

    p_dt = doc.add_paragraph()
    p_dt.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_dt.paragraph_format.space_before = Pt(12)
    p_dt.paragraph_format.space_after = Pt(14)
    rd = p_dt.add_run("Pelotas/RS, _____ de ______________ de 2026.")
    rd.font.name = "Arial"
    rd.font.size = Pt(10.0)

    p_sec = doc.add_paragraph()
    p_sec.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sec.paragraph_format.space_before = Pt(6)
    p_sec.paragraph_format.space_after = Pt(2)
    p_sec.paragraph_format.line_spacing = 1.15
    cargo_txt = (
        f"{META['secretaria_cargo']} — Autorizo" if autorizo_suffix else META["secretaria_cargo"]
    )
    for line, is_bold in (
        ("__________________________________________", False),
        (META["secretaria_nome"], True),
        (cargo_txt, False),
    ):
        r = p_sec.add_run(line + "\n")
        r.font.name = "Arial"
        r.font.size = Pt(10.0)
        r.bold = is_bold
        r.font.color.rgb = COR_TEXTO
