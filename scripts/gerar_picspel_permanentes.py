#!/usr/bin/env python3
"""Gera e valida DFD Revisado, ETP Revisado e Termo de Referência (TR)
para o Processo 2 — PICSPel: Aquisição de Materiais Permanentes.

Emenda Parlamentar nº 43530013 — Secretaria Municipal de Saúde de Pelotas/RS.
Padrão institucional SMS Pelotas / RP05_TR (brasão, Arial, Página X de Y).
"""

from __future__ import annotations

import argparse
import io
import re
import sys
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor


ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "docs" / "09-PICSPEL-MATERIAIS-PERMANENTES"

PDF_ENTRADA_PADRAO = Path(
    "/cursor/stores/bc-27f07d00-f56f-45e6-ad39-463e9ed80767"
    "/internal/uploads-picspel/ETP_PICSPel_Materiais_Permanentes_d7d3.pdf"
)

COR_CABECALHO_TABELA = "E6ECF2"
COR_TOTAL_TABELA = "F2F5F8"
COR_DESTAQUE_BORDA = "B8C4D0"
COR_TEXTO_CINZA = RGBColor(0x55, 0x55, 0x55)

EMENDA = "43530013"
RECURSO_GLOBAL_EMENDA = Decimal("100000.00")
RECURSO_COTA_SERVICOS = Decimal("40000.00")
RECURSO_COTA_MATERIAIS_REF = Decimal("60000.00")

GESTOR = ("Dariane Lima Portela", "44151", "Gestor")
FISCAL = ("Aline Kohler Geppert", "29371", "Fiscal")
SECRETARIA = ("Angela Moreira Vitória", "Secretária Municipal de Saúde")

EQUIPES_EMULTI = [
    "Centro/Porto",
    "Fragata 1",
    "Fragata 2",
    "Três Vendas 1",
    "Três Vendas 2",
    "Areal",
    "Colônia",
    "Praias",
]

ENDERECO_ALMOXARIFADO = (
    "Almoxarifado da Secretaria Municipal de Saúde de Pelotas/RS, localizado na "
    "Av. Fernando Osório, nº 273, Bairro Três Vendas, Pelotas/RS, CEP 96020-150, "
    "de segunda a sexta-feira, das 8h às 11h30min e das 13h30min às 16h, em dias "
    "úteis, mediante prévio agendamento pelo e-mail smspel.almoxarifado@gmail.com"
)


@dataclass(frozen=True)
class ItemPermanente:
    numero: int
    nome_curto: str
    especificacao: str
    catmat: str
    unidade: str
    quantidade: int
    valor_unitario: Decimal
    destinacao: str
    grupo_destinacao: str
    certificacao_principal: str

    @property
    def valor_total(self) -> Decimal:
        return (self.valor_unitario * self.quantidade).quantize(Decimal("0.01"))


ITENS: list[ItemPermanente] = [
    ItemPermanente(
        numero=1,
        nome_curto="Caixa de som portátil recarregável",
        especificacao=(
            "Caixa de som portátil recarregável para apoio a práticas corporais, "
            "meditação e atividades educativas coletivas: sistema de áudio estéreo "
            "com dois alto-falantes (ou transdutores equivalentes), potência total "
            "mínima de 20 W RMS, faixa de resposta de frequência abrangendo no "
            "mínimo de 65 Hz a 20 kHz (ou faixa ampla equivalente até 20 kHz), "
            "conectividade sem fio Bluetooth (versão 5.0 ou superior), bateria "
            "interna recarregável com autonomia mínima de 6 horas em uso moderado, "
            "carregamento via cabo USB/USB-C ou fonte compatível com rede elétrica "
            "220 V / bivolt automático (100–240 V), acompanhada de cabo de "
            "carregamento (e adaptador/fonte padrão brasileiro ABNT NBR 14136, "
            "quando aplicável). Dimensões compactas para transporte manual: "
            "referência aproximada de 170 mm × 80 mm × 80 mm (admitida variação "
            "de até ±20% nas dimensões, desde que mantida a portabilidade). "
            "Produto novo, com manual em português, garantia mínima de 12 (doze) "
            "meses e certificado de homologação na ANATEL."
        ),
        catmat="[A CONFIRMAR]",
        unidade="UN",
        quantidade=8,
        valor_unitario=Decimal("600.00"),
        destinacao="eMulti (1 un. × 8 equipes)",
        grupo_destinacao="eMulti",
        certificacao_principal="Homologação ANATEL (dispositivo Bluetooth)",
    ),
    ItemPermanente(
        numero=2,
        nome_curto="Garrafa térmica em aço inoxidável (aprox. 1,8 L)",
        especificacao=(
            "Garrafa térmica de mesa/transporte confeccionada em aço inoxidável de "
            "grau alimentício (corpo externo e parede/ampola interna 100% em aço "
            "inox, dupla parede com isolamento a vácuo, inquebrável, sem ampola de "
            "vidro), capacidade nominal aproximada de 1,8 L (admitida variação de "
            "1,7 L a 2,0 L), sistema de acionamento por bombeamento de pressão "
            "superior com trava de segurança e bico direcionador antigotejamento, "
            "tampa com vedação hermética, alça resistente para transporte e base "
            "estável. Desempenho térmico mínimo de conservação de líquidos quentes "
            "por 6 (seis) horas. Produto novo, de fácil higienização, em "
            "conformidade com as normas técnicas aplicáveis (ABNT NBR 13282 / "
            "certificação INMETRO, quando aplicável), com garantia mínima de 12 "
            "(doze) meses contra defeitos de fabricação e vedação térmica."
        ),
        catmat="[A CONFIRMAR]",
        unidade="UN",
        quantidade=24,
        valor_unitario=Decimal("150.00"),
        destinacao="eMulti (3 un. × 8 equipes)",
        grupo_destinacao="eMulti",
        certificacao_principal="Conformidade ABNT NBR 13282 / INMETRO (quando aplicável)",
    ),
    ItemPermanente(
        numero=3,
        nome_curto="Maleta organizadora rígida em ABS (tipo frasqueira)",
        especificacao=(
            "Maleta organizadora rígida (tipo frasqueira de transporte) para "
            "acondicionamento e deslocamento seguro dos kits de insumos das "
            "equipes eMulti: estrutura rígida confeccionada em polímero ABS (100% "
            "ABS rígido ou polímero de resistência mecânica equivalente ou "
            "superior), leve e resistente a impactos, fechamento perimetral por "
            "zíper duplo na horizontal (abertura tipo tampa), alça superior "
            "reforçada para transporte manual, interior forrado com "
            "compartimento/divisória organizadora e superfície de fácil "
            "higienização. Dimensões aproximadas: altura entre 25 cm e 30 cm, "
            "largura entre 32 cm e 38 cm e profundidade entre 16 cm e 22 cm "
            "(referência: 27,5 cm × 36 cm × 19 cm). Produto novo, com garantia "
            "mínima de 12 (doze) meses contra defeitos de fabricação, estrutura "
            "e zíper."
        ),
        catmat="[A CONFIRMAR]",
        unidade="UN",
        quantidade=8,
        valor_unitario=Decimal("250.00"),
        destinacao="eMulti (1 un. × 8 equipes)",
        grupo_destinacao="eMulti",
        certificacao_principal="Ficha técnica / garantia do fabricante",
    ),
    ItemPermanente(
        numero=4,
        nome_curto="Jarra elétrica em aço inoxidável (1,7 L a 2,2 L – 220 V)",
        especificacao=(
            "Jarra/chaleira elétrica para fervura de água destinada ao preparo de "
            "escalda-pés e atividades correlatas das PICS: corpo em aço inoxidável, "
            "capacidade útil entre 1,7 L e 2,2 L, potência mínima de 1.500 W "
            "(referência 1.800 W), tensão elétrica 220 V (60 Hz), plugue padrão "
            "brasileiro (ABNT NBR 14136 compatível com a potência nominal), base "
            "elétrica destacável com rotação de 360° e porta-fio, alça isotérmica, "
            "tampa com trava de segurança, indicador luminoso de funcionamento, "
            "proteção contra funcionamento a seco e desligamento automático após "
            "atingir a fervura. Produto novo, com manual em português, garantia "
            "mínima de 12 (doze) meses e certificação compulsória de segurança "
            "elétrica no INMETRO (Portaria INMETRO nº 148/2022 ou norma "
            "substitutiva vigente)."
        ),
        catmat="[A CONFIRMAR]",
        unidade="UN",
        quantidade=8,
        valor_unitario=Decimal("100.00"),
        destinacao="eMulti (1 un. × 8 equipes)",
        grupo_destinacao="eMulti",
        certificacao_principal="Certificação compulsória INMETRO (Portaria nº 148/2022)",
    ),
    ItemPermanente(
        numero=5,
        nome_curto="Plastificadora / laminadora (A3/A4/A5/A6 – 220 V)",
        especificacao=(
            "Plastificadora/laminadora (polisseladora) de mesa para confecção de "
            "materiais educativos, protocolos e guias de orientação do Programa "
            "PICSPel: compatível com formatos A3, A4, A5 e A6 (largura útil "
            "mínima de 320 mm), operação em laminação a quente (e opcionalmente a "
            "frio), compatível com filmes/plásticos para plastificação (polaseal) "
            "de espessura entre 0,05 mm (125 micras) e 0,07 mm (175 micras) ou "
            "superior — assegurando plena compatibilidade com os insumos A4 de "
            "0,07 mm previstos no processo correlato de materiais de consumo —, "
            "sistema de desobstrução/reversão anti-atolamento, indicador luminoso "
            "de aquecimento/pronto para uso, tensão elétrica 220 V (60 Hz) ou "
            "bivolt, plugue padrão brasileiro (ABNT NBR 14136). Produto novo, com "
            "manual em português, garantia mínima de 12 (doze) meses e "
            "conformidade com as normas brasileiras de segurança elétrica "
            "aplicáveis."
        ),
        catmat="[A CONFIRMAR]",
        unidade="UN",
        quantidade=1,
        valor_unitario=Decimal("400.00"),
        destinacao="Coordenação PICSPel (1 un.)",
        grupo_destinacao="Coordenação",
        certificacao_principal="Conformidade elétrica / NBR 14136 e garantia do fabricante",
    ),
]

TOTAL_UNIDADES = sum(i.quantidade for i in ITENS)
SUBTOTAL_EMULTI = sum(i.valor_total for i in ITENS if i.grupo_destinacao == "eMulti")
SUBTOTAL_COORD = sum(i.valor_total for i in ITENS if i.grupo_destinacao == "Coordenação")
VALOR_TOTAL = sum(i.valor_total for i in ITENS)

assert TOTAL_UNIDADES == 49
assert SUBTOTAL_EMULTI == Decimal("11200.00")
assert SUBTOTAL_COORD == Decimal("400.00")
assert VALOR_TOTAL == Decimal("11600.00")


def brl(v: Decimal | float) -> str:
    d = v if isinstance(v, Decimal) else Decimal(str(v))
    return f"R$ {d:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def carregar_brasao_bytes(brasao_path: Path | None) -> bytes | None:
    candidatos: list[Path] = []
    if brasao_path:
        candidatos.append(brasao_path)
    candidatos.extend(
        [
            ROOT / "docs" / "assets" / "brasao_rp_header.png",
            ROOT / "docs" / "assets" / "brasao_pelotas.png",
        ]
    )
    for c in candidatos:
        if c.exists():
            return c.read_bytes()

    if PDF_ENTRADA_PADRAO.exists():
        try:
            from pypdf import PdfReader

            reader = PdfReader(str(PDF_ENTRADA_PADRAO))
            for img in reader.pages[0].images:
                if img.data:
                    return bytes(img.data)
        except Exception:
            pass
    return None


# ---------------------------------------------------------------------------
# Utilitários de formatação python-docx (padrão institucional SMS Pelotas)
# ---------------------------------------------------------------------------


def _font(
    run,
    *,
    name: str = "Arial",
    size: float = 10.5,
    bold: bool = False,
    italic: bool = False,
    color: RGBColor | None = None,
):
    run.font.name = name
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), name)
    rFonts.set(qn("w:hAnsi"), name)
    rFonts.set(qn("w:eastAsia"), name)
    rFonts.set(qn("w:cs"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def _shade_cell(cell, fill: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill}" w:val="clear"/>')
    tcPr.append(shd)


def _cell_margins(cell, *, top: int = 75, bottom: int = 75, left: int = 100, right: int = 100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def _set_table_borders(table, color: str = COR_DESTAQUE_BORDA, sz: str = "4"):
    tblPr = table._tbl.tblPr
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
    tblPr.append(borders)


def _set_no_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        "</w:tblBorders>"
    )
    tblPr.append(borders)


def _row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))


def _row_header(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))


def _paragraph_bottom_border(paragraph, color: str = "999999", sz: str = "6"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="4" w:color="{color}"/>'
        f"</w:pBdr>"
    )
    pPr.append(pBdr)


def _add_page_fields(paragraph):
    def _fld(instr: str):
        r1 = paragraph.add_run()
        _font(r1, size=8.5, color=COR_TEXTO_CINZA)
        fc1 = OxmlElement("w:fldChar")
        fc1.set(qn("w:fldCharType"), "begin")
        r1._r.append(fc1)

        r2 = paragraph.add_run()
        _font(r2, size=8.5, color=COR_TEXTO_CINZA)
        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = instr
        r2._r.append(it)

        r3 = paragraph.add_run()
        _font(r3, size=8.5, color=COR_TEXTO_CINZA)
        fc2 = OxmlElement("w:fldChar")
        fc2.set(qn("w:fldCharType"), "end")
        r3._r.append(fc2)

    r = paragraph.add_run("Página ")
    _font(r, size=8.5, color=COR_TEXTO_CINZA)
    _fld("PAGE")
    r = paragraph.add_run(" de ")
    _font(r, size=8.5, color=COR_TEXTO_CINZA)
    _fld("NUMPAGES")


def new_doc(brasao_bytes: bytes | None, rodape_id: str) -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.4)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)
    sec.header_distance = Cm(0.9)
    sec.footer_distance = Cm(1.0)

    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal._element.get_or_add_rPr()
    normal.font.size = Pt(10.5)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    normal.paragraph_format.space_after = Pt(5)

    # Cabeçalho institucional oficial
    header = sec.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.clear()
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.space_before = Pt(0)
    hp.paragraph_format.space_after = Pt(2)
    if brasao_bytes:
        run_img = hp.add_run()
        run_img.add_picture(io.BytesIO(brasao_bytes), width=Cm(1.55))

    linhas_cab = [
        ("ESTADO DO RIO GRANDE DO SUL", False),
        ("MUNICÍPIO DE PELOTAS", False),
        ("SECRETARIA MUNICIPAL DE SAÚDE", True),
    ]
    for idx, (texto, negrito) in enumerate(linhas_cab):
        p = header.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2 if idx == len(linhas_cab) - 1 else 0)
        r = p.add_run(texto)
        _font(r, size=9.5, bold=negrito)
        if idx == len(linhas_cab) - 1:
            _paragraph_bottom_border(p, color="888888", sz="4")

    # Rodapé institucional com paginação nativa
    footer = sec.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.clear()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(2)
    fp.paragraph_format.space_after = Pt(0)
    r_id = fp.add_run(f"{rodape_id}   |   ")
    _font(r_id, size=8.5, color=COR_TEXTO_CINZA)
    _add_page_fields(fp)

    return doc


def add_title(doc: Document, text: str, subtitle: str | None = None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(4 if subtitle else 10)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    _font(r, size=12, bold=True)
    if subtitle:
        ps = doc.add_paragraph()
        ps.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ps.paragraph_format.space_before = Pt(0)
        ps.paragraph_format.space_after = Pt(10)
        ps.paragraph_format.keep_with_next = True
        rs = ps.add_run(subtitle)
        _font(rs, size=9.5, italic=True, color=COR_TEXTO_CINZA)


def add_h2(doc: Document, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(11)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    _font(r, size=11, bold=True)
    return p


def add_h3(doc: Document, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    _font(r, size=10.5, bold=True)
    return p


def add_p(
    doc: Document,
    text: str,
    *,
    bold_prefix: str | None = None,
    bold: bool = False,
    italic: bool = False,
    justify: bool = True,
    space_after: float = 5,
):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.space_after = Pt(space_after)
    if bold_prefix:
        r1 = para.add_run(bold_prefix)
        _font(r1, bold=True)
        r2 = para.add_run(text)
        _font(r2, bold=bold, italic=italic)
    else:
        r = para.add_run(text)
        _font(r, bold=bold, italic=italic)
    return para


def add_bullet(doc: Document, text: str, *, bold_prefix: str | None = None):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para.paragraph_format.left_indent = Cm(0.6)
    para.paragraph_format.first_line_indent = Cm(-0.3)
    para.paragraph_format.space_after = Pt(3)
    r0 = para.add_run("• ")
    _font(r0, bold=True)
    if bold_prefix:
        r1 = para.add_run(bold_prefix)
        _font(r1, bold=True)
    r2 = para.add_run(text)
    _font(r2)
    return para


def add_checkbox(doc: Document, text: str, *, checked: bool = False):
    para = doc.add_paragraph()
    para.paragraph_format.left_indent = Cm(0.4)
    para.paragraph_format.space_after = Pt(2)
    mark = "X" if checked else " "
    r = para.add_run(f"( {mark} ) {text}")
    _font(r, bold=checked)
    return para


def add_table(
    doc: Document,
    headers: list[str],
    rows: list[list[str]],
    *,
    col_widths_cm: list[float],
    alignments: list[WD_ALIGN_PARAGRAPH] | None = None,
    font_size: float = 8.5,
    last_row_total: bool = False,
):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    _set_table_borders(t)

    # Cabeçalho
    hdr_row = t.rows[0]
    _row_header(hdr_row)
    _row_cant_split(hdr_row)
    for ci, h_text in enumerate(headers):
        cell = hdr_row.cells[ci]
        cell.width = Cm(col_widths_cm[ci])
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _shade_cell(cell, COR_CABECALHO_TABELA)
        _cell_margins(cell, top=80, bottom=80, left=80, right=80)
        p = cell.paragraphs[0]
        p.alignment = (
            alignments[ci]
            if alignments and alignments[ci] == WD_ALIGN_PARAGRAPH.CENTER
            else WD_ALIGN_PARAGRAPH.CENTER
        )
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h_text)
        _font(r, size=font_size, bold=True)

    # Linhas de dados
    for ri, row_data in enumerate(rows, start=1):
        row = t.rows[ri]
        _row_cant_split(row)
        is_total = last_row_total and (ri == len(rows))
        for ci, val in enumerate(row_data):
            cell = row.cells[ci]
            cell.width = Cm(col_widths_cm[ci])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            if is_total:
                _shade_cell(cell, COR_TOTAL_TABELA)
            _cell_margins(cell, top=65, bottom=65, left=80, right=80)
            p = cell.paragraphs[0]
            if alignments:
                p.alignment = alignments[ci]
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(str(val))
            _font(r, size=font_size, bold=is_total)

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(4)
    return t


def add_assinaturas(doc: Document, *, titulo_secao: str, incluir_equipe_texto: str | None = None):
    add_h2(doc, titulo_secao)
    if incluir_equipe_texto:
        add_p(doc, incluir_equipe_texto)
    add_p(
        doc,
        "Certificamos que somos responsáveis pela elaboração do presente instrumento e que o "
        "mesmo está em consonância com a legislação vigente.",
    )

    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    _set_no_borders(t)
    _row_cant_split(t.rows[0])
    for idx, pessoa in enumerate((GESTOR, FISCAL)):
        cell = t.rows[0].cells[idx]
        cell.width = Cm(8.5)
        _cell_margins(cell, top=160, bottom=60, left=60, right=60)
        linhas = [
            "______________________________",
            pessoa[0],
            f"Matrícula: {pessoa[1]}",
            pessoa[2],
        ]
        for j, ln in enumerate(linhas):
            para = cell.paragraphs[0] if j == 0 else cell.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(1)
            para.paragraph_format.keep_with_next = j < len(linhas) - 1
            r = para.add_run(ln)
            _font(r, size=9.5, bold=(j == 1))

    p_data = doc.add_paragraph()
    p_data.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_data.paragraph_format.space_before = Pt(12)
    p_data.paragraph_format.space_after = Pt(14)
    p_data.paragraph_format.keep_with_next = True
    r_data = p_data.add_run("Pelotas/RS, _____ de ______________ de 2026.")
    _font(r_data, size=10.5)

    for idx_ln, (txt, negrito) in enumerate(
        [
            ("__________________________________________", False),
            (SECRETARIA[0], True),
            (f"{SECRETARIA[1]} — Autorizo", False),
        ]
    ):
        pa = doc.add_paragraph()
        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pa.paragraph_format.space_before = Pt(0)
        pa.paragraph_format.space_after = Pt(1)
        pa.paragraph_format.keep_with_next = idx_ln < 2
        ra = pa.add_run(txt)
        _font(ra, size=10, bold=negrito)


# ---------------------------------------------------------------------------
# DOCUMENTO 1: DFD REVISADO
# ---------------------------------------------------------------------------


def build_dfd(brasao_bytes: bytes | None, out_path: Path) -> None:
    doc = new_doc(brasao_bytes, "DFD Revisado — PICSPel (Materiais Permanentes) | Emenda nº 43530013")
    add_title(
        doc,
        "DOCUMENTO DE FORMALIZAÇÃO DA DEMANDA nº [___]/2026 – PICSPel\n(MATERIAIS PERMANENTES E EQUIPAMENTOS)",
        "Processo Administrativo / Memorando nº [______]/2026  •  Emenda Parlamentar nº 43530013",
    )

    add_p(
        doc,
        "Secretaria Municipal de Saúde – SMS / Diretoria de Atenção Primária / Programa Municipal "
        "de Práticas Integrativas e Complementares em Saúde (PICSPel).",
        bold_prefix="UNIDADE DEMANDANTE: ",
    )

    # 1 OBJETO
    add_h2(doc, "1 OBJETO")
    add_p(
        doc,
        "Aquisição pontual de equipamentos e materiais permanentes — compreendendo 5 (cinco) itens "
        "e 49 (quarenta e nove) unidades físicas (caixa de som portátil recarregável, garrafa térmica "
        "em aço inoxidável, maleta organizadora rígida em ABS, jarra elétrica em aço inoxidável e "
        "plastificadora/laminadora) — destinados à estruturação e ao suporte das ações do Programa "
        "Municipal de Práticas Integrativas e Complementares em Saúde (PICSPel), para uso pelas "
        "08 (oito) equipes eMulti (em suporte às Unidades Básicas de Saúde de referência) e pela "
        "coordenação do programa, com recursos da Emenda Parlamentar nº 43530013.",
    )

    add_h3(doc, "1.1 TIPO DE OBJETO")
    add_checkbox(doc, "Material de Consumo", checked=False)
    add_checkbox(doc, "Material Permanente / Equipamento", checked=True)
    add_checkbox(doc, "Serviço continuado com dedicação exclusiva de mão de obra", checked=False)
    add_checkbox(doc, "Serviço continuado sem dedicação exclusiva de mão de obra", checked=False)
    add_checkbox(doc, "Serviço não continuado", checked=False)
    add_p(
        doc,
        "o objeto é composto exclusivamente por materiais permanentes e equipamentos de apoio às "
        "práticas corporais, integrativas, educativas e organizacionais do PICSPel (classificação "
        "contábil 4.4.90.52). As demais parcelas da Emenda Parlamentar nº 43530013 — Processo 1 "
        "(Materiais de Consumo — 3.3.90.30) e Processo 3 (Serviços de Capacitação em Auriculoterapia "
        "— 3.3.90.39) — tramitam em DFDs e processos autônomos, sem sobreposição de itens.",
        bold_prefix="Observação: ",
    )

    # 2 FORMA DE CONTRATAÇÃO SUGERIDA
    add_h2(doc, "2 FORMA DE CONTRATAÇÃO SUGERIDA")
    add_checkbox(doc, "Pregão Eletrônico", checked=True)
    add_checkbox(doc, "Pregão Eletrônico com Sistema/Ata de Registro de Preços", checked=False)
    add_checkbox(doc, "Dispensa com Disputa Eletrônica", checked=False)
    add_checkbox(doc, "Adesão à Ata do Consórcio Público do Extremo Sul", checked=False)
    add_checkbox(doc, "Adesão à Ata de Registro de Preços de Outro(s) Órgão(s)", checked=False)
    add_p(
        doc,
        "menor preço por item, com adjudicação por item e participação exclusiva para Microempresas "
        "(ME) e Empresas de Pequeno Porte (EPP), nos termos do art. 48, inciso I, da Lei Complementar "
        "nº 123/2006 e do art. 4º da Lei Federal nº 14.133/2021.",
        bold_prefix="Critério de julgamento sugerido: ",
    )
    add_p(
        doc,
        "embora o valor estimado total (R$ 11.600,00) situe-se abaixo do limite do art. 75, inciso II, "
        "da Lei nº 14.133/2021, sugere-se o Pregão Eletrônico como via prioritária para prevenir "
        "eventual fracionamento indevido de despesa no exercício (art. 75, § 1º), ampliar a "
        "competitividade e manter padronização com os demais processos da emenda — ressalvada a "
        "possibilidade de utilização de Dispensa Eletrônica com disputa caso a Contabilidade/Compras "
        "certifique formalmente a não extrapolação do limite anual por ramo de atividade, ou adesão "
        "a ata vigente comprovadamente mais vantajosa.",
        bold_prefix="Nota técnica sobre a modalidade: ",
    )

    # 3 JUSTIFICATIVA DA CONTRATAÇÃO
    add_h2(doc, "3 JUSTIFICATIVA DA CONTRATAÇÃO")
    add_p(
        doc,
        "A formalização desta demanda destina-se a viabilizar a estruturação física e logística da "
        "oferta de Práticas Integrativas e Complementares em Saúde (PICS) na Atenção Primária do "
        "Município de Pelotas/RS, em conformidade com a Política Nacional de Práticas Integrativas e "
        "Complementares (PNPIC) e com o planejamento do PICSPel, utilizando recursos da Emenda "
        "Parlamentar nº 43530013.",
    )
    add_p(
        doc,
        "Os equipamentos e materiais permanentes solicitados são indispensáveis para estruturar as "
        "atividades coletivas e terapêuticas executadas pelas 08 (oito) equipes eMulti (Centro/Porto, "
        "Fragata 1, Fragata 2, Três Vendas 1, Três Vendas 2, Areal, Colônia e Praias) — provendo "
        "sonorização portátil para práticas corporais e meditação, aquecimento e conservação térmica "
        "de água para escalda-pés e chás terapêuticos e transporte protegido dos kits de insumos —, "
        "bem como para dotar a coordenação do programa de equipamento para confecção e conservação "
        "de materiais educativos e protocolos (plastificadora). A aquisição por Pregão Eletrônico com "
        "julgamento por item assegura competitividade, publicidade no PNCP, isonomia e economicidade.",
    )
    add_p(
        doc,
        "[INFORMAR se a contratação consta do Plano de Contratações Anual (PCA) e o número do item; "
        "caso não conste, registrar justificativa de inclusão excepcional vinculada à Emenda "
        "Parlamentar nº 43530013].",
        bold_prefix="Previsão no Plano de Contratações Anual (PCA): ",
    )

    # 4 DESCRIÇÃO DO OBJETO
    add_h2(doc, "4 DESCRIÇÃO DO OBJETO E ESTIMATIVA PRELIMINAR DE VALORES")
    rows_dfd = [
        [
            f"{it.numero:02d}",
            f"{it.nome_curto}: {it.especificacao}",
            it.unidade,
            str(it.quantidade),
            brl(it.valor_unitario),
            brl(it.valor_total),
        ]
        for it in ITENS
    ]
    rows_dfd.append(
        [
            "",
            "VALOR TOTAL ESTIMADO (5 itens / 49 unidades)",
            "",
            str(TOTAL_UNIDADES),
            "",
            brl(VALOR_TOTAL),
        ]
    )
    add_table(
        doc,
        ["Item", "Especificação Técnica", "Unid.", "Qtd.", "Valor unit. (R$)", "Valor total (R$)"],
        rows_dfd,
        col_widths_cm=[1.0, 9.6, 1.1, 1.1, 2.1, 2.1],
        alignments=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ],
        font_size=8.0,
        last_row_total=True,
    )
    add_p(
        doc,
        f"{brl(VALOR_TOTAL)} (onze mil e seiscentos reais), correspondente à soma integral dos 5 "
        f"(cinco) itens permanentes ({brl(SUBTOTAL_EMULTI)} destinados às 08 equipes eMulti — "
        "R$ 1.400,00 por equipe — e R$ 400,00 destinados à coordenação do PICSPel). Os valores "
        "unitários são referenciais preliminares e deverão ser confirmados/atualizados mediante "
        "pesquisa de preços formal (art. 23 da Lei nº 14.133/2021).",
        bold_prefix="Valor total estimado: ",
    )

    # 5 DOTAÇÃO ORÇAMENTÁRIA
    add_h2(doc, "5 DOTAÇÃO ORÇAMENTÁRIA")
    rows_dot = [
        ["Emenda Parlamentar", f"Emenda Parlamentar nº {EMENDA} (valor global: {brl(RECURSO_GLOBAL_EMENDA)})"],
        ["Processo Administrativo / Memorando", "[INFORMAR Nº DO PROCESSO / MEMORANDO]/2026"],
        ["Projeto/Atividade", "[INFORMAR — Atenção Primária / PICSPel]"],
        ["Fonte de Recurso", "[INFORMAR — vinculada à Emenda Parlamentar nº 43530013]"],
        ["Dotação / Reduzido", "[INFORMAR — código de dotação/reduzido orçamentário]"],
        [
            "Classificação da Despesa",
            "4.4.90.52 (Equipamentos e Material Permanente — Despesa de Capital / GND 4). "
            "ATENÇÃO: validar previamente junto à Contabilidade/CCGOF a compatibilidade de "
            "empenho em GND 4 na programação da Emenda nº 43530013 (ou fonte complementar) "
            "e o enquadramento patrimonial municipal de cada item.",
        ],
        ["Valor Total Estimado", f"{brl(VALOR_TOTAL)} (onze mil e seiscentos reais)"],
        [
            "Enquadramento no Rateio da Emenda",
            "Parcela exclusiva de materiais permanentes (5 itens / 49 un.), autônoma em relação "
            "aos processos de Materiais de Consumo (3.3.90.30) e Serviços de Capacitação "
            "(3.3.90.39), observando-se o limite global de R$ 100.000,00 da emenda.",
        ],
    ]
    add_table(
        doc,
        ["Campo", "Informação"],
        rows_dot,
        col_widths_cm=[5.2, 11.8],
        alignments=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
        font_size=8.5,
    )

    # 6 LOCAL DE ENTREGA
    add_h2(doc, "6 LOCAL DE ENTREGA E CONTROLE PATRIMONIAL")
    add_p(
        doc,
        f"Os bens deverão ser entregues no {ENDERECO_ALMOXARIFADO}, ou em outro local dentro do "
        "Município de Pelotas/RS formalmente indicado pela Diretoria de Atenção Primária / "
        "coordenação do PICSPel na Ordem de Fornecimento.",
    )
    add_p(
        doc,
        "Após o recebimento provisório, a conferência técnica (incluindo verificação de tensão "
        "elétrica 220 V e teste funcional dos eletroportáteis) e o recebimento definitivo, os bens "
        "permanentes deverão passar por tombamento/incorporação patrimonial junto ao setor "
        "competente da SMS/Prefeitura antes da distribuição, mediante Termo de Responsabilidade, "
        "às 08 (oito) equipes eMulti (Centro/Porto, Fragata 1, Fragata 2, Três Vendas 1, Três "
        "Vendas 2, Areal, Colônia e Praias) e à coordenação do PICSPel.",
    )

    # 7 PREVISÃO DE ENTREGA, GARANTIA E FLUXO PÓS-ENTREGA
    add_h2(doc, "7 PREVISÃO DE ENTREGA, GARANTIA E FLUXO PÓS-ENTREGA")
    add_p(
        doc,
        "Entrega integral por item em até 30 (trinta) dias corridos, contados do recebimento da "
        "Nota de Empenho / Ordem de Fornecimento. Todos os bens deverão possuir garantia mínima de "
        "12 (doze) meses a contar do recebimento definitivo (garantia legal do CDC acrescida de "
        "garantia contratual complementar) e assistência técnica autorizada no Brasil. Após a "
        "chegada e incorporação patrimonial, será realizada reunião com as equipes eMulti e com "
        "as equipes da APS de referência para entrega formal, orientação de guarda, controle "
        "patrimonial e fluxos de utilização.",
    )

    add_assinaturas(doc, titulo_secao="EQUIPE DE PLANEJAMENTO / FISCALIZAÇÃO")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))


# ---------------------------------------------------------------------------
# DOCUMENTO 2: ETP REVISADO
# ---------------------------------------------------------------------------


def build_etp(brasao_bytes: bytes | None, out_path: Path) -> None:
    doc = new_doc(brasao_bytes, "ETP Revisado — PICSPel (Materiais Permanentes) | Emenda nº 43530013")
    add_title(
        doc,
        "ESTUDO TÉCNICO PRELIMINAR – PICSPel\n(MATERIAIS PERMANENTES E EQUIPAMENTOS)",
        "Lei Federal nº 14.133/2021 (art. 18, § 1º)  •  Emenda Parlamentar nº 43530013",
    )

    # 1 IDENTIFICAÇÃO DO PROCESSO
    add_h2(doc, "1 IDENTIFICAÇÃO DO PROCESSO")
    add_p(
        doc,
        "Secretaria Municipal de Saúde – Diretoria de Atenção Primária – Programa Municipal de "
        "Práticas Integrativas e Complementares em Saúde (PICSPel).",
        bold_prefix="1.1 Unidade responsável pela contratação: ",
    )
    add_p(
        doc,
        "Aquisição pontual de materiais de natureza permanente e equipamentos destinados à "
        "implementação e ao fortalecimento das Práticas Integrativas e Complementares em Saúde "
        "(PICS) na rede de Atenção Primária à Saúde do Município de Pelotas/RS, compreendendo "
        "5 (cinco) itens e 49 (quarenta e nove) unidades físicas (caixa de som portátil "
        "recarregável, garrafa térmica em aço inoxidável, maleta organizadora rígida em ABS, "
        "jarra elétrica em aço inoxidável e plastificadora/laminadora), para uso pelas 08 (oito) "
        "equipes eMulti e pela coordenação do programa, no âmbito da Emenda Parlamentar nº 43530013.",
        bold_prefix="1.2 Objeto da contratação: ",
    )
    add_p(
        doc,
        "Os equipamentos e materiais permanentes destinam-se ao suporte logístico, operacional e "
        "estrutural das práticas corporais, meditação, escalda-pés, rotinas organizacionais e "
        "atividades educativas, contribuindo para a ampliação da oferta de PICS, a humanização do "
        "cuidado e o fortalecimento da Atenção Primária à Saúde.",
    )
    add_p(
        doc,
        "Processo Administrativo / Memorando nº [______]/2026 – DFD nº [___]/2026 – Emenda "
        "Parlamentar nº 43530013.",
        bold_prefix="1.3 Número do processo: ",
    )
    add_p(
        doc,
        f"Emenda Parlamentar nº {EMENDA}, no valor global de {brl(RECURSO_GLOBAL_EMENDA)}. O "
        "presente ETP abrange exclusivamente a parcela correspondente aos materiais de natureza "
        f"permanente e equipamentos (estimativa de {brl(VALOR_TOTAL)} — 5 itens / 49 unidades, "
        "classificação 4.4.90.52). As demais parcelas da mesma emenda — Processo 1 (Materiais de "
        "Consumo — 3.3.90.30) e Processo 3 (Serviços de Capacitação em Auriculoterapia — "
        f"3.3.90.39, estimado em {brl(RECURSO_COTA_SERVICOS)}) — tramitam em processos autônomos, "
        "sujeita a presente parcela à validação contábil prévia pela CCGOF quanto à compatibilidade "
        "do Grupo de Natureza de Despesa (GND 4 – Investimentos) na programação da emenda.",
        bold_prefix="1.4 Fonte de recurso: ",
    )

    # 2 DESCRIÇÃO DA NECESSIDADE
    add_h2(doc, "2 DESCRIÇÃO DA NECESSIDADE")
    add_p(
        doc,
        "A Política Nacional de Práticas Integrativas e Complementares (PNPIC) e a organização "
        "local do PICSPel demandam equipamentos e bens permanentes específicos para estruturar a "
        "infraestrutura das atividades de PICS na Atenção Primária. A Secretaria Municipal de Saúde "
        "necessita adquirir o conjunto de materiais permanentes descritos na Lista de Materiais "
        "PICSPel, de modo a municiar as equipes eMulti e a coordenação do programa.",
    )
    add_p(
        doc,
        "A ausência desses equipamentos limita o suporte estrutural e logístico às atividades "
        "coletivas, terapêuticas e organizacionais (sonorização em grupos de práticas corporais e "
        "meditação, aquecimento e conservação térmica de água para escalda-pés e chás terapêuticos, "
        "transporte protegido de insumos entre UBS e produção de guias educativos plastificados). "
        "A demanda foi dimensionada a partir do planejamento do PICSPel, considerando as 08 (oito) "
        "equipes eMulti contempladas e a necessidade de estruturação da coordenação do programa.",
    )
    add_p(
        doc,
        "Equipes eMulti contempladas (quantitativos por equipe multiplicados por 8):",
        bold=True,
        space_after=3,
    )
    for eq in EQUIPES_EMULTI:
        add_bullet(doc, eq)
    add_bullet(
        doc,
        "Eventual inclusão futura de até 4 (quatro) equipes adicionais a definir não altera "
        "automaticamente os quantitativos deste ETP; eventual suplementação dependerá de novo "
        "planejamento e disponibilidade orçamentária.",
    )
    add_p(
        doc,
        "Após a entrega, conferência técnica e incorporação patrimonial dos equipamentos, será "
        "realizada reunião com as equipes eMulti e com as equipes da APS de referência para "
        "organizar guarda, termo de responsabilidade patrimonial, controle e fluxos de utilização, "
        "conforme encaminhamento do Programa PICSPel.",
    )

    # 3 DA CONTRATAÇÃO E DO PLANEJAMENTO
    add_h2(doc, "3 DA CONTRATAÇÃO E DO PLANEJAMENTO")
    add_p(
        doc,
        "A contratação tem por finalidade a aquisição pontual dos equipamentos e materiais "
        "permanentes necessários à estruturação do PICSPel com recursos da Emenda Parlamentar nº "
        "43530013. Os itens possuem padrões de desempenho e qualidade objetivamente definíveis por "
        "especificações usuais de mercado, enquadrando-se como bens comuns (art. 6º, inciso XIII, "
        "da Lei Federal nº 14.133/2021) e não se caracterizando como artigos de luxo (art. 20 da "
        "mesma Lei).",
    )
    add_p(
        doc,
        "Pregão Eletrônico (art. 28, inciso I, da Lei nº 14.133/2021), critério de julgamento "
        "menor preço por item, com adjudicação por item e participação exclusiva para Microempresas "
        "(ME) e Empresas de Pequeno Porte (EPP), nos termos do art. 48, inciso I, da Lei "
        "Complementar nº 123/2006 e do art. 4º da Lei nº 14.133/2021 (haja vista que todos os "
        "itens têm valor estimado individual inferior a R$ 80.000,00), observando-se os princípios "
        "da legalidade, isonomia, competitividade, economicidade, planejamento e seleção da "
        "proposta mais vantajosa.",
        bold_prefix="Forma de contratação sugerida: ",
    )
    add_p(
        doc,
        "Embora o valor global estimado desta aquisição (R$ 11.600,00) seja inferior ao limite "
        "previsto no art. 75, inciso II, da Lei nº 14.133/2021 (atualizado anualmente por decreto "
        "federal), recomenda-se o Pregão Eletrônico por: (i) eliminar qualquer risco de "
        "fracionamento indevido de despesa no exercício financeiro no mesmo ramo de atividade "
        "(art. 75, § 1º); (ii) ampliar a disputa entre fornecedores especializados; e (iii) manter "
        "padronização procedimental com os demais processos da Emenda nº 43530013. Caso a "
        "Contabilidade/Setor de Compras certifique formalmente a inexistência de extrapolação do "
        "limite anual por ramo de atividade, a autoridade competente poderá avaliar, "
        "motivadamente, a adoção de Dispensa Eletrônica com disputa.",
        bold_prefix="Enquadramento frente ao art. 75, II, da Lei nº 14.133/2021: ",
    )
    add_p(
        doc,
        "Justifica-se a aquisição pontual (e não o Sistema de Registro de Preços) porque: (i) os "
        "quantitativos estão vinculados a recurso de emenda parlamentar com objeto e prazo "
        "definidos; (ii) a demanda de equipamentos foi previamente dimensionada em quantidade "
        "certa para as 08 equipes eMulti e para a coordenação do programa; e (iii) não se trata "
        "de demanda contínua indeterminada de reposição, e sim de implantação/estruturação "
        "patrimonial com quantitativo exato.",
    )
    add_p(
        doc,
        "[INFORMAR se a contratação consta do PCA e o respectivo número do item — caso não conste, "
        "justificar inclusão excepcional vinculada à Emenda Parlamentar nº 43530013].",
        bold_prefix="Previsão no Plano de Contratações Anual (PCA): ",
    )
    add_p(
        doc,
        "Serão observadas todas as etapas da Lei nº 14.133/2021, incluindo o DFD, o presente "
        "Estudo Técnico Preliminar, a pesquisa formal de preços (art. 23), o Termo de Referência, "
        "a confirmação orçamentária e de GND pela Contabilidade/CCGOF, a designação de gestor e "
        "fiscal e a divulgação no Portal Nacional de Contratações Públicas (PNCP).",
    )

    # 4 REQUISITOS DA CONTRATAÇÃO
    add_h2(doc, "4 REQUISITOS DA CONTRATAÇÃO")
    add_p(
        doc,
        "Os bens deverão atender integralmente às especificações técnicas de cada item, ser novos, "
        "de primeiro uso, de linha de fabricação corrente, livres de vícios ou defeitos, entregues "
        "em embalagem original lacrada e íntegra, acompanhados de manual de instruções em língua "
        "portuguesa, acessórios e cabos completos, número de série (quando aplicável) e termo ou "
        "certificado de garantia do fabricante/fornecedor. Os equipamentos elétricos deverão ser "
        "compatíveis com a rede elétrica padrão de Pelotas/RS — 220 V (60 Hz) ou bivolt automático "
        "— e possuir plugue padrão brasileiro (ABNT NBR 14136).",
    )
    add_p(doc, "Requisitos específicos por família de produtos:", bold=True, space_after=3)
    add_bullet(
        doc,
        "caixa de som portátil com sistema estéreo (dois alto-falantes ou transdutores "
        "equivalentes), potência mínima de 20 W RMS, conectividade Bluetooth (v5.0 ou superior), "
        "bateria recarregável com autonomia mínima de 6 horas, carregamento compatível com 220 V / "
        "bivolt (ou USB/USB-C) e certificado de homologação na ANATEL.",
        bold_prefix="Equipamentos portáteis de áudio (Item 01): ",
    )
    add_bullet(
        doc,
        "garrafas térmicas 100% em aço inoxidável de grau alimentício (sem ampola de vidro), "
        "isolamento a vácuo, capacidade aproximada de 1,8 L (1,7 L a 2,0 L), bombeamento por "
        "pressão com vedação hermética e desempenho térmico mínimo de 6 horas; e jarras elétricas "
        "em aço inoxidável (1,7 L a 2,2 L), potência mínima de 1.500 W (referência 1.800 W), "
        "tensão 220 V, base 360°, desligamento automático, proteção contra funcionamento a seco e "
        "certificação compulsória de segurança elétrica no INMETRO (Portaria nº 148/2022 ou "
        "substitutiva).",
        bold_prefix="Equipamentos e utensílios térmicos (Itens 02 e 04): ",
    )
    add_bullet(
        doc,
        "maletas rígidas tipo frasqueira em polímero ABS (100% ABS rígido ou equivalente), com "
        "abertura horizontal por zíper duplo, alça reforçada e dimensões proporcionais ao "
        "transporte seguro dos kits de insumos pelas equipes eMulti.",
        bold_prefix="Maletas organizadoras (Item 03): ",
    )
    add_bullet(
        doc,
        "plastificadora/laminadora (polisseladora) de mesa compatível com formatos A3, A4, A5 e "
        "A6 (largura útil mínima de 320 mm), tensão 220 V (ou bivolt), sistema anti-atolamento e "
        "compatibilidade expressa com filmes/pouches de 0,05 mm a 0,07 mm (125 a 175 micras), "
        "garantindo plena integração com os plásticos A4 de 0,07 mm adquiridos no processo "
        "correlato de materiais de consumo.",
        bold_prefix="Apoio administrativo e educativo (Item 05): ",
    )
    add_p(
        doc,
        "mínimo de 12 (doze) meses para todos os itens (compreendendo a garantia legal de 90 dias "
        "prevista no art. 26, inciso II, do Código de Defesa do Consumidor acrescida de garantia "
        "contratual complementar de 9 meses), contados a partir da data do recebimento definitivo, "
        "com cobertura de assistência técnica autorizada no território nacional e sem ônus de "
        "transporte/coleta para a Administração durante o período de garantia. Em caso de defeito "
        "ou desconformidade constatada no recebimento, a contratada responderá pela substituição, "
        "sem ônus, no prazo máximo de 5 (cinco) dias úteis do recebimento da notificação.",
        bold_prefix="Garantia mínima e assistência técnica: ",
    )
    add_p(
        doc,
        "Não haverá exigência ou direcionamento de marca/modelo específico (art. 41, inciso I, da "
        "Lei nº 14.133/2021). Para verificação objetiva da conformidade técnica, será exigida "
        "ficha técnica/catálogo do fabricante junto à proposta classificada provisoriamente em "
        "primeiro lugar, bem como comprovação de certificação INMETRO (item 04 e demais itens "
        "abrangidos) e homologação ANATEL (item 01). Não se exige registro/AFE na ANVISA por não "
        "se tratarem de produtos para saúde/correlatos médicos.",
    )
    add_p(
        doc,
        "habilitação jurídica; regularidade fiscal, social e trabalhista (Fazendas Federal, "
        "Estadual e Municipal, FGTS, CNDT e declarações legais do art. 7º, XXXIII, da CF e art. "
        "63, IV, da Lei nº 14.133/2021); e qualificação econômico-financeira proporcional ao valor "
        "e à baixa complexidade da aquisição, vedadas exigências restritivas ou desarrazoadas.",
        bold_prefix="Habilitação: ",
    )

    # 5 ESTIMATIVA DAS QUANTIDADES
    add_h2(doc, "5 ESTIMATIVA DAS QUANTIDADES E ESPECIFICAÇÕES TÉCNICAS")
    add_p(
        doc,
        "As quantidades foram estimadas com base na Lista de Materiais para Aquisição do PICSPel e "
        "no DFD correspondente, considerando: (a) 08 (oito) equipes eMulti (1 caixa de som portátil, "
        "3 garrafas térmicas em inox, 1 maleta organizadora rígida e 1 jarra elétrica por equipe, "
        "totalizando 48 unidades); e (b) 1 (uma) plastificadora/laminadora para suporte "
        "administrativo e educativo da coordenação do programa. Por se tratar de aquisição pontual "
        "com quantitativos certos, a Administração adquirirá os totais adjudicados, observados o "
        "empenho e a disponibilidade orçamentária da emenda.",
    )

    rows_itens = [
        [
            f"{it.numero:02d}",
            f"{it.nome_curto}: {it.especificacao}",
            it.unidade,
            str(it.quantidade),
            brl(it.valor_unitario),
            brl(it.valor_total),
        ]
        for it in ITENS
    ]
    rows_itens.append(
        [
            "",
            "VALOR TOTAL ESTIMADO (5 itens / 49 unidades)",
            "",
            str(TOTAL_UNIDADES),
            "",
            brl(VALOR_TOTAL),
        ]
    )
    add_table(
        doc,
        ["Item", "Descrição / Especificação Técnica", "Unid.", "Qtd.", "Valor unit. (R$)", "Valor total (R$)"],
        rows_itens,
        col_widths_cm=[1.0, 9.6, 1.1, 1.1, 2.1, 2.1],
        alignments=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ],
        font_size=8.0,
        last_row_total=True,
    )

    add_p(
        doc,
        "Classificação por natureza da despesa e destinação (para fins de empenho e patrimônio):",
        bold=True,
        space_after=4,
    )
    rows_nat = [
        [
            f"{it.numero:02d}",
            "4.4.90.52 (Permanente)",
            it.destinacao,
            it.certificacao_principal,
            brl(it.valor_total),
        ]
        for it in ITENS
    ]
    rows_nat.append(
        [
            "",
            "Subtotal — Equipes eMulti (Itens 01 a 04)",
            "48 un. (R$ 1.400,00 / equipe)",
            "",
            brl(SUBTOTAL_EMULTI),
        ]
    )
    rows_nat.append(
        [
            "",
            "Subtotal — Coordenação PICSPel (Item 05)",
            "1 un.",
            "",
            brl(SUBTOTAL_COORD),
        ]
    )
    rows_nat.append(
        [
            "",
            "TOTAL GERAL — MATERIAL PERMANENTE",
            f"{TOTAL_UNIDADES} unidades",
            "",
            brl(VALOR_TOTAL),
        ]
    )
    add_table(
        doc,
        ["Item", "Natureza / Elemento", "Destinação", "Conformidade / Certificação", "Valor (R$)"],
        rows_nat,
        col_widths_cm=[1.0, 3.8, 3.8, 6.0, 2.4],
        alignments=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ],
        font_size=8.0,
        last_row_total=True,
    )
    add_p(
        doc,
        "a Contabilidade / CCGOF deverá validar previamente: (i) o enquadramento orçamentário "
        "dos 5 itens no elemento de despesa 4.4.90.52 (Equipamentos e Material Permanente — "
        "Despesa de Capital / GND 4), segundo o MCASP e a Portaria STN nº 448/2002; (ii) a "
        "compatibilidade de execução em GND 4 dentro da programação da Emenda Parlamentar nº "
        "43530013 (ou eventual indicação de fonte/readequação orçamentária); e (iii) os critérios "
        "da norma patrimonial municipal para tombamento e controle individualizado dos bens "
        "(inclusive jarras elétricas e garrafas térmicas).",
        bold_prefix="Observação contábil e patrimonial: ",
    )

    # 6 LEVANTAMENTO DE MERCADO
    add_h2(doc, "6 LEVANTAMENTO DE MERCADO")
    add_p(
        doc,
        "Foi realizado levantamento preliminar de mercado com base na Lista de Materiais PICSPel "
        f"e no DFD de Materiais Permanentes, totalizando a estimativa referencial de {brl(VALOR_TOTAL)} "
        "(onze mil e seiscentos reais) para os 5 itens permanentes. Verificou-se ampla oferta de "
        "fabricantes, marcas e fornecedores locais, regionais e nacionais (eletroportáteis, "
        "equipamentos de escritório, utilidades em aço inox e artigos de viagem/organização) "
        "aptos a atender às especificações sem restrição indevida à competitividade.",
    )
    add_p(doc, "Alternativas consideradas:", bold=True, space_after=3)
    add_bullet(
        doc,
        "(I) Licitação própria por Pregão Eletrônico com julgamento por menor preço por item e "
        "participação exclusiva para ME/EPP (art. 48, I, da LC nº 123/2006);",
    )
    add_bullet(
        doc,
        "(II) Adesão a ata de registro de preços vigente de outro órgão público ou do Consórcio "
        "Público do Extremo Sul, desde que comprovada a vantajosidade técnica e econômica;",
    )
    add_bullet(
        doc,
        "(III) Contratação direta por Dispensa Eletrônica com disputa (art. 75, inciso II, da Lei "
        "nº 14.133/2021), exclusivamente caso a Contabilidade e o Setor de Compras certifiquem que "
        "o somatório das despesas do exercício no mesmo ramo de atividade (art. 75, § 1º) não "
        "ultrapassa o teto legal vigente.",
    )
    add_p(
        doc,
        "Conclui-se que o Pregão Eletrônico com julgamento por menor preço por item é a solução "
        "mais segura e adequada como regra geral, por ampliar a competitividade (permitindo que "
        "fornecedores especializados disputem apenas os itens de seu ramo), eliminar o risco de "
        "fracionamento indevido de despesa e assegurar plena transparência no PNCP. A existência "
        "de atas vigentes compatíveis deverá ser verificada previamente à publicação do edital; "
        "caso se comprove vantajosidade para item específico, este poderá ser excluído do certame "
        "com justificativa nos autos.",
    )
    add_p(
        doc,
        "A pesquisa formal de preços será consolidada nos termos do art. 23 da Lei nº 14.133/2021 "
        "(Painel de Preços / PNCP, contratações similares de outros entes públicos, mídia/sítios "
        "especializados e/ou pesquisa direta com fornecedores), adotando-se critério estatístico "
        "justificado (mediana ou média saneada), com mapa comparativo anexo ao processo.",
    )

    # 7 ESTIMATIVA DO VALOR DA CONTRATAÇÃO
    add_h2(doc, "7 ESTIMATIVA DO VALOR DA CONTRATAÇÃO")
    add_p(
        doc,
        f"O valor total estimado para a aquisição dos 5 (cinco) itens permanentes (49 unidades) é "
        f"de {brl(VALOR_TOTAL)} (onze mil e seiscentos reais), sendo {brl(SUBTOTAL_EMULTI)} "
        "correspondentes aos equipamentos das 08 (oito) equipes eMulti (R$ 1.400,00 por equipe) e "
        f"{brl(SUBTOTAL_COORD)} correspondentes à plastificadora da coordenação do PICSPel. O "
        f"montante é plenamente compatível com o recurso global de {brl(RECURSO_GLOBAL_EMENDA)} da "
        f"Emenda Parlamentar nº {EMENDA}.",
    )
    add_p(
        doc,
        "Os valores unitários lançados neste ETP são referenciais preliminares e deverão ser "
        "validados e/ou atualizados pela pesquisa de preços oficial (art. 23 da Lei nº 14.133/2021) "
        "antes da divulgação do edital. Nota de revisão aritmética: corrigiu-se a soma global "
        "constante da versão preliminar (R$ 11.200,00), que omitia os R$ 400,00 do Item 05 "
        "(plastificadora) no total consolidado.",
    )

    # 8 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO
    add_h2(doc, "8 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO")
    add_p(
        doc,
        "A solução consiste na aquisição pontual, mediante Pregão Eletrônico com julgamento por "
        "item, dos 5 (cinco) itens permanentes necessários à estruturação do PICSPel, com entrega "
        f"no {ENDERECO_ALMOXARIFADO} (ou local indicado pela Diretoria de Atenção Primária / "
        "coordenação do PICSPel), seguida de conferência técnica e funcional (inclusive verificação "
        "de tensão 220 V), recebimento provisório e definitivo, tombamento/incorporação patrimonial "
        "e distribuição mediante Termo de Responsabilidade às 08 equipes eMulti e à coordenação.",
    )
    add_p(
        doc,
        "Considerado o ciclo de vida do objeto (art. 11, I, e art. 18, § 1º, VII, da Lei nº "
        "14.133/2021), a solução abrange: planejamento e formalização da demanda; seleção do "
        "fornecedor; entrega e teste funcional; recebimento provisório e definitivo; tombamento e "
        "carga patrimonial; guarda e distribuição às equipes; operação assistencial, educativa e "
        "administrativa; cobertura de garantia mínima de 12 (doze) meses e assistência técnica "
        "autorizada; manutenção preventiva/corretiva ao longo da vida útil; e desfazimento/descarte "
        "ambientalmente adequado de resíduos eletroeletrônicos e baterias ao fim da vida útil.",
    )

    # 9 JUSTIFICATIVA PARA PARCELAMENTO
    add_h2(doc, "9 JUSTIFICATIVA PARA PARCELAMENTO")
    add_p(
        doc,
        "A contratação será realizada com parcelamento do objeto por itens e julgamento por menor "
        "preço por item, nos termos do art. 40, inciso V, alínea 'b', e § 2º, da Lei nº 14.133/2021, "
        "por ser tecnicamente viável e economicamente vantajoso. Os 5 itens pertencem a segmentos "
        "mercadológicos distintos (eletrônicos de áudio, utilidades domésticas/térmicas em inox, "
        "artigos de viagem/maletas em ABS, eletroportáteis de cozinha e equipamentos de escritório), "
        "inexistindo dependência técnica de integração que exija fornecedor único. A divisão por "
        "item amplia a participação de micro e pequenas empresas especializadas e reduz o risco de "
        "intermediação onerosa.",
    )

    # 10 DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS
    add_h2(doc, "10 DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS")
    add_p(doc, "Com a contratação, pretende-se alcançar os seguintes resultados:", space_after=3)
    add_bullet(
        doc,
        "estruturar o suporte físico e logístico para a oferta regular de PICS nas 08 (oito) "
        "equipes eMulti contempladas e na coordenação do programa;",
    )
    add_bullet(
        doc,
        "padronizar o conjunto de equipamentos permanentes entre as 08 equipes eMulti "
        "(1 caixa de som, 3 garrafas térmicas, 1 maleta rígida e 1 jarra elétrica por equipe);",
    )
    add_bullet(
        doc,
        "viabilizar a confecção e durabilidade de materiais educativos, cartazes e guias de "
        "orientação de PICS pela coordenação do programa (plastificadora A3/A4/A5/A6);",
    )
    add_bullet(
        doc,
        "fortalecer a atenção integral, a resolutividade e a humanização do cuidado na APS;",
    )
    add_bullet(
        doc,
        "aplicar de forma eficiente, rastreável e transparente os recursos da Emenda Parlamentar "
        f"nº {EMENDA}, assegurando controle patrimonial adequado dos bens adquiridos;",
    )
    add_bullet(
        doc,
        "obter propostas economicamente vantajosas por meio da disputa competitiva em Pregão "
        "Eletrônico.",
    )

    # 11 PROVIDÊNCIAS PRÉVIAS AO CONTRATO
    add_h2(doc, "11 PROVIDÊNCIAS PRÉVIAS AO CONTRATO")
    for prov in [
        "Formalização do DFD revisado, numeração administrativa e aprovação pela autoridade competente;",
        "validação prévia junto à Contabilidade/CCGOF da dotação orçamentária, fonte de recurso e compatibilidade do elemento 4.4.90.52 (GND 4 – Capital) na Emenda Parlamentar nº 43530013;",
        "aprovação do Termo de Referência (com confirmação dos códigos CATMAT no Compras.gov.br);",
        "realização de pesquisa formal de preços (art. 23 da Lei nº 14.133/2021) e elaboração do mapa comparativo;",
        "verificação de atas de registro de preços vigentes (Consórcio Público do Extremo Sul / outros órgãos) para eventual adesão comprovadamente mais vantajosa;",
        "formalização da designação de gestor e fiscal da contratação (art. 117 da Lei nº 14.133/2021);",
        "articulação prévia com o Almoxarifado da SMS e o setor de Patrimônio para recebimento, teste funcional em 220 V, tombamento e emissão dos Termos de Responsabilidade;",
        "planejamento da reunião pós-entrega com as equipes eMulti e APS de referência para distribuição e controle patrimonial.",
    ]:
        add_bullet(doc, prov)

    # 12 CONTRATAÇÕES CORRELATAS/INTERDEPENDENTES
    add_h2(doc, "12 CONTRATAÇÕES CORRELATAS/INTERDEPENDENTES")
    add_p(
        doc,
        f"No âmbito da mesma Emenda Parlamentar nº {EMENDA} (valor global de "
        f"{brl(RECURSO_GLOBAL_EMENDA)}), tramitam de forma autônoma e correlata outros dois "
        "processos do Programa PICSPel, sem qualquer sobreposição de itens:",
    )
    add_bullet(
        doc,
        "aquisição de 15 itens de consumo para auriculoterapia, aromaterapia, escalda-pés, "
        "práticas corporais e apoio educativo (incluindo plásticos para plastificação A4 de "
        "0,07 mm, bacias e sal grosso), classificados no elemento 3.3.90.30;",
        bold_prefix="Processo 1 — Materiais de Consumo: ",
    )
    add_bullet(
        doc,
        "contratação de pessoa jurídica especializada para curso de capacitação em "
        f"Auriculoterapia (80h, 30 a 50 vagas), com estimativa de {brl(RECURSO_COTA_SERVICOS)}, "
        "classificada no elemento 3.3.90.39.",
        bold_prefix="Processo 3 — Serviços de Capacitação: ",
    )
    add_p(
        doc,
        "Não há interdependência impeditiva para a licitação autônoma deste Processo 2 (Materiais "
        "Permanentes — R$ 11.600,00); contudo, há complementaridade funcional relevante — em "
        "especial entre a plastificadora (Item 05) e os plásticos de plastificação A4 (0,07 mm) do "
        "Processo 1, entre as maletas rígidas (Item 03) e os kits de consumo das equipes eMulti, e "
        "entre as jarras elétricas/garrafas térmicas (Itens 02 e 04) e as práticas de escalda-pés. "
        "Recomenda-se, portanto, o alinhamento de cronogramas entre os três processos e a "
        "conciliação consolidada do saldo financeiro da Emenda nº 43530013 pela área financeira.",
    )

    # 13 POSSÍVEIS IMPACTOS AMBIENTAIS E CRITÉRIOS DE SUSTENTABILIDADE
    add_h2(doc, "13 POSSÍVEIS IMPACTOS AMBIENTAIS E CRITÉRIOS DE SUSTENTABILIDADE")
    add_p(
        doc,
        "Os impactos ambientais associados à aquisição e ao uso dos materiais permanentes "
        "relacionam-se principalmente: (i) à geração de resíduos de embalagens (papelão, filmes "
        "plásticos e espumas/isopor de proteção); (ii) ao consumo de energia elétrica pelos "
        "equipamentos eletroportáteis (caixa de som, jarra elétrica e plastificadora); e (iii) ao "
        "descarte de equipamentos eletroeletrônicos e da bateria interna recarregável da caixa de "
        "som ao final de sua vida útil.",
    )
    add_p(
        doc,
        "Como medidas de mitigação e sustentabilidade (art. 11, inciso IV, e art. 18, § 1º, inciso "
        "XII, da Lei nº 14.133/2021), deverão ser observados: (a) destinação das embalagens "
        "recicláveis à coleta seletiva municipal conforme a Política Nacional de Resíduos Sólidos "
        "(Lei nº 12.305/2010); (b) prioridade para equipamentos duráveis (aço inox sem ampola de "
        "vidro, ABS rígido e eletroportáteis com desligamento automático e conformidade INMETRO), "
        "reduzindo substituições prematuras e desperdício de energia; e (c) destinação final "
        "ambientalmente adequada de Resíduos de Equipamentos Eletroeletrônicos (REEE) e baterias "
        "por meio de logística reversa ou pontos de coleta licenciados, nos termos do Decreto "
        "Federal nº 10.240/2020 e da Resolução CONAMA nº 401/2008.",
    )

    # 14 ANÁLISE DE RISCOS
    add_h2(doc, "14 ANÁLISE DE RISCOS")
    riscos = [
        (
            "Risco 1: Atraso na entrega dos equipamentos",
            "Média",
            "Médio",
            "Adiamento da estruturação física das equipes eMulti e da coordenação do PICSPel, "
            "com possível impacto no cronograma de execução da emenda parlamentar.",
            "Fixação de prazo objetivo de entrega (até 30 dias corridos da Nota de Empenho / "
            "Ordem de Fornecimento); acompanhamento ativo pelo fiscal; previsão de penalidades "
            "contratuais.",
            "Notificação formal da contratada, instauração de procedimento sancionatório conforme "
            "Lei nº 14.133/2021 e Decreto Municipal nº 6.799/2023 e, se necessário, convocação "
            "de licitante remanescente (art. 90, § 2º).",
        ),
        (
            "Risco 2: Entrega em desacordo técnico ou incompatibilidade elétrica (127 V vs. 220 V)",
            "Média",
            "Alto",
            "Inutilização imediata do equipamento (risco de queima na rede 220 V de Pelotas/RS), "
            "falta de segurança elétrica ou ausência de homologação ANATEL/INMETRO.",
            "Especificações técnicas claras e objetivas no TR com exigência explícita de tensão "
            "220 V / bivolt, plugue NBR 14136, ficha técnica, certificação INMETRO e homologação "
            "ANATEL; recebimento provisório e teste funcional prévio ao recebimento definitivo.",
            "Recusa formal do item no recebimento, exigência de retirada e substituição sem ônus "
            "no prazo de 5 (cinco) dias úteis e aplicação das sanções cabíveis.",
        ),
        (
            "Risco 3: Defeito ou vício oculto durante a garantia / falha na assistência técnica",
            "Baixa/Média",
            "Médio",
            "Indisponibilidade temporária do equipamento nas equipes eMulti ou na coordenação, "
            "prejudicando as atividades coletivas e educativas.",
            "Exigência de garantia mínima de 12 (doze) meses a partir do recebimento definitivo, "
            "com assistência técnica autorizada no Brasil, sem ônus logístico para a SMS e prazo "
            "máximo de reparo/substituição definido no TR.",
            "Acionamento formal da garantia junto à contratada/fabricante, exigência de "
            "substituição por equipamento novo em caso de não reparo no prazo legal (art. 18 do "
            "CDC) e responsabilização administrativa em caso de recusa.",
        ),
        (
            "Risco 4: Itens desertos ou fracassados / variação de preços",
            "Média",
            "Médio",
            "Lacuna na padronização dos kits das 08 equipes eMulti ou necessidade de repetição "
            "do procedimento licitatório.",
            "Parcelamento e julgamento por item; pesquisa de preços atualizada e realista "
            "(art. 23); especificação sem direcionamento de marca; ampla publicidade no PNCP.",
            "Verificação de atas vigentes para adesão vantajosa, republicação do item fracassado/"
            "deserto ou contratação direta nos termos do art. 75, inciso III, da Lei nº 14.133/2021.",
        ),
        (
            "Risco 5: Incompatibilidade de GND (Capital vs. Custeio) ou falha no controle patrimonial",
            "Média",
            "Alto",
            "Impedimento de empenho na fonte da emenda, ou extravio/dano aos bens permanentes "
            "sem rastreabilidade patrimonial.",
            "Validação prévia da classificação 4.4.90.52 (GND 4) pela Contabilidade/CCGOF antes "
            "da publicação do edital; tombamento patrimonial obrigatório antes da entrega às "
            "equipes e assinatura de Termo de Responsabilidade em reunião pós-entrega.",
            "Readequação orçamentária/fonte indicada pela CCGOF antes do certame; em caso de "
            "extravio ou dano pós-entrega, apuração de responsabilidade funcional e reposição.",
        ),
    ]
    for titulo_r, prob, imp, dano, prev, cont in riscos:
        add_h3(doc, titulo_r)
        add_p(doc, f"a) Probabilidade: {prob}.   b) Impacto: {imp}.", space_after=2)
        add_p(doc, dano, bold_prefix="c) Dano: ", space_after=2)
        add_p(doc, prev, bold_prefix="d) Ação preventiva: ", space_after=2)
        add_p(doc, cont, bold_prefix="e) Ação de contingência: ", space_after=5)

    # 15 DECLARAÇÃO DE VIABILIDADE DA CONTRATAÇÃO
    add_h2(doc, "15 DECLARAÇÃO DE VIABILIDADE DA CONTRATAÇÃO")
    add_p(
        doc,
        "A solução proposta mostra-se tecnicamente adequada, operacionalmente viável e "
        f"economicamente compatível com os recursos da Emenda Parlamentar nº {EMENDA} (estimativa "
        f"de {brl(VALOR_TOTAL)} para os 5 itens permanentes / 49 unidades). Há ampla oferta de "
        "fornecedores no mercado, os quantitativos estão objetivamente fundamentados no "
        "planejamento do PICSPel (08 equipes eMulti e coordenação do programa) e os riscos "
        "identificados são gerenciáveis. Declara-se a viabilidade técnica, operacional, econômica "
        "e administrativa da contratação, condicionada à confirmação orçamentária e de GND pela "
        "Contabilidade/CCGOF, recomendando-se o prosseguimento com a aprovação do Termo de "
        "Referência e a instauração do Pregão Eletrônico.",
    )

    # 16 EQUIPE DE PLANEJAMENTO
    add_assinaturas(
        doc,
        titulo_secao="16 EQUIPE DE PLANEJAMENTO / FISCALIZAÇÃO",
        incluir_equipe_texto=(
            "Equipe de planejamento – Programa PICSPel, Diretoria de Atenção Primária e área "
            "administrativa da SMS, com apoio do Almoxarifado e do setor de Patrimônio para "
            "recebimento, tombamento e distribuição."
        ),
    )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))


# ---------------------------------------------------------------------------
# DOCUMENTO 3: TERMO DE REFERÊNCIA (PADRÃO FIEL RP05_TR.docx_3bae.pdf)
# ---------------------------------------------------------------------------


def build_tr(brasao_bytes: bytes | None, out_path: Path) -> None:
    doc = new_doc(brasao_bytes, "Termo de Referência — PICSPel (Materiais Permanentes) | Emenda nº 43530013")
    add_title(
        doc,
        "TERMO DE REFERÊNCIA – PICSPel (MATERIAIS PERMANENTES)",
        "Processo Administrativo / Memorando nº [______]/2026  •  DFD nº [___]/2026  •  Emenda Parlamentar nº 43530013",
    )

    # 1 DO OBJETO, VALOR E CONDIÇÕES GERAIS DA CONTRATAÇÃO
    add_h2(doc, "1 DO OBJETO, VALOR E CONDIÇÕES GERAIS DA CONTRATAÇÃO")
    add_p(
        doc,
        "1.1 O presente Termo de Referência tem por objeto a aquisição pontual de equipamentos e "
        "materiais permanentes destinados à estruturação e ao fortalecimento do Programa Municipal "
        "de Práticas Integrativas e Complementares em Saúde (PICSPel), compreendendo 5 (cinco) "
        "itens e 49 (quarenta e nove) unidades físicas, para atendimento às 08 (oito) equipes "
        "eMulti e à coordenação do programa na rede de Atenção Primária à Saúde da Secretaria "
        "Municipal de Saúde de Pelotas/RS, custeada com recursos da Emenda Parlamentar nº "
        "43530013. Os quantitativos, especificações técnicas, valores estimados e condições "
        "encontram-se estabelecidos neste instrumento.",
    )

    rows_tr = [
        [
            f"{it.numero:02d}",
            f"{it.nome_curto.upper()} – {it.especificacao}",
            it.catmat,
            it.unidade,
            str(it.quantidade),
            brl(it.valor_unitario),
            brl(it.valor_total),
        ]
        for it in ITENS
    ]
    rows_tr.append(
        [
            "",
            "VALOR TOTAL ESTIMADO (5 itens / 49 unidades)",
            "",
            "",
            str(TOTAL_UNIDADES),
            "",
            brl(VALOR_TOTAL),
        ]
    )
    add_table(
        doc,
        ["Item", "Descrição", "CATMAT", "Unid.", "Qtd.", "Valor unit. (R$)", "Valor total (R$)"],
        rows_tr,
        col_widths_cm=[0.9, 8.3, 1.9, 1.0, 0.9, 2.0, 2.0],
        alignments=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ],
        font_size=7.8,
        last_row_total=True,
    )

    add_p(
        doc,
        "1.2 Os objetos deste instrumento não se enquadram como bens de luxo, conforme prevê o "
        "art. 20 da Lei Federal nº 14.133/2021 e a regulamentação vigente.",
    )
    add_p(
        doc,
        "1.3 Os objetos são caracterizados como bens comuns (art. 6º, inciso XIII, da Lei Federal "
        "nº 14.133/2021), pois apresentam padrões de desempenho e qualidade objetivamente definidos "
        "por meio de especificações usuais de mercado.",
    )
    add_p(
        doc,
        "1.4 Trata-se de aquisição pontual com entrega imediata e integral por item (sem adoção "
        "do Sistema de Registro de Preços). Nos termos do art. 95, incisos I e II, da Lei Federal "
        "nº 14.133/2021, o termo de contrato poderá ser substituído pela Nota de Empenho / Ordem "
        "de Fornecimento, com vigência adstrita ao cumprimento das obrigações de entrega, "
        "recebimento definitivo e pagamento (estimada em até 120 dias a contar do empenho), sem "
        "prejuízo da subsistência integral da obrigação de garantia mínima de 12 (doze) meses e "
        "assistência técnica prevista neste Termo.",
    )
    add_p(
        doc,
        "1.5 A entrega dos bens será única/integral por item adjudicado, mediante emissão de Nota "
        "de Empenho e Ordem de Fornecimento pela Secretaria Municipal de Saúde, observados os "
        "quantitativos certos definidos na tabela do subitem 1.1 e a disponibilidade orçamentária "
        "da Emenda Parlamentar nº 43530013.",
    )
    add_p(
        doc,
        "1.6 O parcelamento do objeto por itens é tecnicamente viável e economicamente vantajoso "
        "(art. 40, inciso V, alínea 'b', da Lei nº 14.133/2021), sendo o certame destinado "
        "exclusivamente à participação de Microempresas (ME) e Empresas de Pequeno Porte (EPP) "
        "em todos os itens, nos termos do art. 48, inciso I, da Lei Complementar nº 123/2006 e "
        "do art. 4º da Lei nº 14.133/2021, ressalvadas as hipóteses do art. 49 da referida Lei "
        "Complementar.",
    )
    add_p(
        doc,
        f"1.7 O valor total estimado é de {brl(VALOR_TOTAL)} (onze mil e seiscentos reais), "
        "conforme tabela acima e mapa comparativo de preços a ser anexado aos autos após a "
        "pesquisa formal nos termos do art. 23 da Lei nº 14.133/2021.",
    )
    add_p(
        doc,
        "1.8 Em razão de eventuais limitações na precisão das descrições do catálogo CATMAT, "
        "poderão ocorrer divergências em relação às especificações deste Termo de Referência. "
        "Nesses casos, prevalecerão integralmente as especificações deste Termo.",
    )

    # 2 DA FUNDAMENTAÇÃO DA CONTRATAÇÃO
    add_h2(doc, "2 DA FUNDAMENTAÇÃO DA CONTRATAÇÃO")
    add_p(
        doc,
        "2.1 A presente contratação visa atender às necessidades da Secretaria Municipal de Saúde "
        "de Pelotas/RS quanto à estruturação do Programa Municipal de Práticas Integrativas e "
        "Complementares em Saúde (PICSPel) na Atenção Primária à Saúde, em consonância com a "
        "Política Nacional de Práticas Integrativas e Complementares (PNPIC), provendo equipamentos "
        "e bens permanentes de apoio às 08 (oito) equipes eMulti (Centro/Porto, Fragata 1, "
        "Fragata 2, Três Vendas 1, Três Vendas 2, Areal, Colônia e Praias) e à coordenação do "
        "programa. A fundamentação técnica e o dimensionamento dos quantitativos constam "
        "detalhados no Estudo Técnico Preliminar.",
    )
    add_p(
        doc,
        "2.2 Justificativa para a aquisição pontual (sem Sistema de Registro de Preços) e para a "
        "modalidade adotada: a contratação pontual justifica-se porque os quantitativos são certos "
        "e previamente definidos para a implantação/estruturação das equipes eMulti e da "
        "coordenação, com custeio vinculado à Emenda Parlamentar nº 43530013 (objeto e prazo "
        "determinados), não se tratando de demanda contínua e indeterminada. Adota-se o Pregão "
        "Eletrônico com julgamento por menor preço por item visando ampliar a competitividade, "
        "prevenir eventual fracionamento indevido de despesa no exercício financeiro (art. 75, "
        "§ 1º, da Lei nº 14.133/2021) e assegurar isonomia e publicidade no PNCP.",
    )
    add_p(
        doc,
        "2.3 A contratação está alinhada ao planejamento da Secretaria Municipal de Saúde e da "
        "Diretoria de Atenção Primária [INFORMAR previsão no Plano de Contratações Anual – PCA "
        "e número do item, ou registrar justificativa de inclusão excepcional vinculada à Emenda "
        "Parlamentar nº 43530013].",
    )

    # 3 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO CONSIDERADO O CICLO DE VIDA DO OBJETO
    add_h2(doc, "3 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO CONSIDERADO O CICLO DE VIDA DO OBJETO")
    add_p(
        doc,
        "A solução consiste no fornecimento integral por item dos 5 (cinco) bens permanentes "
        "relacionados no subitem 1.1, com entrega no Almoxarifado da Secretaria Municipal de Saúde "
        "(ou local indicado pela Diretoria de Atenção Primária / coordenação PICSPel), conferência "
        "documental e teste funcional básico (especialmente verificação de tensão 220 V e "
        "funcionamento dos eletroportáteis), recebimento provisório e definitivo, incorporação "
        "patrimonial e tombamento (emissão de plaqueta/registro patrimonial e Termo de "
        "Responsabilidade), distribuição às 08 equipes eMulti e à coordenação do programa, "
        "utilização assistencial/educativa com cobertura de garantia mínima de 12 (doze) meses e "
        "assistência técnica autorizada, e descarte ambientalmente adequado ao fim da vida útil "
        "conforme a Política Nacional de Resíduos Sólidos (Lei nº 12.305/2010), o Decreto Federal "
        "nº 10.240/2020 (logística reversa de eletroeletrônicos) e a Resolução CONAMA nº 401/2008 "
        "(baterias). A fundamentação encontra-se pormenorizada no Estudo Técnico Preliminar.",
    )

    # 4 GESTÃO/ANÁLISE DE RISCOS
    add_h2(doc, "4 GESTÃO/ANÁLISE DE RISCOS")
    add_p(
        doc,
        "Os principais riscos identificados na contratação são: (i) atraso na entrega dos "
        "equipamentos, comprometendo o cronograma de estruturação das equipes eMulti e execução "
        "da emenda (probabilidade média; impacto médio); (ii) entrega em desacordo técnico, "
        "incompatibilidade de tensão elétrica (127 V em vez de 220 V) ou ausência de certificação "
        "INMETRO/homologação ANATEL (probabilidade média; impacto alto); (iii) defeito ou vício "
        "oculto no período de garantia ou dificuldade de acionamento de assistência técnica "
        "(probabilidade baixa/média; impacto médio); (iv) itens desertos ou fracassados "
        "(probabilidade média; impacto médio); e (v) falha na compatibilização orçamentária de "
        "GND 4 ou extravio por ausência de controle patrimonial (probabilidade média; impacto "
        "alto).",
    )
    add_p(
        doc,
        "Como medidas de prevenção e mitigação, serão observados o detalhamento objetivo das "
        "especificações técnicas (sem direcionamento de marca), a exigência de ficha técnica e "
        "certificações obrigatórias (INMETRO/ANATEL), a conferência física e o teste funcional em "
        "220 V antes do recebimento definitivo, a garantia mínima de 12 (doze) meses com "
        "assistência técnica sem ônus logístico para a SMS, a validação prévia da dotação/GND pela "
        "CCGOF, o tombamento patrimonial antes da entrega às equipes e a fiscalização ativa da "
        "contratação. Em caso de descumprimento, serão aplicadas as sanções previstas na Lei "
        "Federal nº 14.133/2021 e no Decreto Municipal nº 6.799/2023.",
    )

    # 5 REQUISITOS DA CONTRATAÇÃO
    add_h2(doc, "5 REQUISITOS DA CONTRATAÇÃO")
    add_p(
        doc,
        "5.1 Não serão exigidas marcas ou modelos específicos, sendo admitidos produtos de "
        "diferentes fabricantes, desde que atendam integralmente às especificações técnicas deste "
        "Termo, preservando-se a isonomia e a competitividade (art. 41, inciso I, da Lei nº "
        "14.133/2021).",
    )
    add_p(doc, "5.2 Requisitos técnicos específicos:", space_after=3)
    add_bullet(
        doc,
        "Caixa de som portátil (Item 01): áudio estéreo (dois alto-falantes ou sistema "
        "equivalente), potência mínima de 20 W RMS, conectividade Bluetooth (v5.0 ou superior), "
        "bateria recarregável (autonomia mínima de 6 horas), carregamento compatível com 220 V / "
        "bivolt (ou USB/USB-C com cabo incluso) e certificado de homologação ANATEL;",
    )
    add_bullet(
        doc,
        "Garrafa térmica (Item 02): 100% aço inoxidável de grau alimentício (corpo e ampola "
        "interna sem vidro), isolamento a vácuo, capacidade aproximada de 1,8 L (1,7 L a 2,0 L), "
        "acionamento por bombeamento de pressão com vedação hermética e desempenho térmico mínimo "
        "de 6 horas para líquidos quentes;",
    )
    add_bullet(
        doc,
        "Maleta organizadora rígida (Item 03): estrutura 100% ABS rígido (ou polímero de "
        "resistência equivalente ou superior), fechamento por zíper duplo na horizontal, alça "
        "reforçada, compartimento interno forrado/organizador e dimensões compatíveis com a faixa "
        "especificada no subitem 1.1;",
    )
    add_bullet(
        doc,
        "Jarra elétrica (Item 04): corpo em aço inoxidável, capacidade entre 1,7 L e 2,2 L, "
        "potência mínima de 1.500 W (referência 1.800 W), tensão 220 V (60 Hz), plugue padrão "
        "brasileiro (ABNT NBR 14136), base destacável 360°, proteção contra funcionamento a seco, "
        "desligamento automático e certificação compulsória de segurança elétrica no INMETRO "
        "(Portaria INMETRO nº 148/2022 ou substitutiva);",
    )
    add_bullet(
        doc,
        "Plastificadora/laminadora (Item 05): formatos A3/A4/A5/A6 (largura útil mínima de "
        "320 mm), laminação a quente (e opcionalmente a frio), compatível com polaseal de 0,05 mm "
        "a 0,07 mm (125 a 175 micras) ou superior, sistema anti-atolamento, tensão 220 V (60 Hz) "
        "ou bivolt e plugue padrão brasileiro (ABNT NBR 14136).",
    )
    add_p(
        doc,
        "5.3 Os produtos deverão ser novos, de primeiro uso, de linha de produção corrente "
        "(vedados itens recondicionados, remanufaturados ou usados), entregues em embalagem "
        "original lacrada e íntegra, com identificação em língua portuguesa do fabricante/modelo, "
        "número de série (quando aplicável), acessórios completos e manual de instruções em "
        "português.",
    )
    add_p(
        doc,
        "5.4 Garantia mínima e assistência técnica: todos os bens deverão contar com garantia "
        "mínima total de 12 (doze) meses — compreendendo a garantia legal de 90 (noventa) dias "
        "(art. 26, inciso II, da Lei nº 8.078/1990 – CDC) acrescida de garantia contratual "
        "complementar de 9 (nove) meses — contados a partir da data do recebimento definitivo. "
        "A contratada deverá assegurar atendimento em rede de assistência técnica autorizada no "
        "Brasil, arcando integralmente com eventuais custos de coleta, frete de ida e volta, "
        "peças e mão de obra durante a vigência da garantia. O prazo máximo para reparo ou "
        "substituição de bem em garantia será de 15 (quinze) dias corridos contados da "
        "notificação formal pela SMS; persistindo o vício ou excedido o prazo legal, deverá "
        "ocorrer substituição por equipamento novo, idêntico ou superior, sem ônus para a "
        "Administração (art. 18 do CDC).",
    )
    add_p(
        doc,
        "5.5 Conformidade técnica e regulatória: a licitante provisoriamente classificada em "
        "primeiro lugar deverá apresentar, com a proposta ou no prazo de diligência fixado pelo "
        "pregoeiro, catálogo, prospecto ou ficha técnica oficial do fabricante que comprove o "
        "pleno atendimento às especificações do item (incluindo tensão 220 V / bivolt), bem como: "
        "(a) certificado/número de homologação na ANATEL para o Item 01 (caixa de som com "
        "Bluetooth); e (b) comprovação de certificação compulsória no INMETRO para o Item 04 "
        "(jarra elétrica) e demais itens sujeitos à certificação compulsória. Não se exige "
        "registro ou AFE da ANVISA por não se tratarem de insumos ou equipamentos médicos.",
    )
    add_p(
        doc,
        "5.6 Amostras: caso a ficha técnica ou catálogo apresentado não seja suficiente para "
        "dirimir dúvida técnica objetiva quanto à conformidade do produto, a Administração poderá "
        "exigir, motivadamente, do licitante provisoriamente classificado em primeiro lugar, "
        "amostra física do item no prazo de 5 (cinco) dias úteis, nos termos do art. 41, inciso "
        "II, da Lei nº 14.133/2021. A reprovação fundamentada resultará na desclassificação da "
        "proposta para o respectivo item.",
    )
    add_p(
        doc,
        "5.7 Para fins de habilitação, a empresa vencedora deverá apresentar, na forma do edital:",
        space_after=3,
    )
    for hab in [
        "Comprovante de inscrição e de situação cadastral ativa no CNPJ, com CNAE compatível com o fornecimento dos itens arrematados;",
        "Contrato social, estatuto ou documento constitutivo equivalente vigente, devidamente registrado, com as alterações consolidadas, quando houver;",
        "Certidões negativas ou positivas com efeito de negativas de regularidade fiscal e trabalhista: Fazenda Federal (inclusive seguridade social), Fazenda Estadual, Fazenda Municipal, FGTS e CNDT;",
        "Declaração de que não emprega menor em condições vedadas, nos termos do inciso XXXIII do art. 7º da Constituição Federal;",
        "Declaração de cumprimento das exigências de reserva de cargos para pessoa com deficiência e para reabilitado da Previdência Social, quando aplicável (art. 63, inciso IV, da Lei nº 14.133/2021);",
        "No ato da entrega dos bens: termo/certificado de garantia mínima de 12 (doze) meses, relação ou canal de assistência técnica autorizada e manual de instruções em língua portuguesa.",
    ]:
        add_bullet(doc, hab)
    add_p(
        doc,
        "5.8 É vedada a subcontratação total ou parcial do fornecimento do objeto, ressalvada a "
        "execução de reparos em garantia por meio da rede de assistência técnica autorizada do "
        "fabricante.",
    )
    add_p(
        doc,
        "5.9 O fornecimento deverá contemplar todos os custos diretos e indiretos, incluindo "
        "transporte (frete CIF), seguro, embalagens protetoras, carga, descarga, tributos e "
        "encargos, de inteira responsabilidade da contratada.",
    )
    add_p(
        doc,
        "5.10 Os bens somente serão aceitos definitivamente após conferência técnica e teste "
        "funcional efetuados pela Secretaria Municipal de Saúde e, caso não satisfaçam às "
        "especificações, possuam tensão incompatível ou apresentem avarias/defeitos, deverão ser "
        "retirados e substituídos pelo fornecedor, às suas expensas, no prazo máximo de 5 (cinco) "
        "dias úteis contados da notificação.",
    )

    # 6 DA EXECUÇÃO DO OBJETO
    add_h2(doc, "6 DA EXECUÇÃO DO OBJETO")
    add_p(
        doc,
        f"6.1 Os bens deverão ser entregues no {ENDERECO_ALMOXARIFADO}, ou em outro endereço "
        "situado no Município de Pelotas/RS formalmente indicado pela Diretoria de Atenção "
        "Primária / coordenação do PICSPel na Ordem de Fornecimento.",
    )
    add_p(
        doc,
        "6.2 O prazo de entrega integral dos itens empenhados será de até 30 (trinta) dias "
        "corridos, contados do recebimento da Nota de Empenho / Ordem de Fornecimento pela "
        "contratada.",
    )
    add_p(
        doc,
        "6.3 Caso não seja possível a entrega na data assinalada, a contratada deverá comunicar "
        "as razões, devidamente fundamentadas e comprovadas, com pelo menos 5 (cinco) dias "
        "corridos de antecedência do término do prazo, para análise do pedido de prorrogação pela "
        "Administração, ressalvadas as hipóteses de caso fortuito e força maior.",
    )
    add_p(
        doc,
        "6.4 Os equipamentos entregues deverão estar rigorosamente acompanhados de todos os "
        "acessórios originais (cabos de alimentação/carregamento, bases destacáveis, tampas, "
        "divisórias), manuais em língua portuguesa, termos de garantia e Nota Fiscal/DANFE "
        "detalhada (com indicação de marca, modelo, número de série quando aplicável e referência "
        "à Emenda Parlamentar nº 43530013).",
    )
    add_p(
        doc,
        "6.5 Instalação e verificação operacional: os equipamentos adquiridos são de operação "
        "imediata ('plug and play'), não demandando obras civis ou instalação predial complexa. "
        "Previamente ao recebimento definitivo, a equipe técnica/fiscalização realizará teste "
        "funcional básico em rede 220 V (itens 01, 04 e 05) e inspeção física/vedação (itens 02 "
        "e 03). Concluído o recebimento definitivo, a SMS promoverá o tombamento/incorporação "
        "patrimonial antes da distribuição às 08 equipes eMulti e à coordenação do PICSPel.",
    )

    # 7 DA GESTÃO E FISCALIZAÇÃO DA CONTRATAÇÃO
    add_h2(doc, "7 DA GESTÃO E FISCALIZAÇÃO DA CONTRATAÇÃO")
    add_p(
        doc,
        "7.1 A contratação (formalizada mediante Nota de Empenho / Ordem de Fornecimento vinculada "
        "ao Edital e a este Termo de Referência, nos termos do art. 95 da Lei nº 14.133/2021) "
        "deverá ser executada fielmente pelas partes, de acordo com as cláusulas avençadas e as "
        "normas da Lei Federal nº 14.133/2021, respondendo cada parte pelas consequências de sua "
        "inexecução total ou parcial.",
    )
    add_p(
        doc,
        "7.2 A execução da contratação será acompanhada e fiscalizada por servidores formalmente "
        "designados pela Secretaria Municipal de Saúde (art. 117 da Lei nº 14.133/2021), sendo "
        f"indicados: Gestor — {GESTOR[0]} (Matrícula: {GESTOR[1]}); e Fiscal — {FISCAL[0]} "
        f"(Matrícula: {FISCAL[1]}), competindo-lhes acompanhar prazos, conferir quantidades e "
        "especificações técnicas, realizar ou acompanhar os testes funcionais, registrar "
        "ocorrências, atestar notas fiscais, articular o tombamento patrimonial e acionar a "
        "garantia quando necessário.",
    )
    add_p(
        doc,
        "7.3 As comunicações entre a Administração e a contratada serão realizadas por escrito "
        "sempre que o ato exigir formalidade, admitindo-se o uso de mensagem eletrônica "
        "institucional para esse fim.",
    )
    add_p(
        doc,
        "7.4 O descumprimento total ou parcial das obrigações assumidas caracterizará "
        "inadimplência, sujeitando a contratada às sanções administrativas previstas nos arts. "
        "155 a 163 da Lei Federal nº 14.133/2021 e no Decreto Municipal nº 6.799/2023 "
        "(advertência, multa moratória e compensatória, impedimento de licitar e contratar e "
        "declaração de inidoneidade), assegurados o contraditório e a ampla defesa, sem prejuízo "
        "da responsabilidade civil e criminal cabível.",
    )

    # 8 RECEBIMENTO DO OBJETO E INCORPORAÇÃO PATRIMONIAL
    add_h2(doc, "8 RECEBIMENTO DO OBJETO E INCORPORAÇÃO PATRIMONIAL")
    add_p(
        doc,
        "8.1 Os bens serão recebidos provisoriamente, de forma sumária, no prazo de até 7 (sete) "
        "dias corridos a contar da entrega, pelo responsável pelo acompanhamento e fiscalização "
        "(art. 140, inciso II, alínea 'a', da Lei nº 14.133/2021), mediante conferência da Nota "
        "Fiscal, quantitativos, integridade das embalagens e documentação de garantia, para "
        "posterior verificação da conformidade técnica e funcional com este Termo e com a proposta.",
    )
    add_p(
        doc,
        "8.2 Os bens poderão ser rejeitados, no todo ou em parte, quando em desacordo com as "
        "especificações técnicas deste Termo ou da proposta, quando apresentarem avarias, defeitos "
        "de funcionamento, incompatibilidade de tensão elétrica (127 V em vez de 220 V) ou "
        "ausência de certificações obrigatórias, devendo ser retirados e substituídos no prazo "
        "máximo de 5 (cinco) dias úteis a contar da notificação, às custas exclusivas da "
        "contratada, sem prejuízo da aplicação das penalidades cabíveis.",
    )
    add_p(
        doc,
        "8.3 Os bens serão recebidos definitivamente no prazo de até 8 (oito) dias corridos, "
        "contados do recebimento provisório (art. 140, inciso II, alínea 'b', da Lei nº "
        "14.133/2021), após verificação detalhada da qualidade, quantidade, documentação "
        "regulatória (INMETRO/ANATEL), teste funcional básico em 220 V e aceitação formal "
        "mediante termo detalhado assinado pela fiscalização.",
    )
    add_p(
        doc,
        "8.4 Incorporação patrimonial e tombamento: concluído o recebimento definitivo e atestada "
        "a Nota Fiscal, a documentação será encaminhada ao setor competente de Patrimônio da "
        "SMS/Prefeitura Municipal de Pelotas para registro contábil-patrimonial, emplaquetamento/"
        "tombamento dos bens permanentes e emissão dos respectivos Termos de Responsabilidade "
        "para distribuição às 08 (oito) equipes eMulti e à coordenação do PICSPel, observados os "
        "critérios patrimoniais validados pela Contabilidade/CCGOF.",
    )

    # 9 DOS CRITÉRIOS DE MEDIÇÃO E PAGAMENTO
    add_h2(doc, "9 DOS CRITÉRIOS DE MEDIÇÃO E PAGAMENTO")
    add_p(
        doc,
        "9.1 A liquidação e o pagamento serão efetuados conforme a efetiva entrega integral de "
        "cada item empenhado, após o recebimento definitivo, o ateste pela fiscalização e a "
        "conferência da Nota Fiscal/DANFE pelo órgão competente da Administração, mediante crédito "
        "em conta bancária específica de titularidade da contratada, dentro do cronograma fixado "
        "pela Secretaria Municipal da Fazenda de Pelotas e observada a ordem cronológica prevista "
        "no art. 141 da Lei Federal nº 14.133/2021.",
    )
    add_p(
        doc,
        "9.2 A Nota Fiscal/DANFE deverá discriminar detalhadamente os bens entregues (descrição, "
        "marca, modelo, número de série quando aplicável, quantidade, valor unitário e valor "
        "total), fazer menção expressa ao número da Nota de Empenho e à Emenda Parlamentar nº "
        "43530013, e estar acompanhada da comprovação da manutenção das condições de regularidade "
        "fiscal e trabalhista da contratada. Serão efetuadas na fonte as retenções tributárias "
        "previstas na legislação vigente (inclusive IRRF nos termos da IN RFB nº 1.234/2012 e "
        "normas municipais aplicáveis).",
    )

    # 10 DA FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR
    add_h2(doc, "10 DA FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR")
    add_p(
        doc,
        "10.1 O fornecedor será selecionado por meio de licitação na modalidade Pregão, em sua "
        "forma eletrônica, para aquisição pontual (sem Sistema de Registro de Preços), nos termos "
        "do art. 28, inciso I, da Lei Federal nº 14.133/2021, por possuir o objeto especificações "
        "usuais de mercado e padrões de desempenho e qualidade objetivamente definidos — "
        "ressalvada eventual análise motivada pela autoridade competente quanto à Dispensa "
        "Eletrônica com disputa (art. 75, inciso II), condicionada à certificação prévia de "
        "ausência de fracionamento de despesa no ramo de atividade no exercício financeiro.",
    )
    add_p(
        doc,
        "10.2 Critério de julgamento: menor preço por item, com adjudicação por item.",
    )
    add_p(
        doc,
        "10.3 Tratamento diferenciado: participação exclusiva para Microempresas (ME) e Empresas "
        "de Pequeno Porte (EPP) em todos os itens (valor individual inferior a R$ 80.000,00), "
        "conforme art. 48, inciso I, da Lei Complementar nº 123/2006 e art. 4º da Lei nº "
        "14.133/2021, ressalvado o disposto no art. 49 da LC nº 123/2006.",
    )
    add_p(
        doc,
        "10.4 Estimativa de preços: consolidada nos termos do art. 23 da Lei nº 14.133/2021, "
        "conforme mapa comparativo de preços anexo [ANEXAR mapa de preços oficial com fontes, "
        "datas de cotação e critério estatístico — média/mediana — antes da publicação do edital].",
    )

    # 11 DAS OBRIGAÇÕES DO CONTRATANTE
    add_h2(doc, "11 DAS OBRIGAÇÕES DO CONTRATANTE")
    for ob in [
        "Exigir o integral cumprimento, pela contratada, de todas as obrigações assumidas;",
        "Receber o objeto no prazo e nas condições estabelecidas neste Termo de Referência;",
        "Prestar à contratada as informações e esclarecimentos necessários à perfeita execução do objeto;",
        "Notificar a contratada, por escrito, sobre imperfeições, defeitos, avarias ou irregularidades verificadas nos bens, fixando prazo para substituição ou reparo em garantia;",
        "Exercer o acompanhamento, a inspeção técnica/funcional e a fiscalização por servidores designados, anotando em registro próprio todas as ocorrências detectadas;",
        "Promover o tombamento e a incorporação patrimonial dos bens permanentes após o recebimento definitivo;",
        "Efetuar os pagamentos no prazo, forma e condições estabelecidos;",
        "Aplicar à contratada as sanções administrativas previstas na Lei nº 14.133/2021, no Decreto Municipal nº 6.799/2023 e no edital;",
        "Não praticar atos de ingerência na administração da contratada;",
        "Emitir decisão fundamentada sobre solicitações e reclamações relacionadas à execução da contratação;",
        "Fiscalizar a manutenção, pela contratada, das condições de habilitação exigidas na licitação.",
    ]:
        add_bullet(doc, ob)

    # 12 DAS OBRIGAÇÕES DO CONTRATADO
    add_h2(doc, "12 DAS OBRIGAÇÕES DO CONTRATADO")
    for ob in [
        "Cumprir integralmente todas as obrigações assumidas neste Termo de Referência, no Edital e em sua proposta, assumindo os riscos e as despesas decorrentes da boa e perfeita execução do objeto;",
        "Entregar bens novos, de primeiro uso, em embalagem original lacrada, acompanhados de manuais em língua portuguesa, cabos/acessórios originais, plugues padrão brasileiro (ABNT NBR 14136) e compatibilidade elétrica 220 V / bivolt;",
        "Assegurar a conformidade técnica e regulatória dos itens fornecidos, incluindo certificação compulsória INMETRO (quando aplicável) e homologação ANATEL (Item 01);",
        "Prestar garantia mínima de 12 (doze) meses contados do recebimento definitivo e assegurar assistência técnica autorizada no Brasil, sem ônus de transporte, coleta ou reparo para a Administração;",
        "Atender às determinações regulares emitidas pelo fiscal, gestor ou autoridade superior;",
        "Reparar, corrigir, retirar ou substituir, às suas expensas e nos prazos fixados neste Termo (5 dias úteis no recebimento; até 15 dias corridos em garantia), os bens com vícios, defeitos ou desconformidades;",
        "Responsabilizar-se pelos vícios e danos decorrentes do fornecimento, nos termos do Código de Defesa do Consumidor (Lei nº 8.078/1990);",
        "Não permitir a utilização de trabalho de menor de dezesseis anos, exceto na condição de aprendiz a partir dos quatorze anos, nem de menor de dezoito anos em atividades noturnas, perigosas ou insalubres;",
        "Cumprir, quando aplicável, a reserva de cargos prevista em lei para pessoa com deficiência, para reabilitado da Previdência Social ou para aprendiz (art. 92, inciso XVII, da Lei nº 14.133/2021);",
        "Responsabilizar-se pelo pagamento de todos os tributos, encargos trabalhistas, previdenciários, fiscais e comerciais incidentes sobre o fornecimento;",
        "Manter, durante toda a execução da contratação, as condições de habilitação e qualificação exigidas na licitação;",
        "Arcar com todas as despesas de transporte (frete CIF), seguro, carga, descarga e embalagens necessárias à entrega no local indicado;",
        "Emitir Nota Fiscal/DANFE com descrição detalhada dos bens (incluindo marca, modelo e número de série, quando houver) e referência à Emenda Parlamentar nº 43530013, viabilizando a incorporação patrimonial.",
    ]:
        add_bullet(doc, ob)

    # 13 DA ADEQUAÇÃO ORÇAMENTÁRIA
    add_h2(doc, "13 DA ADEQUAÇÃO ORÇAMENTÁRIA")
    add_p(
        doc,
        "As despesas decorrentes da presente contratação correrão à conta de recursos vinculados "
        f"à Emenda Parlamentar nº {EMENDA} (valor global de {brl(RECURSO_GLOBAL_EMENDA)}), "
        "conforme indicação orçamentária abaixo, condicionada à confirmação prévia de dotação, "
        "fonte e compatibilidade do elemento de despesa de capital (GND 4 – 4.4.90.52) pela "
        "Contabilidade/CCGOF:",
    )
    rows_orc = [
        ["Emenda Parlamentar", f"nº {EMENDA}"],
        ["Projeto/Atividade", "[INFORMAR — Atenção Primária / PICSPel]"],
        ["Fonte de Recurso", "[INFORMAR — vinculada à Emenda Parlamentar nº 43530013]"],
        ["Dotação / Reduzido", "[INFORMAR]"],
        [
            "Classificação da Despesa",
            "4.4.90.52 (Equipamentos e Material Permanente — validar compatibilidade GND 4 com Contabilidade/CCGOF)",
        ],
        ["Valor Total Estimado", f"{brl(VALOR_TOTAL)} (onze mil e seiscentos reais)"],
    ]
    add_table(
        doc,
        ["Campo Orçamentário", "Detalhamento"],
        rows_orc,
        col_widths_cm=[5.2, 11.8],
        alignments=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
        font_size=8.5,
    )

    # 14 EQUIPE DE PLANEJAMENTO
    add_assinaturas(
        doc,
        titulo_secao="14 EQUIPE DE PLANEJAMENTO",
        incluir_equipe_texto=(
            "Equipe de planejamento – Departamento/Serviço demandante: Programa PICSPel, "
            "Diretoria de Atenção Primária, Diretoria Administrativa e Financeira e apoio do "
            "Almoxarifado e do setor de Patrimônio da SMS."
        ),
    )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))


# ---------------------------------------------------------------------------
# VALIDAÇÃO PROGRAMÁTICA DOS .DOCX GERADOS
# ---------------------------------------------------------------------------


def _extrair_texto_completo(doc: Document) -> str:
    partes: list[str] = []
    for sec in doc.sections:
        partes.extend(p.text for p in sec.header.paragraphs)
        partes.extend(p.text for p in sec.footer.paragraphs)
    partes.extend(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                partes.append(cell.text)
    return "\n".join(partes)


def validar_documento(path: Path, tipo: str) -> list[str]:
    erros: list[str] = []
    if not path.exists():
        return [f"Arquivo inexistente: {path}"]

    doc = Document(str(path))
    texto = _extrair_texto_completo(doc)

    # 1. Conferência de brasão/imagem embutida no cabeçalho (quando disponível)
    rels = doc.sections[0].header.part.rels.values()
    tem_imagem_cabecalho = any("image" in r.reltype for r in rels)
    if not tem_imagem_cabecalho:
        erros.append(f"[{tipo}] Cabeçalho sem imagem de brasão embutida.")

    # 2. Conferência institucional e equipe correta
    obrigatorios = [
        "ESTADO DO RIO GRANDE DO SUL",
        "MUNICÍPIO DE PELOTAS",
        "SECRETARIA MUNICIPAL DE SAÚDE",
        "43530013",
        "4.4.90.52",
        "R$ 11.600,00",
        "Dariane Lima Portela",
        "44151",
        "Aline Kohler Geppert",
        "29371",
        "Angela Moreira Vitória",
        "14.133/2021",
    ]
    for ob in obrigatorios:
        if ob not in texto:
            erros.append(f"[{tipo}] Texto obrigatório ausente: {ob!r}")

    # 3. Termos proibidos (resquícios de copia-e-cola de consumo ou RP05)
    proibidos = [
        "protetores descartáveis",
        "fluidos corporais",
        "Óleos e frascos",
        "CÂNULA DE GUEDEL",
        "TUBO ENDOTRAQUEAL",
        "Josué Corrêa da Silva",
        "27.5 cm",
        "automático.Capacidade",
        "20 RMS watts",
        "20 KHz",
        "Poliseladora",
    ]
    for pr in proibidos:
        if pr in texto:
            erros.append(f"[{tipo}] Resquício indevido encontrado: {pr!r}")

    # 4. Verificação aritmética rigorosa da tabela de itens
    tabela_itens = None
    for t in doc.tables:
        hdr = [c.text.strip() for c in t.rows[0].cells]
        if hdr and hdr[0] == "Item" and any("Valor unit." in h for h in hdr):
            tabela_itens = t
            break

    if tabela_itens is None:
        erros.append(f"[{tipo}] Tabela de itens com valores unitários/totais não encontrada.")
    else:
        hdr = [c.text.strip() for c in tabela_itens.rows[0].cells]
        idx_qtd = next(i for i, h in enumerate(hdr) if "Qtd" in h)
        idx_vu = next(i for i, h in enumerate(hdr) if "Valor unit." in h)
        idx_vt = next(i for i, h in enumerate(hdr) if "Valor total" in h)

        def parse_brl(s: str) -> Decimal:
            m = re.search(r"R\$\s*([\d\.]+,\d{2})", s)
            if not m:
                raise ValueError(f"Valor BRL inválido: {s!r}")
            return Decimal(m.group(1).replace(".", "").replace(",", "."))

        soma_vt = Decimal("0.00")
        soma_qtd = 0
        linhas_itens = tabela_itens.rows[1:-1]
        if len(linhas_itens) != len(ITENS):
            erros.append(f"[{tipo}] Esperados {len(ITENS)} itens na tabela, encontrados {len(linhas_itens)}.")
        for idx_row, (row, esperado) in enumerate(zip(linhas_itens, ITENS), start=1):
            num_txt = row.cells[0].text.strip()
            qtd_val = int(row.cells[idx_qtd].text.strip())
            vu_val = parse_brl(row.cells[idx_vu].text.strip())
            vt_val = parse_brl(row.cells[idx_vt].text.strip())
            if num_txt != f"{esperado.numero:02d}":
                erros.append(f"[{tipo}] Linha {idx_row}: item {num_txt!r} != {esperado.numero:02d}")
            if qtd_val != esperado.quantidade:
                erros.append(f"[{tipo}] Item {esperado.numero:02d}: qtd {qtd_val} != {esperado.quantidade}")
            if vu_val != esperado.valor_unitario:
                erros.append(f"[{tipo}] Item {esperado.numero:02d}: VU {vu_val} != {esperado.valor_unitario}")
            if (vu_val * qtd_val) != vt_val or vt_val != esperado.valor_total:
                erros.append(f"[{tipo}] Item {esperado.numero:02d}: VT {vt_val} != {esperado.valor_total}")
            soma_vt += vt_val
            soma_qtd += qtd_val

        linha_total = tabela_itens.rows[-1]
        total_declarado = parse_brl(linha_total.cells[idx_vt].text.strip())
        qtd_total_declarada = int(linha_total.cells[idx_qtd].text.strip())
        if soma_vt != VALOR_TOTAL or total_declarado != VALOR_TOTAL:
            erros.append(
                f"[{tipo}] Soma/total divergente: soma={soma_vt}, declarado={total_declarado}, esperado={VALOR_TOTAL}"
            )
        if soma_qtd != TOTAL_UNIDADES or qtd_total_declarada != TOTAL_UNIDADES:
            erros.append(
                f"[{tipo}] Quantidade total divergente: soma={soma_qtd}, declarado={qtd_total_declarada}"
            )

    # 5. Verificações estruturais específicas por documento
    if tipo == "ETP":
        secoes_etp = [
            "1 IDENTIFICAÇÃO DO PROCESSO",
            "2 DESCRIÇÃO DA NECESSIDADE",
            "3 DA CONTRATAÇÃO E DO PLANEJAMENTO",
            "4 REQUISITOS DA CONTRATAÇÃO",
            "5 ESTIMATIVA DAS QUANTIDADES",
            "6 LEVANTAMENTO DE MERCADO",
            "7 ESTIMATIVA DO VALOR DA CONTRATAÇÃO",
            "8 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO",
            "9 JUSTIFICATIVA PARA PARCELAMENTO",
            "10 DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS",
            "11 PROVIDÊNCIAS PRÉVIAS AO CONTRATO",
            "12 CONTRATAÇÕES CORRELATAS/INTERDEPENDENTES",
            "13 POSSÍVEIS IMPACTOS AMBIENTAIS",
            "14 ANÁLISE DE RISCOS",
            "15 DECLARAÇÃO DE VIABILIDADE DA CONTRATAÇÃO",
            "16 EQUIPE DE PLANEJAMENTO",
        ]
        for s in secoes_etp:
            if s not in texto:
                erros.append(f"[ETP] Seção ausente: {s}")
        # Checar que não há parágrafo duplicado acidental (excluindo classificações de probabilidade/impacto)
        paras = [
            p.text.strip()
            for p in doc.paragraphs
            if len(p.text.strip()) > 40 and not p.text.strip().startswith("a) Probabilidade:")
        ]
        vistos = set()
        for pt in paras:
            if pt in vistos:
                erros.append(f"[ETP] Parágrafo duplicado detectado: {pt[:70]}...")
            vistos.add(pt)

    elif tipo == "TR":
        secoes_tr = [
            "1 DO OBJETO, VALOR E CONDIÇÕES GERAIS DA CONTRATAÇÃO",
            "2 DA FUNDAMENTAÇÃO DA CONTRATAÇÃO",
            "3 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO CONSIDERADO O CICLO DE VIDA DO OBJETO",
            "4 GESTÃO/ANÁLISE DE RISCOS",
            "5 REQUISITOS DA CONTRATAÇÃO",
            "6 DA EXECUÇÃO DO OBJETO",
            "7 DA GESTÃO E FISCALIZAÇÃO DA CONTRATAÇÃO",
            "8 RECEBIMENTO DO OBJETO E INCORPORAÇÃO PATRIMONIAL",
            "9 DOS CRITÉRIOS DE MEDIÇÃO E PAGAMENTO",
            "10 DA FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR",
            "11 DAS OBRIGAÇÕES DO CONTRATANTE",
            "12 DAS OBRIGAÇÕES DO CONTRATADO",
            "13 DA ADEQUAÇÃO ORÇAMENTÁRIA",
            "14 EQUIPE DE PLANEJAMENTO",
        ]
        for s in secoes_tr:
            if s not in texto:
                erros.append(f"[TR] Seção ausente (padrão RP05_TR): {s}")

    elif tipo == "DFD":
        secoes_dfd = [
            "1 OBJETO",
            "1.1 TIPO DE OBJETO",
            "2 FORMA DE CONTRATAÇÃO SUGERIDA",
            "3 JUSTIFICATIVA DA CONTRATAÇÃO",
            "4 DESCRIÇÃO DO OBJETO",
            "5 DOTAÇÃO ORÇAMENTÁRIA",
            "6 LOCAL DE ENTREGA",
            "7 PREVISÃO DE ENTREGA",
            "EQUIPE DE PLANEJAMENTO / FISCALIZAÇÃO",
        ]
        for s in secoes_dfd:
            if s not in texto:
                erros.append(f"[DFD] Seção ausente: {s}")

    # 6. Conferir presença das 8 equipes eMulti em todos os 3 documentos
    for eq in EQUIPES_EMULTI:
        if eq not in texto:
            erros.append(f"[{tipo}] Equipe eMulti ausente: {eq!r}")

    return erros


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Gera e valida os documentos DOCX do Processo 2 — PICSPel (Materiais Permanentes)."
    )
    parser.add_argument(
        "--out-dir",
        type=Path,
        default=OUT_DIR,
        help="Diretório de saída dos arquivos .docx",
    )
    parser.add_argument(
        "--brasao",
        type=Path,
        default=None,
        help="Caminho opcional para o arquivo PNG do brasão de Pelotas",
    )
    args = parser.parse_args(list(argv) if argv is not None else None)

    brasao_bytes = carregar_brasao_bytes(args.brasao)
    if brasao_bytes:
        print(f"Brasão carregado com sucesso ({len(brasao_bytes)} bytes).")
    else:
        print("AVISO: brasão não encontrado; gerando cabeçalho institucional apenas textual.")

    arquivos = [
        ("DFD", args.out_dir / "PICSPel_Materiais_Permanentes_DFD_Revisado.docx", build_dfd),
        ("ETP", args.out_dir / "PICSPel_Materiais_Permanentes_ETP_Revisado.docx", build_etp),
        ("TR", args.out_dir / "PICSPel_Materiais_Permanentes_TR.docx", build_tr),
    ]

    erros_totais: list[str] = []
    for tipo, caminho, builder in arquivos:
        builder(brasao_bytes, caminho)
        erros = validar_documento(caminho, tipo)
        if erros:
            erros_totais.extend(erros)
            print(f"[FALHA] {tipo}: {caminho}")
            for e in erros:
                print(f"   - {e}")
        else:
            doc = Document(str(caminho))
            print(
                f"[OK] {tipo}: {caminho} "
                f"({caminho.stat().st_size:,} bytes | {len(doc.paragraphs)} parágrafos | "
                f"{len(doc.tables)} tabelas | Total={brl(VALOR_TOTAL)})"
            )

    if erros_totais:
        print(f"\nValidação encerrada com {len(erros_totais)} erro(s).", file=sys.stderr)
        return 1

    print("\nTodos os 3 documentos .docx foram gerados e validados com 100% de conformidade.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
