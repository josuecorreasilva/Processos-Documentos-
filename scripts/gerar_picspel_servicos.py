#!/usr/bin/env python3
"""Gera e valida DFD revisado, ETP revisado e TR para o Processo 1:
PICSPel Serviços (Capacitação em Auriculoterapia — Emenda Parlamentar nº 43530013).

Padrão documental oficial SMS Pelotas / RP05 (A4, brasão oficial de Pelotas,
Arial, cabeçalho institucional, rodapé 'Página X de Y', tabelas estruturadas).
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = Path(__file__).resolve().parents[1]
OUT_DIR = ROOT / "docs" / "08-PICSPEL-SERVICOS"
BRASAO_CANDIDATES = [
    ROOT / "docs" / "assets" / "brasao_pelotas.png",
    ROOT / "docs" / "assets" / "brasao_rp_header.png",
]

# ---------------------------------------------------------------------------
# Parâmetros consolidados do Processo 1 — PICSPel Serviços (Emenda 43530013)
# ---------------------------------------------------------------------------
EMENDA = "43530013"
RECURSO_GLOBAL = 100_000.00
COTA_SERVICOS = 40_000.00
COTA_MATERIAIS = 60_000.00
ESTIMADO_CONSUMO = 47_200.00
ESTIMADO_PERMANENTES = 11_200.00
ESTIMADO_MATERIAIS_TOTAL = ESTIMADO_CONSUMO + ESTIMADO_PERMANENTES  # 58.400,00
ESTIMADO_GLOBAL_3_PROC = COTA_SERVICOS + ESTIMADO_MATERIAIS_TOTAL  # 98.400,00

QTD_TURMAS = 1
VALOR_UNITARIO = 40_000.00
VALOR_TOTAL = QTD_TURMAS * VALOR_UNITARIO  # 40.000,00

CARGA_HORARIA_TOTAL = 80
CARGA_TEORICA = 48
CARGA_PRATICA = 32
VAGAS_MIN = 30
VAGAS_MAX = 50
CUSTO_POR_HORA = VALOR_TOTAL / CARGA_HORARIA_TOTAL  # 500,00
CUSTO_POR_VAGA_MAX = VALOR_TOTAL / VAGAS_MAX  # 800,00
CUSTO_POR_VAGA_MIN = VALOR_TOTAL / VAGAS_MIN  # 1.333,33

GESTOR = ("Dariane Lima Portela", "44151", "Gestor")
FISCAL = ("Aline Kohler Geppert", "29371", "Fiscal")
SECRETARIA = ("Angela Moreira Vitória", "Secretária Municipal de Saúde — Autorizo")

UNIDADE_DEMANDANTE = (
    "Secretaria Municipal de Saúde – SMS / Diretoria de Atenção Primária / "
    "Programa Municipal de Práticas Integrativas e Complementares em Saúde (PICSPel)"
)


def brl(v: float) -> str:
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _find_brasao() -> Path | None:
    for candidate in BRASAO_CANDIDATES:
        if candidate.exists():
            return candidate
    return None


# ---------------------------------------------------------------------------
# Utilitários de formatação DOCX (padrão SMS Pelotas / RP05)
# ---------------------------------------------------------------------------
def _font(run, *, name: str = "Arial", size: float = 11, bold: bool = False, italic: bool = False):
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


def _shade(cell, fill: str = "EDEDED"):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def _cell_margins(cell, top: int = 70, bottom: int = 70, left: int = 100, right: int = 100):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for side, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def _set_table_borders(table, color: str = "595959", sz: str = "4"):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), color)
        borders.append(el)
    tblPr.append(borders)


def _set_table_no_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none")
        borders.append(el)
    tblPr.append(borders)


def _set_row_repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:tblHeader"))


def _set_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(OxmlElement("w:cantSplit"))


def _set_cell_width(cell, width_cm: float):
    cell.width = Cm(width_cm)
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(int(round(width_cm * 566.929))))
    tcW.set(qn("w:type"), "dxa")


def _page_fields(paragraph):
    def _fld(instr: str):
        r1 = paragraph.add_run()
        _font(r1, size=9)
        fc1 = OxmlElement("w:fldChar")
        fc1.set(qn("w:fldCharType"), "begin")
        r1._r.append(fc1)

        r2 = paragraph.add_run()
        _font(r2, size=9)
        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = instr
        r2._r.append(it)

        r3 = paragraph.add_run()
        _font(r3, size=9)
        fc2 = OxmlElement("w:fldChar")
        fc2.set(qn("w:fldCharType"), "end")
        r3._r.append(fc2)

    r_pre = paragraph.add_run("Página ")
    _font(r_pre, size=9)
    _fld("PAGE")
    r_mid = paragraph.add_run(" de ")
    _font(r_mid, size=9)
    _fld("NUMPAGES")


def new_doc() -> Document:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    sec.top_margin = Cm(2.4)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)
    sec.header_distance = Cm(0.8)
    sec.footer_distance = Cm(0.9)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    rPr = style._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rFonts.set(qn(attr), "Arial")
    style.font.size = Pt(11)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.space_after = Pt(5)

    # Cabeçalho oficial SMS Pelotas (brasão centralizado + 3 linhas institucionais)
    header = sec.header
    header.is_linked_to_previous = False
    hp = header.paragraphs[0]
    hp.clear()
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.space_before = Pt(0)
    hp.paragraph_format.space_after = Pt(2)
    brasao = _find_brasao()
    if brasao is not None:
        hp.add_run().add_picture(str(brasao), width=Cm(1.65))

    for idx, line_text in enumerate(
        [
            "ESTADO DO RIO GRANDE DO SUL",
            "MUNICÍPIO DE PELOTAS",
            "SECRETARIA MUNICIPAL DE SAÚDE",
        ]
    ):
        p_h = header.add_paragraph()
        p_h.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_h.paragraph_format.space_before = Pt(0)
        p_h.paragraph_format.space_after = Pt(0 if idx < 2 else 4)
        r_h = p_h.add_run(line_text)
        _font(r_h, size=10, bold=False)

    # Rodapé com numeração de páginas
    footer = sec.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.clear()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(0)
    fp.paragraph_format.space_after = Pt(0)
    _page_fields(fp)
    return doc


def doc_title(doc: Document, text: str):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(10)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    _font(r, size=12, bold=True)
    return p


def h2(doc: Document, text: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    _font(r, size=11, bold=True)
    return p


def p(
    doc: Document,
    text: str,
    *,
    bold_prefix: str | None = None,
    bold: bool = False,
    justify: bool = True,
    space_after: float = 5,
    keep_with_next: bool = False,
):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.space_after = Pt(space_after)
    para.paragraph_format.keep_with_next = keep_with_next
    if bold_prefix is not None:
        r1 = para.add_run(bold_prefix)
        _font(r1, size=11, bold=True)
        r2 = para.add_run(text)
        _font(r2, size=11, bold=False)
    else:
        r = para.add_run(text)
        _font(r, size=11, bold=bold)
    return para


def bullet(doc: Document, text: str, *, bold_prefix: str | None = None):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    para.paragraph_format.left_indent = Cm(0.6)
    para.paragraph_format.first_line_indent = Cm(-0.35)
    para.paragraph_format.space_after = Pt(3.5)
    rb = para.add_run("• ")
    _font(rb, size=11, bold=False)
    if bold_prefix is not None:
        r1 = para.add_run(bold_prefix)
        _font(r1, size=11, bold=True)
        r2 = para.add_run(text)
        _font(r2, size=11, bold=False)
    else:
        r = para.add_run(text)
        _font(r, size=11, bold=False)
    return para


def checkbox(doc: Document, text: str, checked: bool = False):
    mark = "X" if checked else " "
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.space_after = Pt(2)
    r = para.add_run(f"( {mark} ) {text}")
    _font(r, size=11, bold=checked)
    return para


def styled_table(
    doc: Document,
    headers: list[str],
    rows: list[list[str]],
    col_widths_cm: list[float],
    *,
    col_aligns: list[WD_ALIGN_PARAGRAPH] | None = None,
    font_size: float = 9.0,
    bold_last_row: bool = False,
    shade_last_row: bool = False,
):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    _set_table_borders(t)

    # Cabeçalho
    hdr_row = t.rows[0]
    _set_row_repeat_header(hdr_row)
    _set_row_cant_split(hdr_row)
    for ci, h_text in enumerate(headers):
        cell = hdr_row.cells[ci]
        _set_cell_width(cell, col_widths_cm[ci])
        _shade(cell, "E6E6E6")
        _cell_margins(cell, top=75, bottom=75, left=90, right=90)
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        cell.text = ""
        para = cell.paragraphs[0]
        para.alignment = (
            col_aligns[ci] if col_aligns is not None else WD_ALIGN_PARAGRAPH.LEFT
        )
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        r = para.add_run(h_text)
        _font(r, size=font_size, bold=True)

    # Linhas de dados
    for ri, row_data in enumerate(rows, start=1):
        row = t.rows[ri]
        _set_row_cant_split(row)
        is_last = ri == len(rows)
        for ci, val in enumerate(row_data):
            cell = row.cells[ci]
            _set_cell_width(cell, col_widths_cm[ci])
            if is_last and shade_last_row:
                _shade(cell, "F5F5F5")
            _cell_margins(cell, top=65, bottom=65, left=90, right=90)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            cell.text = ""
            lines = str(val).split("\n")
            for li, line in enumerate(lines):
                para = cell.paragraphs[0] if li == 0 else cell.add_paragraph()
                para.alignment = (
                    col_aligns[ci]
                    if col_aligns is not None
                    else WD_ALIGN_PARAGRAPH.LEFT
                )
                para.paragraph_format.space_before = Pt(0)
                para.paragraph_format.space_after = Pt(1.5 if li < len(lines) - 1 else 0)
                r = para.add_run(line)
                _font(r, size=font_size, bold=(bold_last_row and is_last))

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(4)
    return t


def assinaturas_padrao(doc: Document, *, incluir_cabecalho_secao: bool = True, titulo_secao: str = "EQUIPE DE PLANEJAMENTO / FISCALIZAÇÃO", texto_equipe: str | None = None):
    if incluir_cabecalho_secao:
        h2(doc, titulo_secao)
    if texto_equipe:
        p(doc, texto_equipe)
    p(
        doc,
        "Certificamos que somos responsáveis pela elaboração do presente instrumento e que o mesmo está em consonância com a legislação vigente.",
        keep_with_next=True,
    )

    # Tabela invisível 1x2 para Gestor e Fiscal lado a lado
    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    _set_table_no_borders(t)
    _set_row_cant_split(t.rows[0])

    for ci, pessoa in enumerate((GESTOR, FISCAL)):
        cell = t.rows[0].cells[ci]
        _set_cell_width(cell, 8.5)
        _cell_margins(cell, top=160, bottom=60, left=80, right=80)
        cell.text = ""
        lines = [
            "__________________________",
            pessoa[0],
            f"Matrícula: {pessoa[1]}",
            pessoa[2],
        ]
        for li, line in enumerate(lines):
            para = cell.paragraphs[0] if li == 0 else cell.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(1)
            para.paragraph_format.keep_with_next = True
            r = para.add_run(line)
            _font(r, size=10, bold=(li == 1))

    p_data = doc.add_paragraph()
    p_data.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_data.paragraph_format.space_before = Pt(12)
    p_data.paragraph_format.space_after = Pt(18)
    p_data.paragraph_format.keep_with_next = True
    r_d = p_data.add_run("Pelotas/RS, _____ de ______________ de 2026.")
    _font(r_d, size=11)

    p_line = doc.add_paragraph()
    p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_line.paragraph_format.space_before = Pt(4)
    p_line.paragraph_format.space_after = Pt(1)
    p_line.paragraph_format.keep_with_next = True
    r_l = p_line.add_run("__________________________________________")
    _font(r_l, size=11)

    p_nome = doc.add_paragraph()
    p_nome.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_nome.paragraph_format.space_after = Pt(1)
    p_nome.paragraph_format.keep_with_next = True
    r_n = p_nome.add_run(SECRETARIA[0])
    _font(r_n, size=10, bold=True)

    p_cargo = doc.add_paragraph()
    p_cargo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_cargo.paragraph_format.space_after = Pt(0)
    r_c = p_cargo.add_run(SECRETARIA[1])
    _font(r_c, size=10)


# ---------------------------------------------------------------------------
# 1) DFD REVISADO — PICSPel Serviços (Capacitação em Auriculoterapia)
# ---------------------------------------------------------------------------
def gerar_dfd(path: Path):
    doc = new_doc()
    doc_title(
        doc,
        "DOCUMENTO DE FORMALIZAÇÃO DA DEMANDA nº [___]/2026 – PICSPel (SERVIÇOS DE CAPACITAÇÃO)",
    )

    p(doc, f"{UNIDADE_DEMANDANTE}.", bold_prefix="UNIDADE DEMANDANTE: ")
    p(
        doc,
        f"Processo Administrativo / Memorando nº [______]/2026 – Emenda Parlamentar nº {EMENDA}.",
        bold_prefix="REFERÊNCIA ADMINISTRATIVA: ",
    )

    h2(doc, "1 OBJETO")
    p(
        doc,
        "Contratação de pessoa jurídica especializada para prestação de serviço técnico-educacional não continuado "
        "de capacitação presencial em Auriculoterapia destinada a profissionais de saúde de nível superior da rede "
        f"municipal de saúde de Pelotas/RS, com capacidade de {VAGAS_MIN} (trinta) a {VAGAS_MAX} (cinquenta) vagas "
        f"e carga horária total de {CARGA_HORARIA_TOTAL} (oitenta) horas presenciais — distribuídas em "
        f"{CARGA_TEORICA}h de Formação Teórica e {CARGA_PRATICA}h de Formação Prática Supervisionada — "
        "incluindo coordenação pedagógica, corpo docente habilitado, fornecimento de material didático-pedagógico, "
        "avaliação de aprendizagem e emissão de certificados, visando ao fortalecimento das Práticas Integrativas "
        f"e Complementares em Saúde (PICSPel), custeada com recursos da Emenda Parlamentar nº {EMENDA}.",
    )

    h2(doc, "1.1 TIPO DE OBJETO")
    checkbox(doc, "Material de Consumo")
    checkbox(doc, "Material Permanente / Equipamento")
    checkbox(doc, "Serviço continuado com dedicação exclusiva de mão de obra")
    checkbox(doc, "Serviço continuado sem dedicação exclusiva de mão de obra")
    checkbox(doc, "Serviço não continuado", checked=True)
    p(
        doc,
        "a presente demanda refere-se estritamente à contratação do serviço de formação/capacitação presencial "
        "em Auriculoterapia para profissionais de saúde de nível superior da rede municipal, financiada pela cota "
        f"de {brl(COTA_SERVICOS)} destinada a Outros Serviços de Terceiros – Pessoa Jurídica (3.3.90.39) da Emenda "
        f"Parlamentar nº {EMENDA} (valor global de {brl(RECURSO_GLOBAL)}). A aquisição de materiais de consumo "
        f"(processo autônomo estimado em {brl(ESTIMADO_CONSUMO)} – 3.3.90.30) e de materiais permanentes/equipamentos "
        f"(processo autônomo estimado em {brl(ESTIMADO_PERMANENTES)} – 4.4.90.52), custeados pela cota de "
        f"{brl(COTA_MATERIAIS)} da mesma emenda, tramita em DFDs/processos próprios, sem sobreposição de escopo.",
        bold_prefix="Observação: ",
    )

    h2(doc, "2 FORMA DE CONTRATAÇÃO SUGERIDA")
    checkbox(doc, "Pregão Eletrônico", checked=True)
    checkbox(doc, "Pregão Eletrônico com Sistema/Ata de Registro de Preços")
    checkbox(doc, "Dispensa com Disputa Eletrônica")
    checkbox(doc, "Adesão à Ata do Consórcio Público do Extremo Sul")
    checkbox(doc, "Adesão à Ata de Registro de Preços de Outro(s) Órgão(s)")
    p(
        doc,
        "menor preço (item único / preço global da turma), sob regime de empreitada por preço global "
        "(art. 6º, XXIX, art. 28, I, e art. 33, I, da Lei Federal nº 14.133/2021).",
        bold_prefix="Critério de julgamento sugerido: ",
    )
    p(
        doc,
        "embora o valor estimado (R$ 40.000,00) seja inferior ao limite legal atualizado de dispensa em razão do valor "
        "(art. 75, II, da Lei nº 14.133/2021), sugere-se a realização de Pregão Eletrônico por garantir ampla "
        "competitividade, publicidade no PNCP e segurança jurídica quanto ao controle de não fracionamento de despesas "
        "de mesma natureza no exercício (art. 75, § 1º), ressalvada a possibilidade de avaliação administrativa quanto "
        "à Dispensa com Disputa Eletrônica caso comprovada a ausência de fracionamento no exercício de 2026.",
        bold_prefix="Nota de enquadramento legal: ",
    )

    h2(doc, "3 JUSTIFICATIVA DA CONTRATAÇÃO")
    p(
        doc,
        f"A formalização desta demanda visa capacitar entre {VAGAS_MIN} (trinta) e {VAGAS_MAX} (cinquenta) "
        "profissionais de saúde de nível superior da Atenção Primária à Saúde (UBS e equipes eMulti) na prática de "
        "Auriculoterapia, promovendo a expansão, a padronização técnica e a inclusão segura de Práticas Integrativas "
        "e Complementares em Saúde (PICS) na rede municipal, em conformidade com a Política Nacional de Práticas "
        "Integrativas e Complementares (PNPIC – Portaria GM/MS nº 971/2006 e atualizações) e as diretrizes do "
        "Programa PICSPel.",
    )
    p(
        doc,
        f"Custeada integralmente pela cota de {brl(COTA_SERVICOS)} de Outros Serviços de Terceiros – Pessoa Jurídica "
        f"da Emenda Parlamentar nº {EMENDA}, a formação teórico-prática presencial de {CARGA_HORARIA_TOTAL} horas "
        "assegura qualificação clínica, domínio dos protocolos de biossegurança e aplicação segura da Auriculoterapia "
        "no Sistema Único de Saúde (SUS), ampliando a resolutividade e a integralidade do cuidado na Atenção Primária.",
    )

    h2(doc, "4 DESCRIÇÃO DO OBJETO")
    desc_item = (
        "Contratação de Serviço de Formação/Capacitação Presencial em Auriculoterapia para Profissionais de Saúde "
        "de Nível Superior da Rede Municipal de Pelotas/RS:\n"
        f"• Carga horária total: {CARGA_HORARIA_TOTAL} (oitenta) horas presenciais.\n"
        f"• Etapa I – Formação Teórica ({CARGA_TEORICA} horas presenciais): PICS no SUS e PNPIC, reflexologia auricular, "
        "Medicina Tradicional Chinesa (MTC), bases biomédicas/neurofisiológicas e integração à Atenção Primária à Saúde.\n"
        f"• Etapa II – Formação Prática Supervisionada ({CARGA_PRATICA} horas presenciais): mapeamento e aplicação prática, "
        "treinamento de habilidades clínicas, biossegurança, gerenciamento de resíduos (PGRSS) e prática supervisionada "
        "(com subdivisão em até 2 grupos de até 25 participantes ou proporcionalidade mínima de 1 supervisor para cada 25 "
        "participantes quando o número de inscritos superar 25).\n"
        f"• Público-alvo / capacidade: {VAGAS_MIN} (trinta) a {VAGAS_MAX} (cinquenta) profissionais de saúde de nível "
        "superior (faixa coberta integralmente pelo valor global da turma, sem custo adicional).\n"
        "• Entregáveis: plano de curso detalhado, corpo docente habilitado, material didático-pedagógico completo para "
        "todos os participantes, controle diário de frequência, avaliação de aprendizagem, relatório final e emissão de "
        "certificados individuais aos concluintes, conforme Estudo Técnico Preliminar e Termo de Referência."
    )
    styled_table(
        doc,
        ["Item", "Especificação", "Unidade", "Qtd.", "Valor unit. estimado", "Valor total estimado"],
        [
            [
                "01",
                desc_item,
                "Serviço\n(Turma:\n30 a 50\nvagas)",
                str(QTD_TURMAS),
                brl(VALOR_UNITARIO),
                brl(VALOR_TOTAL),
            ],
            ["", "VALOR TOTAL ESTIMADO", "", "", "", brl(VALOR_TOTAL)],
        ],
        col_widths_cm=[1.0, 9.2, 1.8, 1.0, 2.0, 2.0],
        col_aligns=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ],
        font_size=8.5,
        bold_last_row=True,
        shade_last_row=True,
    )
    p(
        doc,
        f"{brl(VALOR_TOTAL)} (quarenta mil reais), correspondente ao teto fixado para a cota de Outros Serviços de "
        f"Terceiros – Pessoa Jurídica na Emenda Parlamentar nº {EMENDA}, a ser validado por pesquisa formal de preços "
        "nos termos do art. 23 da Lei nº 14.133/2021.",
        bold_prefix="Valor total estimado: ",
    )

    h2(doc, "5 DOTAÇÃO ORÇAMENTÁRIA")
    styled_table(
        doc,
        ["Campo", "Informação"],
        [
            ["Emenda Parlamentar", f"{EMENDA} (valor global: {brl(RECURSO_GLOBAL)})"],
            ["Projeto/Atividade", "[A INFORMAR — Atenção Primária / PICSPel]"],
            ["Fonte de recurso", f"[A INFORMAR — vinculada à Emenda Parlamentar nº {EMENDA}]"],
            ["Reduzido / Dotação orçamentária", "[A INFORMAR — validação Contabilidade / CCGOF]"],
            [
                "Classificação da despesa",
                "3.3.90.39 – Outros Serviços de Terceiros – Pessoa Jurídica "
                "(subelemento de seleção/treinamento, ex.: 3.3.90.39.48, a validar pela Contabilidade/CCGOF)",
            ],
            ["Valor estimado da contratação", brl(VALOR_TOTAL)],
            [
                "Teto da cota de serviços na emenda",
                f"{brl(COTA_SERVICOS)} (cota de Outros Serviços de Terceiros – PJ)",
            ],
            [
                "Previsão no PCA 2026",
                "[A INFORMAR — nº do item no PCA 2026 ou justificativa de inclusão excepcional vinculada à emenda]",
            ],
        ],
        col_widths_cm=[5.6, 11.4],
        col_aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
        font_size=9.0,
    )

    h2(doc, "6 LOCAL DE EXECUÇÃO DOS SERVIÇOS")
    p(
        doc,
        "Auditório ou espaço físico disponibilizado pela Secretaria Municipal de Saúde de Pelotas / Diretoria de "
        "Atenção Primária, no Município de Pelotas/RS, para a realização integral das aulas teóricas e práticas "
        "presenciais, cujo endereço exato será indicado na Ordem de Serviço.",
    )

    h2(doc, "7 PREVISÃO DE INÍCIO E PRAZO DE EXECUÇÃO")
    p(
        doc,
        "Início da execução das aulas em até 30 (trinta) dias corridos após a emissão da Ordem de Serviço / Nota de "
        "Empenho, precedido pela aprovação do plano de curso, cronograma de encontros e comprovação documental do corpo "
        "docente (em até 10 dias corridos da Ordem de Serviço). A execução integral das 80 horas presenciais ocorrerá "
        "em até 90 (noventa) dias corridos a partir do início das aulas, conforme cronograma aprovado entre a contratada "
        "e a coordenação do PICSPel, com entrega dos certificados e relatório final em até 15 (quinze) dias úteis após "
        "o encerramento das aulas.",
    )

    assinaturas_padrao(doc)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


# ---------------------------------------------------------------------------
# 2) ETP REVISADO — PICSPel Serviços (Capacitação em Auriculoterapia)
# ---------------------------------------------------------------------------
def gerar_etp(path: Path):
    doc = new_doc()
    doc_title(doc, "ESTUDO TÉCNICO PRELIMINAR – PICSPel (SERVIÇOS DE CAPACITAÇÃO)")

    h2(doc, "1 IDENTIFICAÇÃO DO PROCESSO")
    p(doc, f"{UNIDADE_DEMANDANTE}.", bold_prefix="1.1 Unidade responsável pela contratação: ")
    p(
        doc,
        "Contratação de pessoa jurídica especializada para prestação de serviço técnico-educacional não continuado "
        "de capacitação presencial em Auriculoterapia destinada a profissionais de saúde de nível superior da rede "
        f"municipal de Pelotas/RS, para quantitativo de {VAGAS_MIN} (trinta) a {VAGAS_MAX} (cinquenta) vagas, com "
        f"carga horária total de {CARGA_HORARIA_TOTAL} (oitenta) horas presenciais ({CARGA_TEORICA}h teóricas e "
        f"{CARGA_PRATICA}h práticas supervisionadas), visando ao fortalecimento das Práticas Integrativas e "
        f"Complementares em Saúde (PICSPel), custeada com recursos da Emenda Parlamentar nº {EMENDA}.",
        bold_prefix="1.2 Objeto da contratação: ",
    )
    p(
        doc,
        "A formação teórico-prática visa à qualificação técnica, padronização em biossegurança e aplicação segura "
        "da Auriculoterapia no Sistema Único de Saúde (SUS), promovendo a expansão e a institucionalização das PICS "
        "na Atenção Primária à Saúde (APS) em conformidade com a Política Nacional de Práticas Integrativas e "
        "Complementares (PNPIC) e as diretrizes do Programa PICSPel.",
    )
    p(
        doc,
        f"Processo Administrativo / Memorando nº [______]/2026 – DFD nº [___]/2026 – PICSPel (SERVIÇOS DE CAPACITAÇÃO) "
        f"– Emenda Parlamentar nº {EMENDA}.",
        bold_prefix="1.3 Número do processo: ",
    )
    p(
        doc,
        f"Emenda Parlamentar nº {EMENDA}, no valor global de {brl(RECURSO_GLOBAL)}, dividida em duas cotas: "
        f"{brl(COTA_SERVICOS)} destinados a Outros Serviços de Terceiros – Pessoa Jurídica (3.3.90.39) e "
        f"{brl(COTA_MATERIAIS)} destinados a materiais (desdobrados em processo autônomo de Materiais de Consumo, "
        f"estimado em {brl(ESTIMADO_CONSUMO)} – 3.3.90.30, e processo autônomo de Materiais Permanentes e Equipamentos, "
        f"estimado em {brl(ESTIMADO_PERMANENTES)} – 4.4.90.52). O presente ETP abrange exclusivamente a contratação do "
        f"serviço de formação/capacitação presencial em Auriculoterapia (estimativa/teto de {brl(VALOR_TOTAL)}), "
        "sem sobreposição de escopo com os processos de aquisição de materiais.",
        bold_prefix="1.4 Fonte de recurso: ",
    )

    h2(doc, "2 DESCRIÇÃO DA NECESSIDADE")
    p(
        doc,
        "A Política Nacional de Práticas Integrativas e Complementares (PNPIC – Portaria GM/MS nº 971/2006 e "
        "atualizações) e a organização local do Programa PICSPel demandam a formação continuada e qualificada dos "
        "profissionais de saúde da rede municipal. A Secretaria Municipal de Saúde necessita contratar pessoa jurídica "
        "especializada para ministrar curso presencial de capacitação em Auriculoterapia voltado a profissionais de "
        f"nível superior da Atenção Primária à Saúde (Unidades Básicas de Saúde e equipes eMulti), abrangendo entre "
        f"{VAGAS_MIN} (trinta) e {VAGAS_MAX} (cinquenta) vagas.",
    )
    p(
        doc,
        "A insuficiência de profissionais formalmente capacitados restringe a oferta segura, padronizada e "
        "resolutiva da Auriculoterapia nas UBS e nas 08 (oito) equipes eMulti da rede municipal, limitando o acesso "
        "dos usuários a práticas integrativas complementares ao cuidado clínico convencional. O treinamento integrará "
        "conhecimentos teóricos em PICS no SUS, reflexologia auricular, Medicina Tradicional Chinesa (MTC), bases "
        "biomédicas e atenção à saúde, articulados com treinamento prático supervisionado de habilidades clínicas e "
        "protocolos de biossegurança.",
    )
    p(
        doc,
        "As aulas teóricas e práticas presenciais serão realizadas em auditório ou espaço físico disponibilizado "
        "pela Secretaria Municipal de Saúde de Pelotas / Diretoria de Atenção Primária, com início previsto em até "
        "30 (trinta) dias corridos após a emissão da Ordem de Serviço / Nota de Empenho e conclusão integral das 80h "
        "em até 90 (noventa) dias corridos contados do início das aulas, conforme cronograma pedagógico pactuado.",
    )

    h2(doc, "3 DA CONTRATAÇÃO E DO PLANEJAMENTO")
    p(
        doc,
        "A contratação tem por finalidade a prestação de serviço técnico-educacional não continuado, por escopo "
        f"definido (art. 6º, XVII, da Lei Federal nº 14.133/2021), de formação presencial em Auriculoterapia com "
        f"recursos da Emenda Parlamentar nº {EMENDA}. Trata-se de serviço comum (art. 6º, XIII, da Lei nº 14.133/2021), "
        "cujos padrões de desempenho, conteúdo programático, carga horária, qualificação docente e qualidade podem ser "
        "objetivamente definidos pelo edital e pelo Termo de Referência.",
    )
    p(
        doc,
        "Pregão Eletrônico (art. 28, I, c/c art. 29 da Lei nº 14.133/2021), com critério de julgamento menor preço "
        "(item único / preço global da turma – art. 33, I) e regime de execução por empreitada por preço global "
        "(art. 6º, XXIX), observando-se os princípios da legalidade, isonomia, competitividade, publicidade e "
        "economicidade.",
        bold_prefix="Forma de contratação sugerida: ",
    )
    p(
        doc,
        "Justifica-se a contratação pontual por escopo definido — e não pelo Sistema de Registro de Preços (SRP) — "
        "pelo fato de o objeto possuir demanda certa, cronograma delimitado, público-alvo previamente identificado e "
        "vinculação específica ao teto financeiro da cota de serviços da Emenda Parlamentar nº 43530013, não se "
        "tratando de demanda contínua ou de quantitativo indeterminado.",
    )
    p(
        doc,
        "Embora o valor estimado (R$ 40.000,00) situe-se abaixo do limite atualizado para dispensa de licitação em "
        "razão do valor (art. 75, II, da Lei nº 14.133/2021), recomenda-se o Pregão Eletrônico por maximizar a "
        "competitividade e eliminar riscos de questionamento quanto ao somatório anual de despesas de mesma natureza "
        "na SMS (art. 75, § 1º), sem prejuízo de eventual enquadramento em Dispensa com Disputa Eletrônica pela "
        "autoridade competente caso o controle interno/compras ateste expressamente a ausência de fracionamento no "
        "exercício de 2026.",
    )
    p(
        doc,
        "[A INFORMAR: confirmar se a contratação consta do Plano de Contratações Anual (PCA) 2026 e indicar o "
        "respectivo número do item; caso não conste, anexar justificativa formal de inclusão excepcional vinculada à "
        f"Emenda Parlamentar nº {EMENDA}, nos termos do art. 12, VII, da Lei nº 14.133/2021].",
        bold_prefix="Previsão no Plano de Contratações Anual (PCA): ",
    )
    p(
        doc,
        "Serão observadas integralmente as etapas da fase preparatória e externa da Lei nº 14.133/2021, incluindo o "
        "DFD, o presente Estudo Técnico Preliminar, a pesquisa formal de preços (art. 23), o Termo de Referência, a "
        "análise jurídica, a designação formal de gestor e fiscal e a divulgação no Portal Nacional de Contratações "
        "Públicas (PNCP).",
    )

    h2(doc, "4 REQUISITOS DA CONTRATAÇÃO")
    p(
        doc,
        "O serviço deverá ser prestado por pessoa jurídica especializada, legalmente constituída, com objeto social "
        "compatível com atividades de ensino, treinamento ou capacitação na área da saúde/PICS, que comprove "
        "experiência prévia na ministração de cursos de Auriculoterapia ou PICS e disponha de corpo docente "
        "qualificado (nível superior na área da saúde, registro ativo no conselho profissional competente, quando "
        "aplicável, e formação específica ou experiência comprovada em Auriculoterapia/PICS e docência).",
    )
    p(doc, "Requisitos específicos técnico-pedagógicos do curso:", bold=True)
    bullet(
        doc,
        f"Carga horária total de {CARGA_HORARIA_TOTAL} (oitenta) horas presenciais, estruturada em: "
        f"Etapa I – Formação Teórica ({CARGA_TEORICA} horas presenciais: PICS no SUS e PNPIC, reflexologia auricular, "
        "Medicina Tradicional Chinesa – MTC, bases biomédicas/neurofisiológicas e atenção à saúde na APS); e "
        f"Etapa II – Formação Prática Supervisionada ({CARGA_PRATICA} horas presenciais: mapeamento e localização de "
        "pontos, técnicas de aplicação, treinamento de habilidades clínicas, biossegurança, PGRSS e prática supervisionada).",
    )
    bullet(
        doc,
        f"Público-alvo e organização pedagógica: capacidade para atender de {VAGAS_MIN} (trinta) a {VAGAS_MAX} "
        "(cinquenta) profissionais de saúde de nível superior da rede municipal de Pelotas, sem variação do valor "
        "global dentro dessa faixa. Nas atividades práticas supervisionadas (32h), quando o número de inscritos "
        "ultrapassar 25 participantes, deverá ser assegurada a divisão em 2 (dois) grupos de até 25 participantes "
        "(ou alocação proporcional mínima de 1 instrutor/supervisor prático para cada 25 participantes), em coerência "
        "com o planejamento do PICSPel e os padrões de biossegurança.",
    )
    bullet(
        doc,
        "Material didático-pedagógico e insumos didáticos: fornecimento, pela contratada, de apostila/guia didático "
        "completo e atualizado (em meio digital e/ou impresso, incluindo mapas auriculares de referência e protocolos "
        "de biossegurança) para todos os participantes, além dos materiais didáticos de demonstração do docente. "
        "Os insumos clínicos de consumo destinados às práticas da rede municipal são objeto de processo autônomo "
        "(Materiais de Consumo PICSPel), sem sobreposição.",
    )
    bullet(
        doc,
        "Frequência, avaliação e certificação: controle diário nominal de frequência, avaliação teórico-prática de "
        "aprendizagem e emissão de certificados individuais válidos aos concluintes que atingirem frequência mínima "
        "de 75% (60 horas) e aproveitamento satisfatório, entregues juntamente com o relatório final em até 15 (quinze) "
        "dias úteis após o encerramento das aulas.",
    )
    p(
        doc,
        "as aulas teóricas e práticas presenciais ocorrerão em auditório ou espaço físico disponibilizado pela "
        "Secretaria Municipal de Saúde de Pelotas / Diretoria de Atenção Primária, no Município de Pelotas/RS. "
        "Todos os custos com corpo docente, coordenação pedagógica, deslocamento, hospedagem, alimentação da equipe "
        "da contratada, material didático, certificação, tributos e encargos correrão por conta exclusiva da contratada.",
        bold_prefix="Instalações físicas e custos inclusos: ",
    )
    p(
        doc,
        "habilitação jurídica (art. 66); regularidade fiscal, social e trabalhista (art. 68 – CNPJ ativo, fazendas "
        "Federal/INSS, Estadual e Municipal, FGTS, CNDT e declaração de cumprimento do art. 7º, XXXIII, da CF); "
        "qualificação econômico-financeira proporcional ao valor e à baixa complexidade econômico-financeira do "
        "contrato (art. 69); e qualificação técnico-operacional e técnico-profissional (art. 67), mediante atestado(s) "
        "de capacidade técnica de execução de curso/capacitação presencial em Auriculoterapia ou PICS com carga horária "
        "mínima de 40 horas (50% do objeto, art. 67, § 2º) e declaração de disponibilidade de equipe docente "
        "habilitada (com comprovação documental dos docentes antes do início das aulas). Vedadas exigências "
        "restritivas ou desarrazoadas.",
        bold_prefix="Habilitação e qualificação técnica: ",
    )
    p(
        doc,
        "é vedada a subcontratação total do objeto ou da coordenação pedagógica/certificação, admitindo-se apenas a "
        "vinculação formal de docentes/instrutores pela pessoa jurídica contratada (vínculo trabalhista, societário "
        "ou prestação de serviços), mantida a responsabilidade integral e exclusiva da contratada (art. 122 da Lei "
        "nº 14.133/2021).",
        bold_prefix="Subcontratação: ",
    )

    h2(doc, "5 ESTIMATIVA DAS QUANTIDADES")
    p(
        doc,
        f"O quantitativo foi dimensionado para atender uma turma/serviço único com capacidade entre {VAGAS_MIN} "
        f"(trinta) e {VAGAS_MAX} (cinquenta) profissionais de saúde de nível superior da rede de Atenção Primária à "
        f"Saúde de Pelotas (UBS e 08 equipes eMulti), estruturado em carga horária total de {CARGA_HORARIA_TOTAL} horas "
        f"presenciais ({CARGA_TEORICA}h teóricas e {CARGA_PRATICA}h práticas supervisionadas), de modo a viabilizar "
        "a cobertura das equipes e a ampliação efetiva da oferta de Auriculoterapia no município.",
    )
    p(
        doc,
        f"O valor estimado / teto da contratação está limitado a {brl(VALOR_TOTAL)}, correspondente à cota integral "
        f"de Outros Serviços de Terceiros – Pessoa Jurídica (3.3.90.39) da Emenda Parlamentar nº {EMENDA}. Como "
        f"parâmetro referencial de razoabilidade, o teto corresponde a {brl(CUSTO_POR_HORA)} por hora-aula total e a "
        f"custo médio por participante entre {brl(CUSTO_POR_VAGA_MAX)} (para 50 vagas) e {brl(CUSTO_POR_VAGA_MIN)} "
        "(para 30 vagas).",
    )
    styled_table(
        doc,
        ["Item", "Especificação / Descrição do Serviço", "Unidade", "Qtd.", "Valor Unit. Estimado", "Valor Total Estimado"],
        [
            [
                "01",
                "Contratação de Serviço de Formação/Capacitação Presencial em Auriculoterapia para Profissionais "
                f"de Saúde de Nível Superior. Carga horária total de {CARGA_HORARIA_TOTAL}h ({CARGA_TEORICA}h teóricas "
                f"presenciais e {CARGA_PRATICA}h práticas presenciais supervisionadas). Capacidade: {VAGAS_MIN} a "
                f"{VAGAS_MAX} vagas (com subdivisão em até 2 grupos de até 25 participantes nas atividades práticas "
                "quando > 25 inscritos). Inclui coordenação pedagógica, corpo docente habilitado, material didático, "
                "avaliação e certificação. Conforme DFD nº [___]/2026 e Termo de Referência.",
                "Serviço\n(Turma:\n30 a 50\nvagas)",
                str(QTD_TURMAS),
                brl(VALOR_UNITARIO),
                brl(VALOR_TOTAL),
            ],
            ["", "VALOR TOTAL ESTIMADO", "", "", "", brl(VALOR_TOTAL)],
        ],
        col_widths_cm=[1.0, 9.2, 1.8, 1.0, 2.0, 2.0],
        col_aligns=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ],
        font_size=8.5,
        bold_last_row=True,
        shade_last_row=True,
    )

    p(doc, "Classificação por natureza da despesa:", bold=True)
    styled_table(
        doc,
        ["Item", "Natureza da Despesa", "Destinação", "Valor"],
        [
            [
                "01",
                "3.3.90.39 (Outros Serviços de Terceiros – Pessoa Jurídica)",
                f"Capacitação presencial em Auriculoterapia ({VAGAS_MIN} a {VAGAS_MAX} profissionais de nível superior da APS – 80h)",
                brl(VALOR_TOTAL),
            ],
            ["", "TOTAL ESTIMADO – SERVIÇOS DE TERCEIROS (PJ)", "", brl(VALOR_TOTAL)],
        ],
        col_widths_cm=[1.1, 5.4, 7.7, 2.8],
        col_aligns=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ],
        font_size=8.5,
        bold_last_row=True,
        shade_last_row=True,
    )
    p(
        doc,
        "a Contabilidade / CCGOF deverá validar a empenhabilidade e o subelemento específico sob o elemento de "
        f"despesa 3.3.90.39 (ex.: 3.3.90.39.48 – Serviços de Seleção e Treinamento) no âmbito da Emenda Parlamentar nº {EMENDA}.",
        bold_prefix="Observação: ",
    )

    h2(doc, "6 LEVANTAMENTO DE MERCADO")
    p(
        doc,
        "Foi realizado levantamento preliminar de mercado indicando a existência de pluralidade de instituições de "
        "ensino, escolas técnicas/pós-graduação em saúde, entidades especializadas e empresas de capacitação em PICS "
        "e Auriculoterapia no Estado do Rio Grande do Sul e em âmbito nacional aptas a executar integralmente o plano "
        "de curso presencial de 80 horas com corpo docente qualificado.",
    )
    p(doc, "Alternativas de contratação analisadas:", bold=True)
    bullet(
        doc,
        "Licitação por Pregão Eletrônico, critério menor preço (item único / preço global): solução recomendada, "
        "por assegurar ampla competitividade, isonomia, publicidade no PNCP, definição objetiva de qualidade no TR e "
        "segurança jurídica frente às regras de controle de fracionamento de despesas.",
        bold_prefix="(I) ",
    )
    bullet(
        doc,
        "Dispensa com Disputa Eletrônica em razão do valor (art. 75, II, da Lei nº 14.133/2021): alternativa "
        "juridicamente condicionada à comprovação prévia de que o somatório das despesas com objetos de mesma natureza "
        "na unidade gestora no exercício de 2026 não ultrapassará o limite legal atualizado (art. 75, § 1º); opta-se "
        "preferencialmente pelo Pregão Eletrônico para afastar qualquer risco de fracionamento e ampliar a disputa.",
        bold_prefix="(II) ",
    )
    bullet(
        doc,
        "Adesão a Ata de Registro de Preços (arts. 82 a 86 da Lei nº 14.133/2021): considerada de baixa aderência, "
        "haja vista a especificidade do curso presencial customizado de 80h em Auriculoterapia voltado à APS/PICSPel "
        "em Pelotas/RS, devendo eventual ata compatível ser verificada apenas em caráter preliminar antes do edital.",
        bold_prefix="(III) ",
    )
    bullet(
        doc,
        "Inexigibilidade de licitação para treinamento e aperfeiçoamento de pessoal (art. 74, III, 'f', da Lei nº "
        "14.133/2021): descartada, pois o mercado apresenta pluralidade de prestadores capacitados e padrões "
        "pedagógicos objetivamente mensuráveis, caracterizando plena viabilidade de competição.",
        bold_prefix="(IV) ",
    )
    bullet(
        doc,
        "Execução direta com servidores próprios da SMS: descartada, visto que o objetivo central da emenda é "
        "ampliar o quadro de profissionais habilitados sem desassistir as UBS e equipes eMulti em funcionamento.",
        bold_prefix="(V) ",
    )
    p(
        doc,
        "A pesquisa formal de preços será consolidada nos termos do art. 23, § 1º, da Lei nº 14.133/2021 (PNCP / "
        "Painel de Preços, contratações similares da Administração Pública, tabelas/mídia especializada e cotações "
        "diretas com instituições de ensino e empresas especializadas), adotando-se critério estatístico justificado "
        "(mediana ou média saneada), com mapa comparativo anexo ao processo antes da fase externa.",
    )

    h2(doc, "7 ESTIMATIVA DO VALOR DA CONTRATAÇÃO")
    p(
        doc,
        f"Valor total estimado preliminar: {brl(VALOR_TOTAL)} (quarenta mil reais), correspondente exatamente ao teto "
        f"fixado para a cota de Outros Serviços de Terceiros – Pessoa Jurídica (3.3.90.39) da Emenda Parlamentar nº {EMENDA}.",
    )
    p(
        doc,
        "O valor de referência definitivo da licitação será validado pela pesquisa formal de preços (art. 23 da Lei nº "
        "14.133/2021) antes da publicação do edital, respeitado o limite orçamentário da emenda parlamentar.",
    )

    h2(doc, "8 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO")
    p(
        doc,
        "A solução consiste na contratação de pessoa jurídica especializada para ministrar curso presencial de "
        f"capacitação em Auriculoterapia com carga horária de {CARGA_HORARIA_TOTAL} horas ({CARGA_TEORICA}h teóricas "
        f"e {CARGA_PRATICA}h práticas supervisionadas) para {VAGAS_MIN} a {VAGAS_MAX} profissionais de saúde de nível "
        "superior da rede municipal de Pelotas/RS, em auditório/espaço físico disponibilizado pela SMS.",
    )
    p(
        doc,
        "O ciclo completo de execução do objeto compreende: planejamento e formalização do processo licitatório; "
        "emissão da Nota de Empenho e da Ordem de Serviço; apresentação e aprovação do plano de curso detalhado, "
        "cronograma de encontros e comprovação do corpo docente; execução presencial da Etapa I (48h teóricas) e da "
        "Etapa II (32h práticas supervisionadas, com subdivisão em até 2 subgrupos de até 25 participantes ou "
        "supervisão proporcional de 1:25); fornecimento de material didático-pedagógico; controle nominal de frequência "
        "e avaliação de aprendizagem; recebimento provisório e definitivo por etapa; emissão e entrega dos certificados "
        "aos concluintes; entrega do relatório final consolidado; e descarte ambientalmente adequado dos resíduos "
        "eventualmente gerados nas aulas práticas.",
    )

    h2(doc, "9 JUSTIFICATIVA PARA NÃO PARCELAMENTO")
    p(
        doc,
        "Por tratar-se de contratação de serviço único e integrado de formação/turma em Auriculoterapia, o objeto é "
        "técnica e pedagogicamente indivisível, justificando-se o não parcelamento (item único / preço global), nos "
        "termos do art. 40, § 3º, e do art. 47, § 1º, da Lei Federal nº 14.133/2021. A fragmentação entre módulo "
        "teórico, prática supervisionada, produção de material didático e certificação comprometeria a coerência "
        "pedagógica, a uniformidade metodológica, a padronização dos protocolos de biossegurança e a responsabilidade "
        "técnico-acadêmica pela certificação final dos profissionais.",
    )

    h2(doc, "10 DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS")
    p(doc, "Com a contratação do serviço de capacitação, pretende-se:")
    bullet(doc, f"capacitar entre {VAGAS_MIN} e {VAGAS_MAX} profissionais de saúde de nível superior da APS na prática de Auriculoterapia;")
    bullet(doc, "assegurar qualificação clínica, domínio técnico e padronização em biossegurança para atendimento seguro no SUS;")
    bullet(doc, "expandir e qualificar a oferta regular de PICS nas UBS e nas 08 equipes eMulti em conformidade com a PNPIC e o PICSPel;")
    bullet(doc, "ampliar a resolutividade, a integralidade e a humanização do cuidado na Atenção Primária à Saúde de Pelotas;")
    bullet(
        doc,
        f"executar de forma eficiente, transparente e dentro do prazo a cota integral de {brl(COTA_SERVICOS)} "
        f"destinada a serviços de terceiros da Emenda Parlamentar nº {EMENDA}.",
    )

    h2(doc, "11 PROVIDÊNCIAS PRÉVIAS AO CONTRATO")
    bullet(doc, "aprovação do DFD nº [___]/2026, do presente ETP e do Termo de Referência pela autoridade competente;")
    bullet(doc, "realização da pesquisa formal de preços (art. 23 da Lei nº 14.133/2021) e consolidação do mapa comparativo;")
    bullet(doc, "confirmação da dotação orçamentária, projeto/atividade, fonte de recurso e subelemento (3.3.90.39) junto à Contabilidade / CCGOF;")
    bullet(doc, "verificação da previsão no PCA 2026 ou formalização da justificativa de inclusão excepcional;")
    bullet(doc, "definição e reserva prévia do auditório/espaço físico adequado na rede SMS para as aulas teóricas e práticas presenciais;")
    bullet(doc, "divulgação interna, seleção, inscrição e liberação formal de carga horária dos 30 a 50 profissionais participantes (incluindo lista de suplentes);")
    bullet(doc, "designação formal do gestor e do fiscal do contrato de serviços.")

    h2(doc, "12 CONTRATAÇÕES CORRELATAS/INTERDEPENDENTES")
    p(
        doc,
        f"No âmbito da mesma Emenda Parlamentar nº {EMENDA} (valor global de {brl(RECURSO_GLOBAL)}), a cota de "
        f"{brl(COTA_MATERIAIS)} destinada a materiais tramita em dois processos autônomos e correlatos, sem "
        "sobreposição de escopo com este processo de capacitação:",
    )
    bullet(
        doc,
        f"Processo de Materiais de Consumo (elemento 3.3.90.30 — estimado em {brl(ESTIMADO_CONSUMO)}): aquisição de "
        "15 itens de consumo (placas de auriculoterapia, pinças Adson, apalpadores Nogier, óleos essenciais, óleo "
        "carreador, frascos âmbar, inaladores nasais, bacias e protetores de escalda-pés, tapetes de yoga, plástico "
        "para plastificação e sal grosso), destinados às equipes eMulti, UBS e suporte operacional às práticas;",
    )
    bullet(
        doc,
        f"Processo de Materiais Permanentes e Equipamentos (elemento 4.4.90.52 — estimado em {brl(ESTIMADO_PERMANENTES)}): "
        "aquisição de 5 itens permanentes (caixas de som portáteis, garrafas térmicas, maletas organizadoras, jarras "
        "elétricas e plastificadora), destinados à estruturação das equipes eMulti e da coordenação do PICSPel.",
    )
    p(
        doc,
        f"A soma estimada dos três processos totaliza {brl(ESTIMADO_GLOBAL_3_PROC)}, integralmente compatível com o "
        f"teto de {brl(RECURSO_GLOBAL)} da emenda parlamentar (saldo residual de {brl(RECURSO_GLOBAL - ESTIMADO_GLOBAL_3_PROC)} "
        "na cota de materiais). Embora não haja dependência impeditiva para a licitação do curso, recomenda-se a "
        "sincronização dos cronogramas para que os insumos clínicos de auriculoterapia estejam disponíveis à rede "
        "concomitantemente à formação dos profissionais.",
    )

    h2(doc, "13 POSSÍVEIS IMPACTOS AMBIENTAIS")
    p(
        doc,
        "Por tratar-se de prestação de serviços técnico-educacionais presenciais, os impactos ambientais são reduzidos "
        "e gerenciáveis, relacionando-se principalmente ao consumo de papel/impressões de materiais didáticos e à "
        "geração de pequenos resíduos durante a Etapa II – Formação Prática Supervisionada (algodão, embalagens, "
        "adesivos e placas utilizadas em treinamento). Como medidas mitigadoras, deverá ser priorizado o fornecimento "
        "de material didático em formato digital (limitando-se impressões ao estritamente necessário, preferencialmente "
        "frente e verso) e assegurado o descarte adequado dos resíduos das aulas práticas conforme o Plano de "
        "Gerenciamento de Resíduos de Serviços de Saúde (PGRSS) do local e a Política Nacional de Resíduos Sólidos "
        "(Lei Federal nº 12.305/2010).",
    )

    h2(doc, "14 ANÁLISE DE RISCOS")
    riscos = [
        (
            "Risco 1: Atraso no início das aulas ou descumprimento do cronograma pedagógico",
            "Média",
            "Médio",
            "Adiamento da formação dos profissionais, atraso na expansão das PICS na APS e risco ao cronograma de execução da emenda parlamentar.",
            "Fixação de prazos objetivos no Termo de Referência (plano de curso em até 10 dias corridos e início em até 30 dias corridos após a Ordem de Serviço); reserva prévia do espaço físico pela SMS; acompanhamento sistemático pelo fiscal.",
            "Notificação formal à contratada, repactuação/reposição de aulas sem ônus adicional e aplicação das sanções administrativas previstas nos arts. 155 a 163 da Lei nº 14.133/2021 e no Decreto Municipal nº 6.799/2023.",
        ),
        (
            "Risco 2: Corpo docente, metodologia ou material didático em desacordo com as especificações técnicas",
            "Média",
            "Alto",
            "Comprometimento da qualidade formativa, falhas em protocolos de biossegurança e prejuízo à aplicação clínica segura na Atenção Primária.",
            "Exigência de atestado de capacidade técnica (art. 67) na habilitação; validação prévia do plano de curso e da qualificação/registro profissional dos docentes antes do início das aulas; avaliação contínua da execução pelo fiscal.",
            "Recusa da etapa ou entregável desconforme, notificação para substituição do docente ou adequação do material didático em até 5 (cinco) dias úteis sem ônus adicional, reposição da carga horária comprometida e aplicação de sanções cabíveis.",
        ),
        (
            "Risco 3: Licitação deserta ou fracassada / propostas acima do teto financeiro da emenda",
            "Média",
            "Alto",
            "Impossibilidade imediata de contratar a capacitação, atraso no cumprimento da meta da Emenda Parlamentar nº 43530013 e necessidade de repetição do procedimento.",
            "Pesquisa formal de preços atualizada e aderente ao mercado (art. 23); especificação técnica clara, objetiva e proporcional sem cláusulas restritivas; ampla divulgação no PNCP e comunicação a entidades de ensino da área.",
            "Revisão da pesquisa de preços e dos requisitos editalícios, republicação do certame ou avaliação fundamentada de contratação direta legalmente cabível (art. 75, II ou III, da Lei nº 14.133/2021, preenchidos os requisitos legais).",
        ),
        (
            "Risco 4: Baixa adesão, evasão de participantes ou falhas logísticas e de biossegurança nas aulas práticas",
            "Média",
            "Médio",
            "Subutilização das vagas contratadas, redução do número de profissionais certificados na rede ou intercorrências operacionais durante o treinamento prático.",
            "Seleção e inscrição prévia pela Diretoria de Atenção Primária com liberação formal de escala e lista de suplentes (faixa de 30 a 50 vagas); vistoria prévia do auditório; controle rigoroso de frequência (mínimo 75%); subdivisão prática em grupos de até 25 participantes e observância estrita ao PGRSS.",
            "Convocação imediata de suplentes no início do curso, adequação logística pela SMS e intervenção imediata da fiscalização para correção de procedimentos de biossegurança durante a etapa prática.",
        ),
    ]
    for titulo_r, prob, imp, dano, prev, cont in riscos:
        p(doc, titulo_r, bold=True, space_after=2, keep_with_next=True)
        p(doc, f"a) Probabilidade: {prob}.  b) Impacto: {imp}.", space_after=2)
        p(doc, f"c) Dano: {dano}", space_after=2)
        p(doc, f"d) Ação preventiva: {prev}", space_after=2)
        p(doc, f"e) Ação de contingência: {cont}", space_after=6)

    h2(doc, "15 DECLARAÇÃO DE VIABILIDADE DA CONTRATAÇÃO")
    p(
        doc,
        "A contratação do serviço presencial de capacitação em Auriculoterapia (80 horas, 30 a 50 vagas) mostra-se "
        "tecnicamente viável, operacionalmente exequível, juridicamente adequada à Lei Federal nº 14.133/2021 e "
        f"economicamente compatível com a cota de {brl(COTA_SERVICOS)} destinada a Outros Serviços de Terceiros – "
        f"Pessoa Jurídica da Emenda Parlamentar nº {EMENDA}. Há mercado competitivo de prestadores especializados e "
        "os riscos identificados são administráveis. Declara-se a viabilidade técnica, operacional, econômica e "
        "administrativa da contratação, recomendando-se o prosseguimento do processo com a aprovação do Termo de "
        "Referência, consolidação da pesquisa formal de preços e instauração do certame.",
    )

    h2(doc, "16 EQUIPE DE PLANEJAMENTO")
    p(
        doc,
        "Equipe de planejamento – Programa PICSPel, Diretoria de Atenção Primária e área administrativa da SMS, "
        "com apoio da equipe de educação permanente / gestão de espaços físicos da SMS para organização logística das aulas.",
    )
    assinaturas_padrao(doc)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


# ---------------------------------------------------------------------------
# 3) TERMO DE REFERÊNCIA (TR) — Padrão RP05 adaptado a Serviços / Cursos
# ---------------------------------------------------------------------------
def gerar_tr(path: Path):
    doc = new_doc()
    doc_title(doc, "TERMO DE REFERÊNCIA – PICSPel (SERVIÇOS DE CAPACITAÇÃO)")

    # SEÇÃO 1 — fiel ao modelo RP05
    h2(doc, "1 DO OBJETO, VALOR E CONDIÇÕES GERAIS DA CONTRATAÇÃO")
    p(
        doc,
        "O presente Termo de Referência tem por objeto a contratação de pessoa jurídica especializada para prestação "
        "de serviço técnico-educacional não continuado de capacitação presencial em Auriculoterapia destinada a "
        "profissionais de saúde de nível superior da rede municipal de saúde da Secretaria Municipal de Saúde de "
        f"Pelotas/RS, com capacidade de {VAGAS_MIN} (trinta) a {VAGAS_MAX} (cinquenta) vagas e carga horária total de "
        f"{CARGA_HORARIA_TOTAL} (oitenta) horas presenciais ({CARGA_TEORICA}h teóricas e {CARGA_PRATICA}h práticas "
        "supervisionadas), no âmbito do Programa Municipal de Práticas Integrativas e Complementares em Saúde "
        f"(PICSPel), custeada com recursos da Emenda Parlamentar nº {EMENDA}. Os quantitativos, especificações, "
        "valores estimados e condições gerais encontram-se estabelecidos neste instrumento.",
        bold_prefix="1.1 ",
    )

    desc_tr = (
        "SERVIÇO DE CAPACITAÇÃO PRESENCIAL EM AURICULOTERAPIA PARA PROFISSIONAIS DE SAÚDE DE NÍVEL SUPERIOR – "
        f"Carga horária total de {CARGA_HORARIA_TOTAL} (oitenta) horas presenciais, distribuídas em Etapa I – "
        f"Formação Teórica ({CARGA_TEORICA}h: PICS no SUS e PNPIC, reflexologia auricular, Medicina Tradicional "
        "Chinesa – MTC, bases biomédicas/neurofisiológicas e integração à Atenção Primária à Saúde) e Etapa II – "
        f"Formação Prática Supervisionada ({CARGA_PRATICA}h: mapeamento e localização de pontos, técnicas de "
        "aplicação, treinamento de habilidades clínicas, biossegurança, gerenciamento de resíduos – PGRSS e prática "
        f"supervisionada). Capacidade: {VAGAS_MIN} (trinta) a {VAGAS_MAX} (cinquenta) vagas (com subdivisão em até "
        "2 grupos de até 25 participantes ou proporcionalidade mínima de 1 supervisor para cada 25 participantes nas "
        "aulas práticas quando houver mais de 25 inscritos). Inclui coordenação pedagógica, corpo docente habilitado, "
        "material didático-pedagógico completo para todos os participantes, controle diário de frequência, avaliação de "
        "aprendizagem, relatório final de execução e emissão de certificados individuais aos concluintes."
    )
    styled_table(
        doc,
        ["Item", "Descrição", "CATSER", "Unid.", "Qtd.", "Valor unit. (R$)", "Valor total (R$)"],
        [
            [
                "01",
                desc_tr,
                "[A INFORMAR\nCompras.gov.br]",
                "Serviço\n(Turma:\n30 a 50\nvagas)",
                str(QTD_TURMAS),
                brl(VALOR_UNITARIO),
                brl(VALOR_TOTAL),
            ],
            ["", "VALOR TOTAL ESTIMADO", "", "", "", "", brl(VALOR_TOTAL)],
        ],
        col_widths_cm=[0.9, 7.6, 1.8, 1.7, 0.8, 2.1, 2.1],
        col_aligns=[
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ],
        font_size=8.0,
        bold_last_row=True,
        shade_last_row=True,
    )

    p(
        doc,
        "O objeto deste instrumento não se enquadra como bem ou serviço de luxo, nos termos da legislação vigente.",
        bold_prefix="1.2 ",
    )
    p(
        doc,
        "O serviço é caracterizado como serviço comum e não continuado, contratado por escopo definido (art. 6º, "
        "incisos XIII e XVII, da Lei Federal nº 14.133/2021), pois apresenta padrões de desempenho, conteúdo "
        "programático, carga horária e qualidade objetivamente definidos por meio de especificações usuais de mercado.",
        bold_prefix="1.3 ",
    )
    p(
        doc,
        "O prazo de vigência da contratação será de 180 (cento e oitenta) dias, contados da assinatura do termo de "
        "contrato ou da emissão do instrumento equivalente, na forma do art. 105 da Lei nº 14.133/2021. Por tratar-se "
        "de contrato por escopo, o prazo de vigência será automaticamente prorrogado quando o objeto não for concluído "
        "no período firmado, ressalvadas as providências cabíveis no caso de culpa da contratada, nos termos do art. "
        "111 da Lei nº 14.133/2021.",
        bold_prefix="1.4 ",
    )
    p(
        doc,
        "A contratação será executada sob o regime de empreitada por preço global (art. 6º, XXIX, da Lei nº "
        "14.133/2021), em contratação pontual por escopo (não sujeita ao Sistema de Registro de Preços), cobrindo "
        f"integralmente o atendimento de {VAGAS_MIN} (trinta) a {VAGAS_MAX} (cinquenta) participantes sem acréscimo "
        "de valor dentro dessa faixa.",
        bold_prefix="1.5 ",
    )
    p(
        doc,
        "O não parcelamento do objeto (item único / adjudicação global) justifica-se técnica e pedagogicamente, nos "
        "termos do art. 40, § 3º, e do art. 47, § 1º, da Lei nº 14.133/2021, pela necessidade de unidade de "
        "coordenação acadêmica, coerência metodológica entre as etapas teórica e prática, padronização de biossegurança "
        "e responsabilização integral pela certificação final dos concluintes.",
        bold_prefix="1.6 ",
    )
    p(
        doc,
        f"O valor total estimado é de {brl(VALOR_TOTAL)} (quarenta mil reais), correspondente ao teto da cota de "
        f"Outros Serviços de Terceiros – Pessoa Jurídica da Emenda Parlamentar nº {EMENDA}, conforme tabela acima e "
        "mapa comparativo de preços a ser anexado nos termos do art. 23 da Lei nº 14.133/2021.",
        bold_prefix="1.7 ",
    )
    p(
        doc,
        "Em razão de eventuais limitações na precisão das descrições do catálogo CATSER, poderão ocorrer divergências "
        "em relação às especificações deste Termo de Referência. Nesses casos, prevalecerão integralmente as "
        "especificações deste Termo.",
        bold_prefix="1.8 ",
    )

    # SEÇÃO 2 — fiel ao modelo RP05
    h2(doc, "2 DA FUNDAMENTAÇÃO DA CONTRATAÇÃO")
    p(
        doc,
        "A presente contratação visa atender às necessidades da Secretaria Municipal de Saúde de Pelotas/RS quanto à "
        "formação continuada e qualificação técnica de profissionais de nível superior da Atenção Primária à Saúde "
        "(UBS e equipes eMulti) em Auriculoterapia, em alinhamento à Política Nacional de Práticas Integrativas e "
        "Complementares (PNPIC – Portaria GM/MS nº 971/2006 e alterações) e ao Programa Municipal PICSPel, com lastro "
        f"orçamentário na Emenda Parlamentar nº {EMENDA} (cota de {brl(COTA_SERVICOS)} para serviços de terceiros – PJ). "
        "A fundamentação completa e o dimensionamento constam do Estudo Técnico Preliminar.",
        bold_prefix="2.1 ",
    )
    p(
        doc,
        "Justificativa para contratação pontual por escopo (não adoção do Sistema de Registro de Preços): a opção "
        "pela contratação pontual decorre de o objeto possuir escopo delimitado, carga horária certa (80 horas), "
        "público-alvo definido (30 a 50 profissionais da rede municipal) e vinculação direta ao teto financeiro da "
        f"Emenda Parlamentar nº {EMENDA}, afastando a hipótese de demanda eventual ou indeterminada típica do SRP. "
        "Registra-se ainda que não há sobreposição de escopo com os processos autônomos de Materiais de Consumo "
        f"({brl(ESTIMADO_CONSUMO)} – 3.3.90.30) e de Materiais Permanentes ({brl(ESTIMADO_PERMANENTES)} – 4.4.90.52), "
        f"custeados pela cota de {brl(COTA_MATERIAIS)} da mesma emenda.",
        bold_prefix="2.2 ",
    )
    p(
        doc,
        "A contratação está alinhada ao planejamento da Secretaria Municipal de Saúde / Diretoria de Atenção Primária. "
        "Previsão no Plano de Contratações Anual (PCA 2026): [A INFORMAR — nº do item no PCA ou justificativa de "
        f"inclusão excepcional vinculada à Emenda Parlamentar nº {EMENDA}].",
        bold_prefix="2.3 ",
    )

    # SEÇÃO 3 — fiel ao modelo RP05
    h2(doc, "3 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO CONSIDERADO O CICLO DE VIDA DO OBJETO")
    p(
        doc,
        "A solução consiste na prestação integral do serviço presencial de capacitação em Auriculoterapia (80 horas, "
        "para 30 a 50 profissionais de nível superior), contemplando todas as etapas de seu ciclo de vida: planejamento "
        "pedagógico e validação prévia do plano de curso, cronograma e corpo docente; execução presencial da Etapa I – "
        "Formação Teórica (48h) e da Etapa II – Formação Prática Supervisionada (32h, com organização em até 2 grupos "
        "de até 25 participantes ou proporcionalidade mínima de 1 supervisor para cada 25 participantes quando houver "
        "mais de 25 inscritos); fornecimento de material didático-pedagógico completo; controle nominal de frequência "
        "e avaliação de aprendizagem; recebimento provisório e definitivo; emissão e entrega dos certificados aos "
        "concluintes e do relatório final consolidado; bem como gestão e descarte ambientalmente adequado dos resíduos "
        "gerados nas atividades práticas conforme o PGRSS do local e a Política Nacional de Resíduos Sólidos. A "
        "fundamentação encontra-se pormenorizada no Estudo Técnico Preliminar.",
    )

    # SEÇÃO 4 — fiel ao modelo RP05
    h2(doc, "4 GESTÃO/ANÁLISE DE RISCOS")
    p(
        doc,
        "Os principais riscos identificados na fase de planejamento são: (I) atraso no início das aulas ou no "
        "cumprimento do cronograma pedagógico (probabilidade média, impacto médio); (II) corpo docente, metodologia "
        "ou material didático em desacordo com as especificações técnicas (probabilidade média, impacto alto); "
        "(III) licitação deserta ou fracassada / variação de preços acima do teto da emenda (probabilidade média, "
        "impacto alto); e (IV) baixa adesão, evasão de participantes ou falhas logísticas/de biossegurança nas aulas "
        "práticas (probabilidade média, impacto médio). Como medidas de prevenção e mitigação, serão observados a "
        "pesquisa formal de preços atualizada, a exigência de qualificação técnica proporcional, a validação prévia "
        "do plano de curso e dos currículos docentes pela coordenação do PICSPel, a reserva prévia do espaço físico "
        "pela SMS, a seleção com lista de suplentes, o pagamento condicionado à efetiva execução e aceite das etapas "
        "e a fiscalização rigorosa. Em caso de descumprimento, serão aplicadas as sanções previstas na Lei Federal nº "
        "14.133/2021, no Decreto Municipal nº 6.799/2023 e no instrumento contratual.",
    )

    # SEÇÃO 5 — fiel ao modelo RP05 (adaptada integralmente a serviços/cursos)
    h2(doc, "5 REQUISITOS DA CONTRATAÇÃO")
    p(
        doc,
        "Não serão exigidas certificações ou credenciamentos exclusivos que restrinjam indevidamente a competição, "
        "sendo admitidas instituições de ensino e pessoas jurídicas especializadas que atendam integralmente às "
        "especificações técnicas e pedagógicas deste Termo, preservando-se a isonomia e a competitividade.",
        bold_prefix="5.1 ",
    )
    p(doc, "Requisitos técnicos e pedagógicos específicos do curso:", bold_prefix="5.2 ")
    bullet(
        doc,
        f"Estrutura e carga horária: {CARGA_HORARIA_TOTAL} (oitenta) horas presenciais obrigatórias, divididas em "
        f"Etapa I – Formação Teórica ({CARGA_TEORICA} horas presenciais) e Etapa II – Formação Prática Supervisionada "
        f"({CARGA_PRATICA} horas presenciais).",
    )
    bullet(
        doc,
        "Conteúdo programático mínimo da Etapa I (48h teóricas): histórico e diretrizes da Política Nacional de "
        "Práticas Integrativas e Complementares (PNPIC) e das PICS no SUS; organização do PICSPel na Atenção Primária; "
        "fundamentos da reflexologia auricular; bases da Medicina Tradicional Chinesa (MTC) aplicadas à auriculoterapia; "
        "bases biomédicas, anatomia do pavilhão auricular e neurofisiologia; avaliação clínica, anamnese, indicações, "
        "contraindicações, intercorrências e registro assistencial no SUS.",
    )
    bullet(
        doc,
        "Conteúdo programático mínimo da Etapa II (32h práticas supervisionadas): inspeção e palpação auricular; "
        "mapeamento e localização de pontos auriculares; técnicas de estimulação e aplicação (sementes, cristais e "
        "protocolos clínicos compatíveis com a APS); treinamento prático de habilidades sob supervisão direta; "
        "protocolos de biossegurança, higienização, assepsia e descarte de resíduos conforme PGRSS; e discussão de "
        "casos clínicos aplicados à realidade das UBS e equipes eMulti.",
    )
    bullet(
        doc,
        f"Capacidade de turma e proporção de supervisão prática: atendimento de {VAGAS_MIN} (trinta) a {VAGAS_MAX} "
        "(cinquenta) profissionais de saúde de nível superior. Nas 32 horas práticas presenciais, caso o total de "
        "participantes inscritos seja superior a 25 (vinte e cinco), a contratada deverá organizar as práticas em "
        "2 (dois) grupos de até 25 participantes ou disponibilizar instrutores/supervisores práticos na proporção "
        "mínima de 1 (um) docente para cada 25 (vinte e cinco) participantes, sem ônus adicional para a SMS.",
    )
    bullet(
        doc,
        "Material didático-pedagógico e delimitação de insumos: a contratada deverá fornecer a todos os participantes "
        "material didático-pedagógico completo e atualizado (apostila/guia de estudos em formato digital e/ou impresso, "
        "mapas auriculares de referência, apresentações e roteiros de biossegurança), além dos materiais didáticos de "
        "demonstração utilizados pelo docente em sala. Os insumos clínicos de consumo adquiridos pela SMS no processo "
        "autônomo de Materiais de Consumo PICSPel darão suporte operacional à rede e às práticas, sem cobrança ou "
        "sobreposição de fornecimento neste contrato.",
    )
    bullet(
        doc,
        "Frequência, avaliação e certificação: a contratada manterá lista nominal de frequência assinada por turno "
        "e aplicará avaliação teórico-prática de aproveitamento. Farão jus ao certificado individual os participantes "
        "que obtiverem frequência mínima de 75% (setenta e cinco por cento) da carga horária total (60 horas) e "
        "aproveitamento mínimo satisfatório conforme critério aprovado no plano de curso. Os certificados (em formato "
        "digital verificável e/ou impresso, conforme definido com a coordenação do PICSPel) deverão conter: nome "
        "completo e CPF/matrícula do participante, título do curso, carga horária total (80h), período de realização, "
        "conteúdo programático no verso ou anexo, razão social e CNPJ da contratada e assinatura do responsável técnico.",
    )
    p(
        doc,
        "Corpo docente e coordenação técnica: os docentes/instrutores deverão possuir graduação de nível superior na "
        "área da saúde, registro ativo no respectivo conselho de classe (quando aplicável à categoria profissional) e "
        "formação específica (pós-graduação/curso de qualificação) ou experiência comprovada em Auriculoterapia/PICS "
        "e atividade docente. A relação nominal do coordenador pedagógico e dos docentes, acompanhada de currículos e "
        "documentos comprobatórios, deverá ser apresentada à coordenação do PICSPel em até 10 (dez) dias corridos "
        "após o recebimento da Ordem de Serviço para aprovação prévia. Eventual substituição de docente durante a "
        "execução dependerá de anuência prévia da SMS e manutenção de qualificação técnica equivalente ou superior.",
        bold_prefix="5.3 ",
    )
    p(
        doc,
        "Instalações físicas e recursos de apoio: as aulas teóricas e práticas presenciais serão realizadas em "
        "auditório ou espaço físico disponibilizado pela Secretaria Municipal de Saúde de Pelotas / Diretoria de "
        "Atenção Primária, no Município de Pelotas/RS, com mobiliário e recursos audiovisuais básicos do local. "
        "Recursos didáticos específicos do instrutor (computador portátil, modelos anatômicos auriculares didáticos, "
        "material pedagógico) serão de responsabilidade da contratada.",
        bold_prefix="5.4 ",
    )
    p(
        doc,
        "Qualificação técnica (art. 67 da Lei nº 14.133/2021): a licitante provisoriamente classificada em primeiro "
        "lugar deverá apresentar: (I) um ou mais atestados de capacidade técnica, emitidos por pessoa jurídica de "
        "direito público ou privado, comprovando a execução satisfatória de serviço de capacitação/curso presencial "
        "na área de Auriculoterapia ou Práticas Integrativas e Complementares em Saúde (PICS), com carga horária "
        "mínima igual ou superior a 40 (quarenta) horas (50% da carga horária contratada, nos termos do art. 67, § 2º, "
        "da Lei nº 14.133/2021); e (II) declaração formal de disponibilidade de corpo docente qualificado e compatível "
        "com o item 5.3 para execução integral do curso no Município de Pelotas/RS.",
        bold_prefix="5.5 ",
    )
    p(
        doc,
        "Proposta técnico-comercial e plano preliminar de curso: juntamente com a proposta ajustada ao lance final "
        "(ou quando solicitada pelo pregoeiro), a licitante provisoriamente vencedora deverá apresentar ementa/plano "
        "preliminar de curso compatível com os conteúdos e carga horária previstos no item 5.2, sob pena de "
        "desclassificação caso não atenda às especificações deste Termo.",
        bold_prefix="5.6 ",
    )
    p(
        doc,
        "Para fins de habilitação jurídica, fiscal, social, trabalhista e econômico-financeira (arts. 66, 68 e 69 "
        "da Lei nº 14.133/2021), a empresa vencedora deverá apresentar, quando solicitado:",
        bold_prefix="5.7 ",
    )
    bullet(doc, "Comprovante de inscrição e de situação cadastral ativa no CNPJ, com CNAE compatível com atividades de ensino/treinamento;")
    bullet(doc, "Contrato social, estatuto ou ato constitutivo consolidado vigente, devidamente registrado, com comprovação dos poderes do representante legal;")
    bullet(doc, "Certidões negativas ou positivas com efeito de negativas de regularidade fiscal e trabalhista: Fazenda Federal (inclusive contribuições previdenciárias/INSS), Fazenda Estadual, Fazenda Municipal do domicílio/sede, FGTS e Certidão Negativa de Débitos Trabalhistas (CNDT);")
    bullet(doc, "Declaração de que não emprega menor em condições vedadas, nos termos do inciso XXXIII do art. 7º da Constituição Federal;")
    bullet(doc, "Declaração de cumprimento das exigências de reserva de cargos para pessoa com deficiência e para reabilitado da Previdência Social, quando aplicável (art. 63, IV, da Lei nº 14.133/2021);")
    bullet(doc, "Certidão negativa de falência ou concordata/recuperação judicial expedida pelo distribuidor da sede da pessoa jurídica, dentro do prazo de validade (qualificação econômico-financeira proporcional ao valor e à natureza do contrato).")
    p(
        doc,
        "É vedada a subcontratação total do objeto ou da coordenação pedagógica e certificação. Admite-se apenas a "
        "contratação ou vinculação formal de docentes/instrutores pela pessoa jurídica contratada (vínculo "
        "trabalhista, societário ou prestação de serviços), permanecendo a contratada integral e exclusivamente "
        "responsável perante a Administração pela execução, qualidade pedagógica, encargos e certificação (art. 122 "
        "da Lei nº 14.133/2021).",
        bold_prefix="5.8 ",
    )
    p(
        doc,
        "O preço global contratado deverá contemplar todos os custos diretos e indiretos necessários à perfeita "
        "execução do objeto, incluindo honorários docentes, coordenação pedagógica, despesas de deslocamento, "
        "transporte, hospedagem e alimentação da equipe da contratada, elaboração e fornecimento de material didático, "
        "emissão de certificados, tributos, taxas e encargos trabalhistas, previdenciários e fiscais.",
        bold_prefix="5.9 ",
    )
    p(
        doc,
        "Critérios de sustentabilidade e biossegurança: a contratada deverá priorizar o uso racional de papel "
        "(disponibilizando material didático em meio digital e limitando impressões ao necessário, em frente e verso "
        "quando possível) e cumprir integralmente as normas de biossegurança e o Plano de Gerenciamento de Resíduos "
        "de Serviços de Saúde (PGRSS) do espaço físico durante a etapa prática presencial.",
        bold_prefix="5.10 ",
    )

    # SEÇÃO 6 — fiel ao modelo RP05
    h2(doc, "6 DA EXECUÇÃO DO OBJETO")
    p(
        doc,
        "Local de execução: os serviços presenciais (48h teóricas e 32h práticas supervisionadas) serão executados "
        "em auditório ou espaço físico disponibilizado pela Secretaria Municipal de Saúde de Pelotas / Diretoria de "
        "Atenção Primária, no Município de Pelotas/RS, cujo endereço exato e horários de funcionamento serão "
        "informados formalmente na Ordem de Serviço.",
        bold_prefix="6.1 ",
    )
    p(
        doc,
        "Prazos e cronograma de execução: (I) em até 10 (dez) dias corridos após o recebimento da Ordem de Serviço / "
        "Nota de Empenho, a contratada deverá submeter à coordenação do PICSPel o plano de curso detalhado, o "
        "cronograma de encontros e a comprovação documental do corpo docente; (II) o início efetivo das aulas deverá "
        "ocorrer em até 30 (trinta) dias corridos contados da emissão da Ordem de Serviço / Nota de Empenho; (III) a "
        "execução integral das 80 (oitenta) horas presenciais deverá ser concluída em até 90 (noventa) dias corridos "
        "a partir da data de início das aulas, em calendário previamente acordado com a Diretoria de Atenção Primária "
        "de modo a compatibilizar a liberação das equipes assistenciais; e (IV) o relatório final consolidado, os "
        "registros de frequência e avaliação e os certificados individuais deverão ser entregues em até 15 (quinze) "
        "dias úteis após o encerramento das aulas.",
        bold_prefix="6.2 ",
    )
    p(
        doc,
        "Caso ocorra impedimento justificado que afete alguma data prevista no cronograma aprovado, a contratada "
        "deverá comunicar formalmente a fiscalização com antecedência mínima de 5 (cinco) dias úteis (ou em até 48 "
        "horas na hipótese comprovada de caso fortuito ou força maior), submetendo proposta de reposição sem ônus "
        "adicional para análise e aprovação da coordenação do PICSPel.",
        bold_prefix="6.3 ",
    )
    p(
        doc,
        "Aulas eventualmente canceladas por responsabilidade da contratada ou executadas em desconformidade técnica "
        "deverão ser integralmente repostas em até 10 (dez) dias úteis (ou na primeira data compatível definida pela "
        "coordenação do PICSPel), e eventual substituição de docente ou correção de material didático/certificados "
        "deverá ocorrer em até 5 (cinco) dias úteis da notificação, sem qualquer custo adicional à Administração.",
        bold_prefix="6.4 ",
    )

    # SEÇÃO 7 — fiel ao modelo RP05 (adaptado de Gestão da Ata para Gestão do Contrato)
    h2(doc, "7 DA GESTÃO DO CONTRATO")
    p(
        doc,
        "O contrato (ou instrumento equivalente) deverá ser executado fielmente pelas partes, de acordo com as "
        "cláusulas avençadas, este Termo de Referência e as normas da Lei Federal nº 14.133/2021, respondendo cada "
        "parte pelas consequências de sua inexecução total ou parcial.",
        bold_prefix="7.1 ",
    )
    p(
        doc,
        f"A gestão e a fiscalização do contrato serão exercidas por servidores formalmente designados pela Secretaria "
        f"Municipal de Saúde de Pelotas/RS: Gestor do Contrato – {GESTOR[0]} (Matrícula nº {GESTOR[1]}); Fiscal do "
        f"Contrato – {FISCAL[0]} (Matrícula nº {FISCAL[1]}), nos termos do art. 117 da Lei nº 14.133/2021 e da "
        "regulamentação municipal vigente.",
        bold_prefix="7.2 ",
    )
    p(
        doc,
        "A execução do curso será acompanhada presencialmente e documentalmente pela fiscalização, mediante registro "
        "próprio de ocorrências, conferência de listas de frequência, verificação do conteúdo ministrado e validação "
        "dos entregáveis pedagógicos. As comunicações entre o órgão e a contratada serão realizadas por escrito "
        "sempre que o ato exigir, admitindo-se mensagem eletrônica institucional.",
        bold_prefix="7.3 ",
    )
    p(
        doc,
        "O descumprimento total ou parcial das obrigações assumidas caracterizará inadimplência, sujeitando a "
        "contratada às sanções administrativas previstas nos arts. 155 a 163 da Lei Federal nº 14.133/2021 e no "
        "Decreto Municipal nº 6.799/2023 (advertência, multa moratória e/ou compensatória, impedimento de licitar e "
        "contratar e declaração de inidoneidade), assegurados o contraditório e a ampla defesa, sem prejuízo de "
        "eventual responsabilidade civil e criminal.",
        bold_prefix="7.4 ",
    )
    p(
        doc,
        "Garantia de execução e reajuste: dispensa-se a exigência de garantia contratual (art. 96 da Lei nº "
        "14.133/2021) em razão do pagamento exclusivamente posterior à execução e ao aceite definitivo das etapas. "
        "Os preços contratados são fixos e irreajustáveis durante o prazo previsto de execução; ocorrendo prorrogação "
        "excepcional superior a 12 (doze) meses sem culpa da contratada, aplicar-se-á o IPCA (art. 25, § 7º, da Lei "
        "nº 14.133/2021).",
        bold_prefix="7.5 ",
    )

    # SEÇÃO 8 — fiel ao modelo RP05 (adaptado ao art. 140, I, da Lei 14.133/2021 — serviços)
    h2(doc, "8 RECEBIMENTO DO OBJETO")
    p(
        doc,
        "Os serviços serão recebidos provisoriamente, mediante termo detalhado, no prazo de até 7 (sete) dias úteis "
        "contados da conclusão de cada etapa prevista para medição (Etapa I – Formação Teórica de 48h; e Etapa II – "
        "Formação Prática Supervisionada de 32h acompanhada do relatório final e certificados), pelo fiscal do "
        "contrato, após verificação documental e técnica do cumprimento das exigências contratuais (listas de presença "
        "assinadas, diário de conteúdo ministrado, entrega do material didático e registros de avaliação), nos termos "
        "do art. 140, inciso I, alínea 'a', da Lei nº 14.133/2021.",
        bold_prefix="8.1 ",
    )
    p(
        doc,
        "Os serviços ou entregáveis (aulas, material didático, relatórios ou certificados) poderão ser rejeitados, "
        "no todo ou em parte, quando em desacordo com as especificações constantes neste Termo de Referência e no "
        "plano de curso aprovado, devendo ser corrigidos, refeitos ou substituídos no prazo de até 5 (cinco) dias "
        "úteis a contar da notificação (ou até 10 dias úteis / data pactuada para reposição presencial de aulas), "
        "às custas exclusivas da contratada, sem prejuízo da aplicação das penalidades cabíveis.",
        bold_prefix="8.2 ",
    )
    p(
        doc,
        "Os serviços serão recebidos definitivamente no prazo de até 8 (oito) dias úteis, contados do recebimento "
        "provisório, pelo gestor do contrato (ou comissão designada), mediante termo detalhado que comprove o "
        "atendimento integral das exigências contratuais (art. 140, inciso I, alínea 'b', da Lei nº 14.133/2021), "
        "incluindo, na etapa final, a conferência do relatório consolidado da turma e da entrega dos certificados "
        "individuais a todos os profissionais concluintes.",
        bold_prefix="8.3 ",
    )

    # SEÇÃO 9 — fiel ao modelo RP05 (adaptado a medição por etapa de serviço)
    h2(doc, "9 DOS CRITÉRIOS DE MEDIÇÃO E PAGAMENTO")
    p(
        doc,
        "A medição e o pagamento serão vinculados exclusivamente à efetiva execução e ao recebimento definitivo das "
        "etapas do serviço, sendo vedado o pagamento antecipado (art. 145 da Lei nº 14.133/2021). O pagamento será "
        "realizado em 2 (duas) parcelas, conforme o cumprimento comprovado do cronograma:",
        bold_prefix="9.1 ",
    )
    bullet(
        doc,
        "1ª Parcela – 50% (cinquenta por cento) do valor global contratado: devida após a conclusão integral e o "
        "recebimento definitivo da Etapa I – Formação Teórica (48 horas presenciais), mediante entrega do material "
        "didático completo, das listas de frequência assinadas e do registro de conteúdo ministrado;",
    )
    bullet(
        doc,
        "2ª Parcela – 50% (cinquenta por cento) do valor global contratado: devida após a conclusão integral e o "
        "recebimento definitivo da Etapa II – Formação Prática Supervisionada (32 horas presenciais), mediante "
        "entrega das listas de frequência, dos registros de avaliação de aprendizagem, do relatório final consolidado "
        "da capacitação e dos certificados individuais de todos os participantes concluintes.",
    )
    p(
        doc,
        "O pagamento de cada parcela será efetuado após o atesto da Nota Fiscal/Fatura de Serviços pelo fiscal e "
        "gestor do contrato e a verificação da regularidade fiscal e trabalhista da contratada, mediante crédito em "
        "conta bancária de titularidade da contratada, observadas as retenções tributárias legais aplicáveis (ISSQN, "
        "INSS e IRRF na fonte, quando couber) e o cronograma financeiro fixado pela Secretaria Municipal da Fazenda "
        "de Pelotas.",
        bold_prefix="9.2 ",
    )

    # SEÇÃO 10 — fiel ao modelo RP05
    h2(doc, "10 DA FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR")
    p(
        doc,
        "O prestador de serviço será selecionado por meio de licitação na modalidade Pregão, em sua forma eletrônica, "
        "nos termos dos arts. 28, I, e 29 da Lei Federal nº 14.133/2021, por possuir o objeto especificações usuais "
        "de mercado e padrões de desempenho e qualidade objetivamente definidos neste Termo (ressalvada a avaliação "
        "administrativa sobre eventual utilização de Dispensa com Disputa Eletrônica, art. 75, II, somente se "
        "comprovada formalmente a ausência de fracionamento de despesas de mesma natureza no exercício de 2026).",
        bold_prefix="10.1 ",
    )
    p(
        doc,
        "Critério de julgamento: menor preço (item único / preço global da turma), nos termos do art. 33, inciso I, "
        "da Lei nº 14.133/2021, observados os requisitos de habilitação jurídica, fiscal, trabalhista, econômico-"
        "financeira e técnico-operacional/profissional previstos na Seção 5.",
        bold_prefix="10.2 ",
    )
    p(
        doc,
        "Estimativa de preços: realizada nos termos do art. 23 da Lei nº 14.133/2021, conforme mapa comparativo de "
        "preços anexo [A INFORMAR — consolidar fontes (PNCP/Painel de Preços, contratações similares, cotações "
        f"diretas) e metodologia estatística (mediana/média) antes da publicação do edital], observado o teto de "
        f"{brl(VALOR_TOTAL)} da cota de serviços da Emenda Parlamentar nº {EMENDA}.",
        bold_prefix="10.3 ",
    )

    # SEÇÃO 11 — fiel ao modelo RP05
    h2(doc, "11 DAS OBRIGAÇÕES DO CONTRATANTE")
    obrig_contratante = [
        "Exigir o integral cumprimento, pela contratada, de todas as obrigações assumidas neste Termo e no contrato;",
        "Disponibilizar auditório ou espaço físico adequado no Município de Pelotas/RS, em condições de uso e nos horários pactuados, para a realização das aulas teóricas e práticas presenciais;",
        "Selecionar, inscrever e divulgar a relação dos 30 (trinta) a 50 (cinquenta) profissionais de saúde de nível superior participantes, inclusive lista de suplentes;",
        "Analisar e aprovar, no prazo pactuado, o plano de curso detalhado, o cronograma de encontros e a qualificação do corpo docente apresentados pela contratada;",
        "Emitir a Nota de Empenho e a Ordem de Serviço, prestando à contratada as informações necessárias à perfeita execução do objeto;",
        "Exercer o acompanhamento e a fiscalização por servidores formalmente designados (gestor e fiscal), anotando em registro próprio todas as ocorrências e falhas detectadas;",
        "Notificar a contratada, por escrito, sobre imperfeições, falhas ou irregularidades constatadas na execução, fixando prazo para correção ou reposição;",
        "Receber provisória e definitivamente o objeto nos prazos e condições estabelecidos e efetuar os pagamentos correspondentes;",
        "Aplicar à contratada as sanções administrativas previstas na Lei nº 14.133/2021, no Decreto Municipal nº 6.799/2023 e no edital, quando couber;",
        "Não praticar atos de ingerência na administração da contratada, emitindo decisão fundamentada sobre solicitações e reclamações relacionadas à execução contratual;",
        "Fiscalizar a manutenção, pela contratada, de todas as condições de habilitação durante a vigência do contrato.",
    ]
    for item_ob in obrig_contratante:
        bullet(doc, item_ob)

    # SEÇÃO 12 — fiel ao modelo RP05
    h2(doc, "12 DAS OBRIGAÇÕES DO CONTRATADO")
    obrig_contratado = [
        "Cumprir integralmente todas as obrigações assumidas, executando as 80 (oitenta) horas presenciais (48h teóricas e 32h práticas supervisionadas) conforme o plano de curso aprovado, assumindo os riscos e despesas inerentes à boa execução do objeto;",
        "Alocar corpo docente qualificado, com formação de nível superior na área da saúde, registro profissional compatível e experiência comprovada em Auriculoterapia/PICS, submetendo previamente os currículos à coordenação do PICSPel e vedada a substituição de docentes sem anuência prévia da SMS;",
        "Garantir, durante as 32 horas práticas supervisionadas, a subdivisão em 2 (dois) grupos de até 25 participantes ou a presença mínima de 1 (um) supervisor prático para cada 25 (vinte e cinco) participantes quando o número de inscritos superar 25, sem custo adicional;",
        "Fornecer material didático-pedagógico completo e atualizado a todos os participantes, bem como os materiais didáticos de demonstração do docente;",
        "Manter controle diário nominal de frequência, aplicar avaliação de aprendizagem, elaborar relatório final consolidado e emitir e entregar certificados individuais válidos a todos os concluintes no prazo de até 15 (quinze) dias úteis após o encerramento das aulas;",
        "Atender prontamente às determinações do fiscal e do gestor do contrato ou da autoridade superior, reparando, corrigindo ou repondo, às suas expensas e nos prazos fixados, quaisquer aulas, materiais ou documentos em desacordo;",
        "Cumprir rigorosamente as normas de biossegurança, conduta ética profissional e o Plano de Gerenciamento de Resíduos de Serviços de Saúde (PGRSS) do local durante as aulas práticas presenciais;",
        "Responsabilizar-se integralmente pelo pagamento de honorários, salários, tributos, encargos trabalhistas, previdenciários, fiscais e comerciais, bem como por despesas de deslocamento, hospedagem e alimentação de sua equipe docente e administrativa;",
        "Responsabilizar-se pelos vícios e danos decorrentes da execução do serviço, nos termos da legislação civil e do Código de Defesa do Consumidor (Lei nº 8.078/1990), no que couber;",
        "Observar integralmente a Lei Geral de Proteção de Dados Pessoais (Lei nº 13.709/2018 – LGPD) no tratamento de dados cadastrais e registros de frequência/certificação dos servidores municipais participantes;",
        "Não permitir a utilização de trabalho de menor de dezesseis anos, exceto na condição de aprendiz a partir dos quatorze anos, nem de menor de dezoito anos em trabalho noturno, perigoso ou insalubre (art. 7º, XXXIII, da CF);",
        "Manter, durante toda a execução do contrato, todas as condições de habilitação e qualificação exigidas na licitação.",
    ]
    for item_ob in obrig_contratado:
        bullet(doc, item_ob)

    # SEÇÃO 13 — fiel ao modelo RP05
    h2(doc, "13 DA ADEQUAÇÃO ORÇAMENTÁRIA")
    p(
        doc,
        "Por tratar-se de contratação pontual por escopo (e não de Sistema de Registro de Preços), a indicação da "
        "dotação orçamentária e a emissão da respectiva nota de reserva/pré-empenho constituem condição prévia à "
        "contratação, nos termos do art. 150 da Lei Federal nº 14.133/2021. As despesas decorrentes desta contratação "
        f"correrão à conta dos recursos da Emenda Parlamentar nº {EMENDA} (valor global de {brl(RECURSO_GLOBAL)}; "
        f"cota destinada a serviços de terceiros – PJ: {brl(COTA_SERVICOS)}):",
    )
    styled_table(
        doc,
        ["Campo Orçamentário", "Especificação"],
        [
            ["Emenda Parlamentar", f"Nº {EMENDA} (cota de Outros Serviços de Terceiros – PJ: {brl(COTA_SERVICOS)})"],
            ["Projeto/Atividade", "[A INFORMAR — Atenção Primária / PICSPel]"],
            ["Fonte de Recurso", f"[A INFORMAR — vinculada à Emenda Parlamentar nº {EMENDA}]"],
            ["Reduzido / Dotação", "[A INFORMAR — Contabilidade / CCGOF]"],
            ["Classificação da Despesa", "3.3.90.39 – Outros Serviços de Terceiros – Pessoa Jurídica (subelemento [A INFORMAR — ex.: 3.3.90.39.48])"],
            ["Nota de Reserva Orçamentária", "[A INFORMAR antes da publicação do edital / contratação]"],
            ["Valor Total Estimado", brl(VALOR_TOTAL)],
        ],
        col_widths_cm=[5.6, 11.4],
        col_aligns=[WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT],
        font_size=9.0,
    )

    # SEÇÃO 14 — fiel ao modelo RP05
    h2(doc, "14 EQUIPE DE PLANEJAMENTO")
    assinaturas_padrao(
        doc,
        incluir_cabecalho_secao=False,
        texto_equipe=(
            "Equipe de planejamento – Departamento/Serviço demandante: Programa Municipal de Práticas Integrativas "
            "e Complementares em Saúde (PICSPel), Diretoria de Atenção Primária e Diretoria Administrativa e Financeira da SMS."
        ),
    )

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)


# ---------------------------------------------------------------------------
# Validação programática integral dos 3 documentos .docx
# ---------------------------------------------------------------------------
def _extract_all_text(doc: Document) -> str:
    parts: list[str] = []
    for para in doc.paragraphs:
        if para.text.strip():
            parts.append(para.text)
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                if cell.text.strip():
                    parts.append(cell.text)
    return "\n".join(parts)


def validar_documentos(dfd_path: Path, etp_path: Path, tr_path: Path) -> None:
    print("\n=== VALIDAÇÃO PROGRAMÁTICA DOS DOCUMENTOS DOCX ===")
    expected_common = [
        "43530013",
        "R$ 40.000,00",
        "3.3.90.39",
        "80",
        "48",
        "32",
        "30",
        "50",
        "Dariane Lima Portela",
        "44151",
        "Aline Kohler Geppert",
        "29371",
        "Angela Moreira Vitória",
        "14.133/2021",
    ]
    # Termos proibidos (resíduos de copy-paste de Materiais/RP05 que NÃO podem aparecer)
    forbidden_terms = [
        "Teto da cota de materiais na emenda",
        "com apoio do Almoxarifado",
        "Controle patrimonial e armazenamento inadequado",
        "extravio de permanentes",
        "tombamento dos permanentes",
        "lacuna em kit de equipe",
        "realocação interna de materiais",
        "CÂNULA DE GUEDEL",
        "TUBO ENDOTRAQUEAL",
        "CATETER DE ASPIRAÇÃO",
        "contratação direta por turma",
    ]

    checks = [
        (
            "DFD",
            dfd_path,
            [
                "DOCUMENTO DE FORMALIZAÇÃO DA DEMANDA",
                "1 OBJETO",
                "1.1 TIPO DE OBJETO",
                "( X ) Serviço não continuado",
                "2 FORMA DE CONTRATAÇÃO SUGERIDA",
                "( X ) Pregão Eletrônico",
                "3 JUSTIFICATIVA DA CONTRATAÇÃO",
                "4 DESCRIÇÃO DO OBJETO",
                "5 DOTAÇÃO ORÇAMENTÁRIA",
                "6 LOCAL DE EXECUÇÃO DOS SERVIÇOS",
                "7 PREVISÃO DE INÍCIO E PRAZO DE EXECUÇÃO",
                "Teto da cota de serviços na emenda",
                "R$ 47.200,00",
                "R$ 11.200,00",
            ],
            3,  # tabela objeto + tabela dotação + tabela assinaturas
        ),
        (
            "ETP",
            etp_path,
            [
                "ESTUDO TÉCNICO PRELIMINAR – PICSPel (SERVIÇOS DE CAPACITAÇÃO)",
                "1 IDENTIFICAÇÃO DO PROCESSO",
                "2 DESCRIÇÃO DA NECESSIDADE",
                "3 DA CONTRATAÇÃO E DO PLANEJAMENTO",
                "4 REQUISITOS DA CONTRATAÇÃO",
                "5 ESTIMATIVA DAS QUANTIDADES",
                "6 LEVANTAMENTO DE MERCADO",
                "7 ESTIMATIVA DO VALOR DA CONTRATAÇÃO",
                "8 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO",
                "9 JUSTIFICATIVA PARA NÃO PARCELAMENTO",
                "10 DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS",
                "11 PROVIDÊNCIAS PRÉVIAS AO CONTRATO",
                "12 CONTRATAÇÕES CORRELATAS/INTERDEPENDENTES",
                "13 POSSÍVEIS IMPACTOS AMBIENTAIS",
                "14 ANÁLISE DE RISCOS",
                "15 DECLARAÇÃO DE VIABILIDADE DA CONTRATAÇÃO",
                "16 EQUIPE DE PLANEJAMENTO",
                "R$ 98.400,00",
                "R$ 100.000,00",
            ],
            3,  # tabela quantitativos + tabela natureza + tabela assinaturas
        ),
        (
            "TR",
            tr_path,
            [
                "TERMO DE REFERÊNCIA – PICSPel (SERVIÇOS DE CAPACITAÇÃO)",
                "1 DO OBJETO, VALOR E CONDIÇÕES GERAIS DA CONTRATAÇÃO",
                "2 DA FUNDAMENTAÇÃO DA CONTRATAÇÃO",
                "3 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO CONSIDERADO O CICLO DE VIDA DO OBJETO",
                "4 GESTÃO/ANÁLISE DE RISCOS",
                "5 REQUISITOS DA CONTRATAÇÃO",
                "6 DA EXECUÇÃO DO OBJETO",
                "7 DA GESTÃO DO CONTRATO",
                "8 RECEBIMENTO DO OBJETO",
                "9 DOS CRITÉRIOS DE MEDIÇÃO E PAGAMENTO",
                "10 DA FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR",
                "11 DAS OBRIGAÇÕES DO CONTRATANTE",
                "12 DAS OBRIGAÇÕES DO CONTRATADO",
                "13 DA ADEQUAÇÃO ORÇAMENTÁRIA",
                "14 EQUIPE DE PLANEJAMENTO",
                "CATSER",
                "6.799/2023",
                "1ª Parcela – 50%",
                "2ª Parcela – 50%",
            ],
            3,  # tabela objeto + tabela dotação + tabela assinaturas
        ),
    ]

    # Verificação matemática das somas do PICSPel
    assert abs(VALOR_TOTAL - 40_000.00) < 1e-6
    assert abs(ESTIMADO_MATERIAIS_TOTAL - 58_400.00) < 1e-6
    assert abs(ESTIMADO_GLOBAL_3_PROC - 98_400.00) < 1e-6
    assert ESTIMADO_GLOBAL_3_PROC <= RECURSO_GLOBAL
    assert CARGA_TEORICA + CARGA_PRATICA == CARGA_HORARIA_TOTAL
    assert abs(CUSTO_POR_HORA - 500.00) < 1e-6
    assert abs(CUSTO_POR_VAGA_MAX - 800.00) < 1e-6

    for label, fpath, required_specific, min_tables in checks:
        assert fpath.exists() and fpath.stat().st_size > 10_000, f"Arquivo inválido: {fpath}"
        doc = Document(str(fpath))
        full_text = _extract_all_text(doc)
        header_text = "\n".join(p.text for p in doc.sections[0].header.paragraphs)

        # Cabeçalho oficial + imagem do brasão
        assert "ESTADO DO RIO GRANDE DO SUL" in header_text
        assert "MUNICÍPIO DE PELOTAS" in header_text
        assert "SECRETARIA MUNICIPAL DE SAÚDE" in header_text
        header_xml = doc.sections[0].header._element.xml
        assert "blip" in header_xml, f"{label}: imagem do brasão ausente no cabeçalho!"

        # Rodapé Página X de Y
        footer_xml = doc.sections[0].footer._element.xml
        assert "PAGE" in footer_xml and "NUMPAGES" in footer_xml, f"{label}: paginação ausente no rodapé!"

        # Tabelas e soma na tabela principal de item
        assert len(doc.tables) >= min_tables, f"{label}: esperado >= {min_tables} tabelas, obtido {len(doc.tables)}"
        t0 = doc.tables[0]
        last_row_text = " | ".join(c.text.strip() for c in t0.rows[-1].cells)
        assert "VALOR TOTAL ESTIMADO" in last_row_text and "R$ 40.000,00" in last_row_text, (
            f"{label}: linha de total divergente -> {last_row_text}"
        )
        # Checar consistência Qtd × Valor Unit = Valor Total na tabela de objeto
        item_row = t0.rows[1].cells
        assert item_row[0].text.strip() == "01"
        assert "R$ 40.000,00" in item_row[-2].text and "R$ 40.000,00" in item_row[-1].text

        # Campos obrigatórios comuns e específicos
        for term in expected_common + required_specific:
            assert term in full_text, f"{label}: termo obrigatório ausente -> {term!r}"

        # Ausência de resíduos de copy-paste
        for bad in forbidden_terms:
            assert bad.lower() not in full_text.lower(), f"{label}: termo proibido encontrado -> {bad!r}"

        print(
            f"  [OK] {label}: {fpath.name} | {fpath.stat().st_size:,} bytes | "
            f"{len(doc.paragraphs)} parágrafos | {len(doc.tables)} tabelas | brasão+paginação OK"
        )

    print("Todas as validações de conteúdo, estrutura, tabelas, somas e integridade foram aprovadas!")


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    dfd_path = OUT_DIR / "PICSPel_Servicos_DFD_Revisado.docx"
    etp_path = OUT_DIR / "PICSPel_Servicos_ETP_Revisado.docx"
    tr_path = OUT_DIR / "PICSPel_Servicos_TR.docx"

    gerar_dfd(dfd_path)
    gerar_etp(etp_path)
    gerar_tr(tr_path)
    validar_documentos(dfd_path, etp_path, tr_path)


if __name__ == "__main__":
    main()
