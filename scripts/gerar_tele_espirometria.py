#!/usr/bin/env python3
"""Gera DFD e ETP para solicitação de diárias e passagens — Projeto Tele-espirometria.

Processo SIS nº 15912/2026 — Servidora: Elem Lira Rodrigues (Matrícula 44777, Enfermeira).
Treinamento presencial em Recife/PE (15/12/2026) — CTS-HC/UFMG / SAPS-MS / Novo PAC.
Padrão documental oficial SARH/SMS Pelotas (Lei 14.133/2021, Decreto Municipal 6.799/2023
e Lei Municipal nº 6.621/2018).
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "docs" / "06-TELE-ESPIROMETRIA-DIARIAS-PASSAGENS"
BRASAO = ROOT / "docs" / "assets" / "brasao_rp_header.png"
if not BRASAO.exists():
    BRASAO = ROOT / "docs" / "assets" / "brasao_pelotas.png"

# Dados oficiais do Processo SIS nº 15912/2026 e anexos
SIS_NUM = "15912/2026"
SIS_ABERTURA = "05/11/2026"
UNIDADE = "Secretaria Municipal de Saúde – SMS / Diretoria de Atenção Primária / Departamento de Saúde Digital"
ENDERECO_UNIDADE = "Rua Voluntários da Pátria, nº 1428, Sala 501 — Centro, Pelotas/RS"
SOLICITANTE = "Angela Roberta Alves de Lima"
SERVIDORA = "Elem Lira Rodrigues"
MATRICULA_SERVIDORA = "44777"
CARGO_SERVIDORA = "Enfermeira"

SAIDA_DATA = "14/12/2026"
SAIDA_HORA = "06h00"
RETORNO_DATA = "16/12/2026"
RETORNO_HORA = "18h00"
EVENTO_DATA = "15/12/2026"
EVENTO_HORARIO = "08h00 às 17h45 (credenciamento 30 minutos antes — 07h30)"
EVENTO_LOCAL = (
    "Sede da Secretaria de Saúde do Estado de Pernambuco (SES-PE) — "
    "Rua Vinte e Quatro de Agosto, nº 209, bairro Santo Amaro, Recife/PE, CEP 50040-190"
)
PROMOTOR = (
    "Centro de Telessaúde do Hospital das Clínicas da Universidade Federal de Minas Gerais, "
    "Filial EBSERH (CTS-HC/UFMG), CNPJ nº 17.217.985/0034-72, em parceria com a Secretaria de "
    "Atenção Primária à Saúde do Ministério da Saúde (SAPS/MS)"
)
COORDENADOR_PROJETO = "Prof. Dr. Antonio Luiz Pinho Ribeiro"

# Estimativas referenciais (pesquisa de mercado 2026 + Lei Municipal nº 6.621/2018)
VAL_AEREA = 2350.00  # Média dez/2026 POA <-> REC (Mundi/KAYAK nov/2026: R$ 2.025 a R$ 2.393)
VAL_RODOV_UNIT = 131.35  # Expresso Embaixador executivo 2026 (convencional: R$ 111,96; DAER/RS)
VAL_RODOV_TOTAL = VAL_RODOV_UNIT * 2  # R$ 262,70
SUBTOTAL_PASSAGENS = VAL_AEREA + VAL_RODOV_TOTAL  # R$ 2.612,70

URM_REF = 111.27  # Último valor oficial publicado localizado (Decreto nº 7.021/2025) — atualizar p/ dez/2026
DIARIAS_INTEGRAIS_QTD = 2
DIARIAS_INTEGRAIS_URM = 4  # Nível 02 - Outros Estados sem transporte oficial (Lei 6.621/2018)
DIARIA_ESTADA_QTD = 1
DIARIA_ESTADA_URM = 2  # Nível 02 - Outros Estados sem transporte oficial (Lei 6.621/2018)
TOTAL_URMS = DIARIAS_INTEGRAIS_QTD * DIARIAS_INTEGRAIS_URM + DIARIA_ESTADA_QTD * DIARIA_ESTADA_URM  # 10 URMs
SUBTOTAL_DIARIAS_REF = TOTAL_URMS * URM_REF  # R$ 1.112,70 (referência mínima)
TOTAL_GLOBAL_REF = SUBTOTAL_PASSAGENS + SUBTOTAL_DIARIAS_REF  # R$ 3.725,40

LIMITE_DISPENSA_2026 = 65492.11  # Decreto Federal nº 12.807/2025 (art. 75, II, Lei 14.133/2021)

DIRETORA = ("Cândida Silveira Rodrigues", "Diretora de Atenção Primária")
SECRETARIA = ("Angela Moreira Vitória", "Secretária Municipal de Saúde")


def brl(v: float) -> str:
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _font(run, *, name="Arial", size=11, bold=False, italic=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    run.italic = italic


def _shade(cell, fill="EDEDED"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


def _cell_margins(cell, top=60, bottom=60, left=100, right=100):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = OxmlElement("w:tcMar")
    for edge, val in (("top", top), ("bottom", bottom), ("left", left), ("right", right)):
        node = OxmlElement(f"w:{edge}")
        node.set(qn("w:w"), str(val))
        node.set(qn("w:type"), "dxa")
        tcMar.append(node)
    tcPr.append(tcMar)


def _page_fields(paragraph):
    def _fld(instr):
        run = paragraph.add_run()
        _font(run, size=9)
        r = run._r
        fc1 = OxmlElement("w:fldChar")
        fc1.set(qn("w:fldCharType"), "begin")
        r.append(fc1)
        run2 = paragraph.add_run()
        _font(run2, size=9)
        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = instr
        run2._r.append(it)
        run3 = paragraph.add_run()
        _font(run3, size=9)
        fc2 = OxmlElement("w:fldChar")
        fc2.set(qn("w:fldCharType"), "end")
        run3._r.append(fc2)

    run = paragraph.add_run("Página ")
    _font(run, size=9)
    _fld("PAGE")
    run = paragraph.add_run(" de ")
    _font(run, size=9)
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
    sec.header_distance = Cm(1.0)
    sec.footer_distance = Cm(0.8)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    style.font.size = Pt(11)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    style.paragraph_format.space_after = Pt(5)

    header = sec.header
    header.is_linked_to_previous = False
    header.paragraphs[0].clear()
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.space_after = Pt(2)
    if BRASAO.exists():
        run = hp.add_run()
        run.add_picture(str(BRASAO), width=Cm(1.65))
    for text, bold in [
        ("ESTADO DO RIO GRANDE DO SUL", True),
        ("MUNICÍPIO DE PELOTAS", True),
        ("SECRETARIA MUNICIPAL DE SAÚDE", True),
    ]:
        para = header.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        r = para.add_run(text)
        _font(r, size=10, bold=bold)

    line = header.add_paragraph()
    line.paragraph_format.space_before = Pt(2)
    line.paragraph_format.space_after = Pt(0)
    pPr = line._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)

    footer = sec.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.clear()
    _page_fields(fp)
    return doc


def title(doc: Document, main_text: str, sub_text: str | None = None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(4 if sub_text else 10)
    r = p.add_run(main_text)
    _font(r, size=12, bold=True)
    if sub_text:
        p2 = doc.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after = Pt(10)
        r2 = p2.add_run(sub_text)
        _font(r2, size=10, italic=True)


def h2(doc: Document, text: str, legal_ref: str | None = None):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(3 if legal_ref else 5)
    r = p.add_run(text)
    _font(r, size=11, bold=True)
    if legal_ref:
        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(5)
        r2 = p2.add_run(f"Fundamentação: {legal_ref}")
        _font(r2, size=9, italic=True)


def para(doc: Document, text: str, *, bold_prefix: str | None = None, justify: bool = True, size: int = 11):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT
    if bold_prefix:
        r1 = p.add_run(bold_prefix)
        _font(r1, size=size, bold=True)
    r2 = p.add_run(text)
    _font(r2, size=size, bold=False)
    return p


def bullet(doc: Document, text: str, *, bold_prefix: str | None = None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(0.6)
    p.paragraph_format.space_after = Pt(3)
    r0 = p.add_run("• ")
    _font(r0, size=10.5, bold=True)
    if bold_prefix:
        r1 = p.add_run(bold_prefix)
        _font(r1, size=10.5, bold=True)
    r2 = p.add_run(text)
    _font(r2, size=10.5)
    return p


def checkbox(doc: Document, text: str, checked: bool = False):
    mark = "X" if checked else " "
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.4)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f"( {mark} ) {text}")
    _font(r, size=10.5, bold=checked)
    return p


def table(doc: Document, headers: list[str], rows: list[list[str]], *, font_size: int = 9):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = ""
        _shade(cell, "E6E6E6")
        _cell_margins(cell)
        p = cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(h)
        _font(r, size=font_size, bold=True)
    for ri, row in enumerate(rows, start=1):
        is_total = any(
            str(c).upper().startswith(("SUBTOTAL", "TOTAL", "CUSTO GLOBAL")) for c in row
        )
        for ci, val in enumerate(row):
            cell = t.rows[ri].cells[ci]
            cell.text = ""
            if is_total:
                _shade(cell, "F5F5F5")
            _cell_margins(cell)
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            r = p.add_run(str(val))
            _font(r, size=font_size, bold=is_total)
    doc.add_paragraph()
    return t


def risk_card(
    doc: Document,
    titulo: str,
    prob: str,
    impacto: str,
    situacao: str,
    preventiva: str,
    resp_prev: str,
    contingencia: str,
    resp_cont: str,
):
    """Matriz de risco no formato oficial do Guia SARH/Pelotas (Decreto 6.799/2023)."""
    t = doc.add_table(rows=4, cols=2)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Linha 0 mesclada: título do risco
    c0 = t.rows[0].cells[0].merge(t.rows[0].cells[1])
    c0.text = ""
    _shade(c0, "E6E6E6")
    _cell_margins(c0)
    r = c0.paragraphs[0].add_run(f"Risco: {titulo}")
    _font(r, size=9.5, bold=True)

    # Linha 1: Probabilidade/Impacto | Situação Prevista
    cells_1 = [
        f"Probabilidade: {prob} | Impacto: {impacto}",
        f"Situação Prevista / Dano: {situacao}",
    ]
    for idx, txt in enumerate(cells_1):
        cell = t.rows[1].cells[idx]
        cell.text = ""
        _cell_margins(cell)
        cell.paragraphs[0].paragraph_format.space_after = Pt(0)
        _font(cell.paragraphs[0].add_run(txt), size=9)

    # Linha 2: Ação Preventiva | Responsável
    cells_2 = [f"Ação Preventiva: {preventiva}", f"Responsável: {resp_prev}"]
    for idx, txt in enumerate(cells_2):
        cell = t.rows[2].cells[idx]
        cell.text = ""
        _cell_margins(cell)
        cell.paragraphs[0].paragraph_format.space_after = Pt(0)
        _font(cell.paragraphs[0].add_run(txt), size=9)

    # Linha 3: Ação de Contingência | Responsável
    cells_3 = [f"Ação de Contingência: {contingencia}", f"Responsável: {resp_cont}"]
    for idx, txt in enumerate(cells_3):
        cell = t.rows[3].cells[idx]
        cell.text = ""
        _cell_margins(cell)
        cell.paragraphs[0].paragraph_format.space_after = Pt(0)
        _font(cell.paragraphs[0].add_run(txt), size=9)

    doc.add_paragraph()


def bloco_assinaturas(doc: Document, incluir_autorizacao_secretaria: bool = True):
    para(
        doc,
        "Certificamos que somos responsáveis pela elaboração do presente instrumento e que este "
        "encontra-se em consonância com a Lei Federal nº 14.133/2021, com o Decreto Municipal "
        "nº 6.799/2023 e com a Lei Municipal nº 6.621/2018.",
    )
    para(doc, "Pelotas/RS, 26 de novembro de 2026.", justify=False)

    t = doc.add_table(rows=1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    membros = [
        [
            "__________________________________",
            SOLICITANTE,
            "Departamento de Saúde Digital / SMS",
            "Matrícula: [informar]",
        ],
        [
            "__________________________________",
            "[Responsável Técnico / Planejamento]",
            "Diretoria de Atenção Primária / SMS",
            "Matrícula: [informar]",
        ],
    ]
    for i, lines in enumerate(membros):
        cell = t.rows[0].cells[i]
        cell.text = ""
        for j, ln in enumerate(lines):
            p = cell.paragraphs[0] if j == 0 else cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            _font(p.add_run(ln), size=9.5, bold=(j == 1))

    doc.add_paragraph()
    autoridades = [[DIRETORA[0], DIRETORA[1], "Ciente / De acordo"]]
    if incluir_autorizacao_secretaria:
        autoridades.append(
            [
                SECRETARIA[0],
                SECRETARIA[1],
                "Autorizo (ratificando acordo registrado no SIS nº 15912/2026 em 09/11/2026)",
            ]
        )
    for bloco in autoridades:
        p_line = doc.add_paragraph()
        p_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_line.paragraph_format.space_before = Pt(8)
        p_line.paragraph_format.space_after = Pt(1)
        _font(p_line.add_run("_" * 46), size=10)
        for idx, txt in enumerate(bloco):
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(0)
            _font(p.add_run(txt), size=9.5, bold=(idx == 0), italic=(idx == 2))


# ---------------------------------------------------------------------------
# 1) DFD — DOCUMENTO DE FORMALIZAÇÃO DA DEMANDA
# ---------------------------------------------------------------------------
def gerar_dfd(path: Path):
    doc = new_doc()
    title(
        doc,
        "DOCUMENTO DE FORMALIZAÇÃO DA DEMANDA – DFD nº [___]/2026",
        f"Expediente: Processo SIS nº {SIS_NUM} (Abertura em {SIS_ABERTURA}) — Projeto Tele-espirometria (CTS-HC/UFMG / MS)",
    )

    table(
        doc,
        ["Campo de Identificação", "Informação"],
        [
            ["Unidade Demandante", UNIDADE],
            ["Endereço da Unidade", ENDERECO_UNIDADE],
            ["Responsável pela Solicitação", SOLICITANTE],
            ["Processo Administrativo (SIS)", f"SIS nº {SIS_NUM} — Tipo: Diária / Pernoite / Passagem"],
            ["Servidora Designada", f"{SERVIDORA} — Cargo: {CARGO_SERVIDORA} — Matrícula: {MATRICULA_SERVIDORA}"],
            [
                "Dados Cadastrais e Bancários (LGPD)",
                f"CPF e conta bancária (Banrisul) devidamente registrados na folha inicial do Processo SIS nº {SIS_NUM}, preservados nesta peça em observância à Lei nº 13.709/2018 (LGPD).",
            ],
            [
                "Situação da Autorização Superior",
                "Acordo da Secretária Municipal de Saúde registrado no SIS em 09/11/2026 e reconfirmado em 10/11/2026; autos devolvidos pelo Setor de Compras em 11/11/2026 para juntada de DFD e ETP.",
            ],
        ],
        font_size=9,
    )

    h2(doc, "1. OBJETO")
    para(
        doc,
        f"Contratação de serviços de fornecimento de passagens (aérea no trecho Porto Alegre/RS ↔ Recife/PE e "
        f"rodoviária intermunicipal no trecho Pelotas/RS ↔ Porto Alegre/RS), cumulada com a concessão de diárias "
        f"civis estatutárias, visando viabilizar a participação da servidora {SERVIDORA}, {CARGO_SERVIDORA}, "
        f"Matrícula nº {MATRICULA_SERVIDORA}, no Treinamento Prático Presencial do Projeto Nacional de "
        f"Tele-espirometria (Projeto TeleRESP), promovido pelo {PROMOTOR}, a realizar-se no dia {EVENTO_DATA}, "
        f"no município de Recife/PE, com deslocamento de {SAIDA_DATA} ({SAIDA_HORA}) a {RETORNO_DATA} ({RETORNO_HORA}).",
    )
    para(
        doc,
        "Distinção de regimes jurídicos: (i) a aquisição das passagens aéreas e terrestres configura contratação de "
        "serviços comuns de transporte/agenciamento, regida pela Lei Federal nº 14.133/2021 e pelo Decreto Municipal "
        "nº 6.799/2023 (Elemento de Despesa 3.3.90.33); (ii) a concessão de diárias à servidora possui natureza "
        "indenizatória estatutária, regida pela Lei Municipal nº 6.621/2018 (Elemento de Despesa 3.3.90.14), não "
        "sujeita a procedimento licitatório, mas integrada a este DFD e ao respectivo ETP para assegurar planejamento "
        "unificado, demonstração do custo global da missão e adequação orçamentária.",
        bold_prefix="Nota técnica de enquadramento — ",
    )

    h2(doc, "1.1 TIPO DO OBJETO")
    checkbox(doc, "Material de Consumo")
    checkbox(doc, "Material Permanente / Equipamento")
    checkbox(doc, "Serviço continuado com dedicação exclusiva de mão de obra")
    checkbox(doc, "Serviço continuado sem dedicação exclusiva de mão de obra")
    checkbox(doc, "Serviço não continuado", checked=True)
    para(
        doc,
        "Observação: enquadramento aplicável à contratação pontual de passagens aéreas e rodoviárias; as diárias "
        "civis seguem o regime indenizatório próprio da Lei Municipal nº 6.621/2018.",
        size=9.5,
    )

    h2(doc, "2. FORMA DE CONTRATAÇÃO SUGERIDA")
    checkbox(doc, "Pregão Eletrônico*  |  Sistema de Registro de Preços: (   ) SIM   (   ) NÃO")
    checkbox(doc, "Concorrência")
    checkbox(doc, "Dispensa de licitação", checked=True)
    checkbox(doc, "Inexigibilidade")
    checkbox(doc, "Adesão à ata de registro de preços de outro(s) Órgão(s)")
    para(
        doc,
        f"Para o fornecimento das passagens (estimativa referencial de {brl(SUBTOTAL_PASSAGENS)}), sugere-se "
        f"contratação direta por Dispensa de Licitação em razão do valor, com fundamento no art. 75, inciso II, da "
        f"Lei Federal nº 14.133/2021 (teto atualizado para {brl(LIMITE_DISPENSA_2026)} em 2026 pelo Decreto Federal "
        f"nº 12.807/2025) e no Decreto Municipal nº 6.799/2023, mediante pesquisa de preços atualizada (art. 23) e "
        f"verificação de não fracionamento da despesa no exercício (art. 75, § 1º), observando-se que o Município de "
        f"Pelotas não dispõe atualmente de contrato ou ata de registro de preços vigente para agenciamento de "
        f"passagens aéreas (conforme precedente municipal de 2026 — Dispensa nº 228292/2026). Para o trecho terrestre "
        f"Pelotas ↔ Porto Alegre, admite-se a aquisição junto à agência contratada, a emissão direta junto à "
        f"concessionária regulada pelo DAER/RS ou a utilização de veículo oficial da SMS, caso disponível na data. "
        f"As diárias civis serão concedidas diretamente à servidora nos termos da Lei Municipal nº 6.621/2018.",
    )

    h2(doc, "3. JUSTIFICATIVA DA NECESSIDADE DA CONTRATAÇÃO DE SERVIÇO OU AQUISIÇÃO")
    para(
        doc,
        "As doenças respiratórias crônicas, em especial a asma e a doença pulmonar obstrutiva crônica (DPOC), "
        "figuram entre as principais causas de morbimortalidade e de sobrecarga dos serviços de urgência e atenção "
        "especializada no SUS. No âmbito do Novo Programa de Aceleração do Crescimento (Novo PAC Saúde — Combo de "
        "Equipamentos para UBS), o Ministério da Saúde promoveu a modernização tecnológica da Atenção Primária à "
        "Saúde (APS) com a distribuição de espirômetros digitais com interlocução com telessaúde/telediagnóstico. "
        "Paralelamente, a Portaria GM/MS nº 3.691/2024 instituiu a Ação Estratégica SUS Digital – Telessaúde "
        "(prevendo a modalidade de telediagnóstico em seu art. 448, IV), e a Portaria GM/MS nº 7.495/2025 reforçou "
        "o telediagnóstico na APS para ampliar a resolubilidade clínica e qualificar o encaminhamento à atenção "
        "especializada.",
    )
    para(
        doc,
        f"Nesse contexto, o Município de Pelotas/RS formalizou adesão ao Projeto Nacional de Tele-espirometria "
        f"(Projeto TeleRESP), desenvolvido pelo {PROMOTOR} (coordenado pelo {COORDENADOR_PROJETO}), mediante envio "
        f"de Termo de Compromisso assinado pela Secretaria Municipal de Saúde e confirmação oficial de adesão em "
        f"09/11/2026. O projeto permitirá que exames de espirometria sejam realizados diretamente na rede de APS de "
        f"Pelotas, com transmissão eletrônica e emissão gratuita de laudos a distância por pneumologistas do HC-UFMG, "
        f"além de teleconsultoria especializada para as equipes municipais.",
    )
    para(
        doc,
        f"Para habilitar o Município a operar o serviço, a capacitação exige duas etapas obrigatórias e "
        f"complementares: (1) módulo teórico em EaD na plataforma Moodle do CTS-HC/UFMG (de 11/11/2026 até 08/12/2026, "
        f"uma semana antes da prática); e (2) módulo prático presencial, agendado para o dia {EVENTO_DATA}, das "
        f"{EVENTO_HORARIO}, na sede da SES-PE em Recife/PE, com obrigatoriedade de levar o espirômetro digital do "
        f"Novo PAC, cabo Micro-USB e turbina para calibração e prática supervisionada. Não há taxa de inscrição "
        f"cobrada pelo projeto, cabendo ao Município exclusivamente o custeio do deslocamento e das diárias da "
        f"profissional indicada.",
    )
    para(
        doc,
        f"A servidora {SERVIDORA} ({CARGO_SERVIDORA}, Matrícula {MATRICULA_SERVIDORA}) atuará como profissional "
        f"multiplicadora no Município de Pelotas, replicando o conhecimento técnico na rede municipal de saúde e "
        f"garantindo a efetiva utilização dos espirômetros recebidos via Novo PAC, evitando ociosidade de bem público "
        f"federal e reduzindo filas de espera por exames de função pulmonar. A contratação das passagens e a "
        f"concessão das diárias são estritamente interdependentes, pois sem a logística completa (deslocamento "
        f"Pelotas ↔ Porto Alegre, passagem aérea Porto Alegre ↔ Recife e manutenção/hospedagem no destino) "
        f"inviabiliza-se a presença da servidora no treinamento presencial obrigatório.",
    )

    h2(doc, "4. QUANTIDADE E DESCRIÇÃO DO MATERIAL OU SERVIÇO")
    para(doc, "4.1 Parcela sujeita à Lei Federal nº 14.133/2021 — Passagens (Elemento 3.3.90.33):", bold_prefix="")
    table(
        doc,
        ["Item", "Especificação (CATSER ref.)", "Unid.", "Qtd.", "Valor Unit. Ref.", "Valor Total Ref."],
        [
            [
                "01",
                "Passagem aérea nacional, ida e volta, classe econômica, trecho Porto Alegre/RS (POA) ↔ Recife/PE (REC), ida em 14/12/2026 e retorno em 16/12/2026, com taxas aeroportuárias inclusas e franquia mínima de bagagem de mão (10 kg) para transporte do espirômetro PAC. (CATSER ref.: 3719 / Agenciamento de viagens)",
                "UN (ida e volta)",
                "1",
                brl(VAL_AEREA),
                brl(VAL_AEREA),
            ],
            [
                "02",
                "Passagem rodoviária intermunicipal (ônibus executivo/convencional — linha regulada pelo DAER/RS), trecho Pelotas/RS ↔ Porto Alegre/RS, ida em 14/12/2026 (saída ~06h00) e retorno em 16/12/2026 (chegada ~18h00). (CATSER ref.: 25089 / Transporte rodoviário)",
                "Trecho",
                "2",
                brl(VAL_RODOV_UNIT),
                brl(VAL_RODOV_TOTAL),
            ],
            ["", "SUBTOTAL PASSAGENS (Lei nº 14.133/2021)", "", "", "", brl(SUBTOTAL_PASSAGENS)],
        ],
        font_size=8.5,
    )

    para(
        doc,
        "4.2 Parcela indenizatória estatutária — Diárias Civis (Lei Municipal nº 6.621/2018 — Elemento 3.3.90.14):",
    )
    table(
        doc,
        ["Modalidade (Lei 6.621/2018)", "Enquadramento", "Período Coberto", "Qtd.", "URMs", "Valor Ref.*"],
        [
            [
                "Diária Integral (24h)",
                "Nível 02 — Demais agentes públicos — Outros Estados (sem transporte oficial)",
                "14/12/2026 (06h) a 16/12/2026 (06h) — 48h",
                "2",
                "8 URMs (4 cada)",
                brl(8 * URM_REF),
            ],
            [
                "Diária Estada (≤ 12h)",
                "Nível 02 — Demais agentes públicos — Outros Estados (sem transporte oficial)",
                "16/12/2026 (06h) a 16/12/2026 (18h) — 12h",
                "1",
                "2 URMs",
                brl(2 * URM_REF),
            ],
            [
                "SUBTOTAL DIÁRIAS (Indenizatório)",
                "Total de 60h de afastamento da sede (art. 2º, § 1º, e art. 3º)",
                "14/12 a 16/12/2026",
                "2 integrais + 1 estada",
                "10 URMs não trib.",
                brl(SUBTOTAL_DIARIAS_REF),
            ],
            [
                "CUSTO GLOBAL PRELIMINAR DA MISSÃO",
                "Passagens (3.3.90.33) + Diárias (3.3.90.14)",
                "—",
                "—",
                "—",
                brl(TOTAL_GLOBAL_REF),
            ],
        ],
        font_size=8.5,
    )
    para(
        doc,
        f"*Notas de referência de valores: (a) o valor da passagem aérea ({brl(VAL_AEREA)}) reflete a média de "
        f"mercado para dezembro/2026 (alta temporada) e deverá ser validado por pesquisa formal de preços (art. 23, "
        f"Lei nº 14.133/2021) imediatamente antes da emissão; (b) o valor rodoviário ({brl(VAL_RODOV_UNIT)}/trecho) "
        f"corresponde à tarifa executiva 2026 regulada pelo DAER/RS (serviço convencional: R$ 111,96/trecho); (c) o "
        f"valor de diárias ({brl(SUBTOTAL_DIARIAS_REF)}) utiliza, a título meramente referencial, a última URM não "
        f"tributária localizada em publicação oficial (R$ 111,27 — Decreto nº 7.021/2025), devendo o setor financeiro/"
        f"Contabilidade aplicar o valor oficial vigente da Tabela URM Não Tributária de novembro/dezembro de 2026 "
        f"publicada pela SEFAZ/Pelotas no ato do empenho.",
        size=9,
    )

    h2(doc, "5. DOTAÇÃO ORÇAMENTÁRIA")
    table(
        doc,
        ["Campo", "Passagens (Lei 14.133/2021)", "Diárias (Lei Municipal nº 6.621/2018)"],
        [
            ["Órgão / Unidade", "Secretaria Municipal de Saúde – SMS", "Secretaria Municipal de Saúde – SMS"],
            [
                "Projeto / Atividade",
                "[Confirmar CCGOF — ex.: 10.301.0105.2167.00 APS / Saúde Digital]",
                "[Confirmar CCGOF — ex.: 10.301.0105.2167.00 APS / Saúde Digital]",
            ],
            ["Fonte de Recurso", "[Confirmar CCGOF — ex.: 1600 / SUS Digital / Próprio]", "[Confirmar CCGOF — ex.: 1600 / Próprio]"],
            ["Elemento de Despesa", "3.3.90.33 — Passagens e Despesas com Locomoção", "3.3.90.14 — Diárias – Pessoal Civil"],
            ["Valor Estimado", f"{brl(SUBTOTAL_PASSAGENS)} (a confirmar na cotação)", f"10 URMs não trib. (ref. mín. {brl(SUBTOTAL_DIARIAS_REF)}*)"],
            ["Forma de Execução", "Nota de Empenho Ordinário", "Nota de Empenho Ordinário (crédito em conta)"],
        ],
        font_size=8.5,
    )

    h2(doc, "6. LOCAL DE PRESTAÇÃO DOS SERVIÇOS OU PARA ENTREGA DO OBJETO")
    bullet(
        doc,
        f"Envio em formato digital (e-ticket/localizador e comprovantes) ao Departamento de Saúde Digital "
        f"({ENDERECO_UNIDADE}) e ao e-mail institucional da servidora designada.",
        bold_prefix="Entrega dos bilhetes: ",
    )
    bullet(
        doc,
        "Trecho terrestre: Terminal Rodoviário de Pelotas/RS ↔ Estação Rodoviária / Aeroporto Internacional "
        "Salgado Filho (Porto Alegre/RS). Trecho aéreo: Aeroporto Internacional Salgado Filho (POA) ↔ Aeroporto "
        "Internacional do Recife/Guararapes (REC).",
        bold_prefix="Execução do transporte: ",
    )
    bullet(
        doc,
        f"{EVENTO_LOCAL}.",
        bold_prefix="Local da capacitação presencial: ",
    )

    h2(doc, "7. PREVISÃO DE DATA EM QUE DEVE SER INICIADA A PRESTAÇÃO DOS SERVIÇOS")
    bullet(
        doc,
        "Imediata após autorização e empenho, recomendada emissão dos bilhetes até o início de dezembro de 2026 "
        "para evitar majoração tarifária típica da alta temporada aérea.",
        bold_prefix="Emissão dos bilhetes e crédito das diárias: ",
    )
    bullet(
        doc,
        f"Saída de Pelotas/RS em {SAIDA_DATA} às {SAIDA_HORA}; treinamento presencial em Recife/PE em "
        f"{EVENTO_DATA} ({EVENTO_HORARIO}); retorno previsto a Pelotas/RS em {RETORNO_DATA} às {RETORNO_HORA}.",
        bold_prefix="Período de deslocamento e evento: ",
    )
    bullet(
        doc,
        "Até 5 (cinco) dias úteis após o retorno (até 23/12/2026), mediante entrega de relatório resumido de "
        "atividades, comprovantes de despesas e certificado/declaração de participação, nos termos do art. 8º da "
        "Lei Municipal nº 6.621/2018.",
        bold_prefix="Prestação de contas das diárias: ",
    )

    h2(doc, "8. INDICAÇÃO DOS MEMBROS DA EQUIPE DE PLANEJAMENTO DA CONTRATAÇÃO")
    bloco_assinaturas(doc, incluir_autorizacao_secretaria=True)

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    print(f"OK {path}")


# ---------------------------------------------------------------------------
# 2) ETP — ESTUDO TÉCNICO PRELIMINAR
# ---------------------------------------------------------------------------
def gerar_etp(path: Path):
    doc = new_doc()
    title(
        doc,
        "ESTUDO TÉCNICO PRELIMINAR – ETP",
        f"Processo SIS nº {SIS_NUM} — Passagens e Diárias para Treinamento Presencial em Tele-espirometria (Recife/PE)",
    )

    # 1
    h2(
        doc,
        "1. IDENTIFICAÇÃO DO PROCESSO",
        "Art. 18, § 1º, I, da Lei Federal nº 14.133/2021 e art. 26, I, do Decreto Municipal nº 6.799/2023.",
    )
    table(
        doc,
        ["Campo", "Detalhamento"],
        [
            ["Unidade Responsável", UNIDADE],
            ["Endereço / Contato", f"{ENDERECO_UNIDADE} — Responsável: {SOLICITANTE}"],
            ["Número do Processo", f"Processo SIS nº {SIS_NUM} (abertura em {SIS_ABERTURA}) — DFD nº [___]/2026"],
            [
                "Objeto da Contratação",
                f"Contratação de passagens aéreas (POA ↔ REC) e rodoviárias (Pelotas ↔ POA), associada à "
                f"concessão de diárias civis (Lei Municipal nº 6.621/2018), para viabilizar a participação da "
                f"servidora {SERVIDORA} ({CARGO_SERVIDORA}, Matrícula {MATRICULA_SERVIDORA}) no Treinamento "
                f"Prático Presencial do Projeto Tele-espirometria (CTS-HC/UFMG / MS), em Recife/PE, em {EVENTO_DATA}.",
            ],
            [
                "Histórico Processual (SIS)",
                "05/11/2026: abertura pelo Depto. Saúde Digital; 06/11/2026: encaminhamento pela Diretoria de "
                "Atenção Primária; 09/11/2026: Secretária Municipal de Saúde manifesta acordo (reconfirmado em "
                "10/11/2026); 11/11/2026: DAF remete ao Setor de Compras, que devolve os autos para instrução "
                "com DFD e ETP.",
            ],
        ],
        font_size=9,
    )

    # 2
    h2(
        doc,
        "2. DESCRIÇÃO DA NECESSIDADE",
        "Art. 18, § 1º, I, da Lei Federal nº 14.133/2021 e art. 26, I, do Decreto Municipal nº 6.799/2023.",
    )
    para(
        doc,
        "A Secretaria Municipal de Saúde de Pelotas/RS enfrenta demanda expressiva relacionada ao diagnóstico "
        "oportuno e ao acompanhamento longitudinal das doenças respiratórias crônicas na Atenção Primária à Saúde "
        "(APS), especialmente asma e doença pulmonar obstrutiva crônica (DPOC). O acesso restrito ao exame de "
        "espirometria na atenção especializada gera subdiagnóstico, filas de espera, agravamento clínico e "
        "internações evitáveis.",
    )
    para(
        doc,
        "Para superar esse gargalo, o Governo Federal estruturou, no âmbito do Novo PAC Saúde (Combo de "
        "Equipamentos para UBS — aporte nacional de cerca de R$ 1,6 bilhão para modernização tecnológica de "
        "10 mil UBS em mais de 5 mil municípios), o fornecimento de espirômetros digitais com interlocução em "
        "telessaúde/telediagnóstico. Essa entrega articula-se à Ação Estratégica SUS Digital – Telessaúde "
        "(Portaria GM/MS nº 3.691/2024, art. 448, IV) e ao Componente SUS Digital do Programa Agora Tem "
        "Especialistas (Portaria GM/MS nº 7.495/2025, art. 12), que fomentam o telediagnóstico na APS para "
        "aumentar a resolubilidade clínica e qualificar os encaminhamentos.",
    )
    para(
        doc,
        f"Em consonância com essas diretrizes, o Município de Pelotas/RS aderiu ao Projeto Tele-espirometria "
        f"(Projeto TeleRESP), executado pelo {PROMOTOR} sob coordenação do {COORDENADOR_PROJETO}. O Termo de "
        f"Compromisso foi enviado pela Secretária Municipal de Saúde e a adesão de Pelotas foi formalmente "
        f"confirmada pelo CTS-HC/UFMG em 09/11/2026. O projeto garante aos municípios aderentes plataforma digital "
        f"de envio de exames, emissão gratuita de laudos por pneumologistas do HC-UFMG, teleconsultoria "
        f"especializada e capacitação técnica em duas etapas: módulo teórico EaD (11/11/2026 a 08/12/2026, no "
        f"Moodle www.telessaude.hc.ufmg.br/eadct) e módulo prático presencial obrigatório em {EVENTO_DATA}, "
        f"das {EVENTO_HORARIO}, no endereço {EVENTO_LOCAL}.",
    )
    para(
        doc,
        f"Sem a participação presencial de ao menos 01 (uma) profissional municipal habilitada — levando o "
        f"espirômetro digital recebido do Novo PAC, cabo Micro-USB e turbina —, o Município não conclui a "
        f"certificação técnica exigida para ativação do serviço de tele-espirometria, mantendo subutilizados os "
        f"equipamentos públicos federais recebidos e privando a população pelotense do diagnóstico pulmonar na APS. "
        f"A indicação da enfermeira {SERVIDORA} (Matrícula {MATRICULA_SERVIDORA}) com atribuição expressa de "
        f"multiplicação local assegura que uma única viagem capacite a referência técnica que replicará o método "
        f"às demais equipes da rede municipal.",
    )

    # 3
    h2(
        doc,
        "3. DA CONTRATAÇÃO E DO PLANEJAMENTO",
        "Art. 18, § 1º, II, da Lei Federal nº 14.133/2021 e art. 26, II, do Decreto Municipal nº 6.799/2023.",
    )
    para(
        doc,
        "Sob o prisma jurídico-administrativo, a missão institucional envolve duas despesas complementares com "
        "regimes distintos que precisam ser claramente segregados para correta instrução processual e contábil:",
    )
    bullet(
        doc,
        "Fornecimento de passagem aérea (POA ↔ REC) e passagens rodoviárias (Pelotas ↔ POA): serviços comuns de "
        "natureza não continuada, regidos pela Lei Federal nº 14.133/2021 e pelo Decreto Municipal nº 6.799/2023 "
        "(Elemento 3.3.90.33), enquadráveis em Dispensa de Licitação em razão do valor (art. 75, inciso II, teto "
        f"de {brl(LIMITE_DISPENSA_2026)} em 2026 pelo Decreto Federal nº 12.807/2025).",
        bold_prefix="Parcela contratual (Passagens): ",
    )
    bullet(
        doc,
        "Indenização das despesas de alimentação, hospedagem e deslocamento urbano em Recife/PE: verba "
        "indenizatória de pessoal regida pela Lei Municipal nº 6.621, de 23 de agosto de 2018 (Elemento 3.3.90.14), "
        "não configurando contratação de terceiros sob a Lei nº 14.133/2021, mas incorporada a este ETP para "
        "visibilidade integral do custo da viagem e controle de economicidade.",
        bold_prefix="Parcela estatutária (Diárias): ",
    )
    para(
        doc,
        "Alinhamento ao planejamento institucional e ao PCA: a ação encontra-se alinhada às diretrizes de "
        "fortalecimento da Atenção Primária e transformação digital do SUS previstas no Plano Plurianual (PPA "
        "2026–2029, instituído pela Lei Municipal nº 7.472/2025), na LDO, na LOA 2026, no Plano Municipal de Saúde "
        "e no Plano de Ação de Transformação para a Saúde Digital (PA Saúde Digital). Caso a demanda específica de "
        "passagens para Recife/PE não conste discriminada no Plano de Contratações Anual (PCA) 2026 — por decorrer "
        "de calendário definido externamente pelo Ministério da Saúde e pelo CTS-HC/UFMG, com adesão confirmada "
        "em 09/11/2026 —, justifica-se sua inclusão/execução superveniente com base no interesse público relevante "
        "e na ausência de previsibilidade da data e sede do treinamento no exercício anterior (conforme orientação "
        "do Guia de Preenchimento do ETP da SARH/Pelotas).",
    )

    # 4
    h2(
        doc,
        "4. REQUISITOS DA CONTRATAÇÃO",
        "Art. 18, § 1º, III, da Lei Federal nº 14.133/2021 e art. 26, III, do Decreto Municipal nº 6.799/2023.",
    )
    para(doc, "4.1 Requisitos técnicos das passagens aéreas (Item 01):")
    bullet(
        doc,
        "Trecho Porto Alegre/RS (Aeroporto Internacional Salgado Filho – POA) ↔ Recife/PE (Aeroporto Internacional "
        "do Recife/Guararapes – REC), em classe econômica, tarifa mais vantajosa disponível na data de emissão.",
    )
    bullet(
        doc,
        "Horários compatíveis com a logística completa: voo de ida em 14/12/2026 com embarque após a chegada do "
        "transporte proveniente de Pelotas (previsão de chegada em POA por volta de 09h30–09h45, recomendando-se "
        "margem mínima de segurança de 2h30 para traslado rodoviária–aeroporto e check-in) e desembarque em Recife "
        "no dia 14/12/2026; voo de retorno em 16/12/2026 com desembarque em POA em horário que permita conexão "
        "rodoviária segura para chegada a Pelotas por volta das 18h00.",
    )
    bullet(
        doc,
        "Franquia de bagagem: obrigatoriedade de bagagem de mão (mínimo 10 kg na cabine, conforme Resolução ANAC "
        "nº 400/2016) e item pessoal, para transporte protegido em cabine do espirômetro digital do Novo PAC, "
        "cabo Micro-USB e turbina, evitando riscos de extravio ou avaria em porão.",
    )
    bullet(
        doc,
        "Emissão de bilhete eletrônico (e-ticket) com código localizador válido, discriminação transparente de "
        "tarifa, taxas aeroportuárias e eventual taxa de agenciamento, além de suporte para eventual remarcação, "
        "alteração ou cancelamento em caso fortuito ou de força maior.",
    )
    para(doc, "4.2 Requisitos técnicos das passagens rodoviárias (Item 02):")
    bullet(
        doc,
        "Trecho Pelotas/RS ↔ Porto Alegre/RS (ida em 14/12/2026, saída prevista às 06h00; retorno em 16/12/2026, "
        "chegada prevista às 18h00), em empresa permissionária/concessionária regularmente autorizada pelo DAER/RS "
        "(modalidade executiva ou convencional com horário compatível com a conexão aérea).",
    )
    para(doc, "4.3 Requisitos de habilitação do fornecedor:")
    bullet(
        doc,
        "Habilitação jurídica, regularidade fiscal federal, estadual, municipal, FGTS e trabalhista (CNDT), "
        "observada a proporcionalidade do art. 70, inciso III, da Lei nº 14.133/2021 para contratações de entrega "
        "imediata e valor inferior a 1/4 do limite de dispensa, bem como registro regular no CADASTUR quando "
        "contratada agência de viagens.",
    )
    para(doc, "4.4 Requisitos institucionais e estatutários da missão (Lei Municipal nº 6.621/2018):")
    bullet(
        doc,
        "Conclusão obrigatória pela servidora do curso EaD no Moodle do CTS-HC/UFMG até 08/12/2026; confirmação "
        "prévia de presença por e-mail ao organizador; porte de documento oficial de identidade (RG) e do "
        "espirômetro PAC com acessórios; e entrega da prestação de contas em até 5 (cinco) dias úteis após o "
        "retorno (art. 8º da Lei nº 6.621/2018).",
    )

    # 5
    h2(
        doc,
        "5. ESTIMATIVA DAS QUANTIDADES",
        "Art. 18, § 1º, IV, da Lei Federal nº 14.133/2021 e art. 26, IV, do Decreto Municipal nº 6.799/2023.",
    )
    para(
        doc,
        "Os quantitativos decorrem diretamente da convocação oficial do CTS-HC/UFMG (que requer a participação de "
        "01 profissional por município, em treinamento presencial de turno integral no dia 15/12/2026, das 08h00 "
        "às 17h45, com chegada às 07h30) combinada com a distância geográfica entre Pelotas/RS e Recife/PE (~3.200 km "
        "por via aérea a partir de Porto Alegre, sem voo direto de Pelotas).",
    )
    para(doc, "Memória de cálculo do período de afastamento e das diárias (Lei Municipal nº 6.621/2018):")
    table(
        doc,
        ["Etapa / Data", "Horário Previsto", "Atividade / Justificativa Logística", "Enquadramento (Lei 6.621/2018)"],
        [
            [
                "Dia 1 — 14/12/2026",
                "06h00 às 23h59",
                "Deslocamento rodoviário Pelotas → POA (~3h30) + conexão aeroportuária + voo POA → Recife (~4h05 direto ou ~6h com escala) + 1º pernoite em Recife (indispensável para credenciamento às 07h30 do dia 15/12).",
                "Início da 1ª Diária Integral (14/12 06h → 15/12 06h = 24h) — 4 URMs",
            ],
            [
                "Dia 2 — 15/12/2026",
                "07h30 às 17h45",
                "Treinamento prático presencial na SES-PE (Recife/PE). Término às 17h45 inviabiliza retorno a Pelotas/RS no mesmo dia (art. 3º da Lei nº 6.621/2018), exigindo 2º pernoite em Recife.",
                "2ª Diária Integral (15/12 06h → 16/12 06h = 24h) — 4 URMs",
            ],
            [
                "Dia 3 — 16/12/2026",
                "06h00 às 18h00",
                "Voo de retorno Recife → POA + deslocamento rodoviário POA → Pelotas, com chegada prevista às 18h00 (período residual de exatas 12h, sem novo pernoite).",
                "1 Diária Estada (≤ 12h, município não limítrofe — art. 2º, § 1º) — 2 URMs",
            ],
            [
                "TOTAL",
                "60 horas de afastamento",
                "1 servidora | 1 passagem aérea (ida/volta) | 2 trechos rodoviários | 2 pernoites",
                "2 Diárias Integrais + 1 Estada = 10 URMs não tributárias",
            ],
        ],
        font_size=8.5,
    )
    para(
        doc,
        "Observação legal (art. 2º, §§ 3º e 4º, da Lei nº 6.621/2018): caso o retorno efetivo ocorra em prazo menor "
        "que o previsto, a servidora restituirá os valores recebidos em excesso; caso atrasos aéreos/rodoviários "
        "comprovados façam o período exceder as 60 horas estimadas, caberá análise de ressarcimento mediante "
        "justificativa fundamentada e autorização expressa da autoridade competente.",
        size=9.5,
    )

    # 6
    h2(
        doc,
        "6. LEVANTAMENTO DE MERCADO",
        "Art. 18, § 1º, V, da Lei Federal nº 14.133/2021 e art. 26, V, do Decreto Municipal nº 6.799/2023.",
    )
    para(
        doc,
        "Foram analisadas as alternativas técnicas, logísticas e contratuais disponíveis no mercado e na "
        "Administração Municipal para atender à demanda:",
    )
    table(
        doc,
        ["Dimensão Analisada", "Alternativa Avaliada", "Análise Técnica e Econômica", "Conclusão"],
        [
            [
                "Formato da capacitação",
                "A1 — Participação exclusivamente remota (EaD)",
                "O CTS-HC/UFMG exige módulo prático presencial obrigatório com o espirômetro do Novo PAC para habilitação técnica. Não existe substituto 100% remoto.",
                "Inviável",
            ],
            [
                "Aeroporto de embarque",
                "B1 — Voo partindo de Pelotas (PET) ↔ Recife (REC)",
                "Malha aérea reduzida em Pelotas, exigindo múltiplas conexões (PET → POA/GRU/VCP → REC), tempo elevado de viagem, maior risco de perda de conexão e tarifas geralmente superiores.",
                "Desvantajosa (avaliar apenas se cotação específica empatar em custo total)",
            ],
            [
                "Aeroporto de embarque",
                "B2 — Transporte terrestre Pelotas ↔ POA + voo POA ↔ REC",
                "Maior oferta de voos diários (Azul, LATAM, GOL), inclusive voo direto Azul POA–REC (~4h05), melhor previsibilidade de horários e menor custo total combinado.",
                "Solução recomendada",
            ],
            [
                "Trecho terrestre Pelotas ↔ POA",
                "C1 — Passagem rodoviária (linha DAER/RS) vs. Veículo oficial da SMS",
                "Contrato nº 098/2025 da SMS com a Expresso Embaixador destina-se exclusivamente ao transporte de pacientes (TFD), vedado desvio de finalidade. A compra de 2 bilhetes rodoviários (R$ 223,92 a R$ 262,70) é econômica; havendo veículo oficial da SMS já escalado para POA nas datas sem custo adicional, seu uso reduzirá a despesa.",
                "Passagem rodoviária (com checagem prévia de veículo oficial)",
            ],
            [
                "Instrumento contratual (passagens)",
                "D1 — Contrato municipal vigente ou Adesão a ARP",
                "Consulta ao Portal da Transparência do Município de Pelotas (2025–2026) não identificou contrato ou ARP vigente da Prefeitura para agenciamento de passagens aéreas. Processo de adesão a ARP externa demandaria prazo incompatível com a data da viagem (14/12/2026).",
                "Indisponível / Intempestiva",
            ],
            [
                "Instrumento contratual (passagens)",
                "D2 — Dispensa de Licitação por valor (art. 75, II, Lei 14.133/2021)",
                "Valor estimado (~R$ 2.612,70) corresponde a menos de 4% do limite legal de 2026 (R$ 65.492,11). Alinhada à prática vigente do Município (ex.: Dispensa nº 228292/2026 – SMHRF) e compatível com a urgência da alta temporada aérea.",
                "Solução escolhida",
            ],
        ],
        font_size=8.5,
    )

    # 7
    h2(
        doc,
        "7. ESTIMATIVA DO VALOR DA CONTRATAÇÃO",
        "Art. 18, § 1º, VI, da Lei Federal nº 14.133/2021 e art. 26, VI, do Decreto Municipal nº 6.799/2023.",
    )
    para(
        doc,
        "A estimativa preliminar de preços foi construída a partir de parâmetros públicos verificáveis de 2026, "
        "conforme art. 23 da Lei nº 14.133/2021: (i) para o trecho aéreo POA ↔ REC em dezembro/2026 (alta temporada), "
        "levantamento em consolidadores especializados (Mundi/KAYAK, dados atualizados em novembro/2026) aponta "
        "média entre R$ 2.025,00 (LATAM) e R$ 2.393,00 (média geral), adotando-se R$ 2.350,00 como estimativa "
        "prudencial de referência; (ii) para o trecho rodoviário Pelotas ↔ POA, a tarifa oficial regulada pelo "
        "DAER/RS (reajuste de 8,18% vigente desde 08/04/2026 — Expresso Embaixador) é de R$ 111,96 (convencional) "
        "e R$ 131,35 (executivo), adotando-se a modalidade executiva (R$ 262,70 ida e volta) como teto referencial; "
        "(iii) para as diárias, aplicam-se 10 URMs não tributárias (Lei Municipal nº 6.621/2018).",
    )
    table(
        doc,
        ["Componente", "Regime / Elemento", "Base de Cálculo", "Valor Estimado Ref."],
        [
            [
                "Item 01 — Passagem aérea POA ↔ REC (ida 14/12 e volta 16/12)",
                "Lei 14.133/2021 | 3.3.90.33",
                "1 bilhete ida e volta (classe econômica + taxas + bagagem mão 10 kg)",
                brl(VAL_AEREA),
            ],
            [
                "Item 02 — Passagem rodoviária Pelotas ↔ POA (ida 14/12 e volta 16/12)",
                "Lei 14.133/2021 | 3.3.90.33",
                "2 trechos × R$ 131,35 (tarifa executiva DAER/RS; convencional R$ 111,96)",
                brl(VAL_RODOV_TOTAL),
            ],
            [
                "SUBTOTAL CONTRATAÇÃO DE PASSAGENS (Art. 75, II)",
                "3,99% do teto anual de R$ 65.492,11",
                "Validar com mínimo de 3 cotações formais (art. 23)",
                brl(SUBTOTAL_PASSAGENS),
            ],
            [
                "Diárias Civis — 2 Integrais (8 URMs) + 1 Estada (2 URMs)",
                "Lei Mun. 6.621/2018 | 3.3.90.14",
                "10 URMs não trib. × R$ 111,27* (atualizar p/ tabela SEFAZ dez/2026)",
                brl(SUBTOTAL_DIARIAS_REF),
            ],
            [
                "CUSTO GLOBAL PRELIMINAR ESTIMADO DA MISSÃO",
                "Passagens + Diárias",
                "Sujeito a ajuste pela cotação final aérea e URM dez/2026",
                brl(TOTAL_GLOBAL_REF),
            ],
        ],
        font_size=8.5,
    )
    para(
        doc,
        "Ressalva operacional: em razão da volatilidade dinâmica das tarifas aéreas na primeira quinzena de "
        "dezembro, o Setor de Compras deverá realizar pesquisa formal contemporânea (cotações em agências de "
        "viagens e/ou sites oficiais das companhias aéreas, instruindo o mapa de preços do art. 23) e proceder à "
        "emissão imediata após o empenho, verificando previamente o somatório anual do ramo de atividade na "
        "unidade gestora (art. 75, § 1º, da Lei nº 14.133/2021). Da mesma forma, o setor financeiro aplicará o "
        "valor vigente da URM não tributária de novembro/dezembro de 2026 publicado pela SEFAZ/Pelotas.",
        size=9.5,
    )

    # 8
    h2(
        doc,
        "8. DESCRIÇÃO DA SOLUÇÃO COMO UM TODO",
        "Art. 18, § 1º, VII, da Lei Federal nº 14.133/2021 e art. 26, VII, do Decreto Municipal nº 6.799/2023.",
    )
    para(
        doc,
        "A solução integrada compreende todo o ciclo operacional necessário ao pleno êxito da capacitação e à "
        "implantação da Tele-espirometria na APS de Pelotas:",
    )
    bullet(
        doc,
        "Conclusão do módulo teórico EaD pela servidora (11/11 a 08/12/2026), confirmação de presença ao "
        "CTS-HC/UFMG, emissão de termo de cautela patrimonial para transporte temporário do espirômetro digital "
        "do Novo PAC (com cabo Micro-USB e turbina), pesquisa formal de preços e emissão dos empenhos (3.3.90.33 "
        "e 3.3.90.14).",
        bold_prefix="Fase 1 — Preparação e contratação: ",
    )
    bullet(
        doc,
        "Emissão dos bilhetes digitais (rodoviário e aéreo) com compatibilidade de conexões e depósito tempestivo "
        "das diárias na conta bancária da servidora (Banrisul, constante no Processo SIS nº 15912/2026) antes da "
        "data de embarque.",
        bold_prefix="Fase 2 — Logística e execução (14 a 16/12/2026): ",
    )
    bullet(
        doc,
        "Participação presencial integral no treinamento prático na SES-PE (15/12/2026), retorno a Pelotas e "
        "entrega da prestação de contas no prazo máximo de 5 (cinco) dias úteis (relatório de atividades, "
        "comprovantes de despesas e certificado/declaração de participação — art. 8º, Lei nº 6.621/2018).",
        bold_prefix="Fase 3 — Prestação de contas: ",
    )
    bullet(
        doc,
        "Atuação da servidora como multiplicadora junto às equipes da APS e do Departamento de Saúde Digital, "
        "viabilizando a realização de espirometrias nas UBS de referência com laudos remotos pelo CTS-HC/UFMG.",
        bold_prefix="Fase 4 — Multiplicação e impacto assistencial: ",
    )

    # 9
    h2(
        doc,
        "9. JUSTIFICATIVA PARA PARCELAMENTO",
        "Art. 18, § 1º, VIII, da Lei Federal nº 14.133/2021 e art. 26, VIII, do Decreto Municipal nº 6.799/2023.",
    )
    para(
        doc,
        "Nos termos do art. 40, inciso V, alínea 'b', e § 2º da Lei nº 14.133/2021, avaliou-se a divisibilidade do "
        "objeto em duas camadas: (1) segregação jurídica obrigatória entre a contratação de passagens (Lei nº "
        "14.133/2021) e o pagamento de diárias civis à servidora (Lei Municipal nº 6.621/2018), que não se "
        "misturam em um mesmo contrato; (2) no âmbito das passagens, o parcelamento em 02 (dois) itens "
        "independentes — Item 01 (passagem aérea POA ↔ REC) e Item 02 (passagem rodoviária Pelotas ↔ POA) — é "
        "tecnicamente viável e economicamente vantajoso, pois permite que agência de viagens cote ambos os itens "
        "quando competitivo ou que o trecho rodoviário seja adquirido diretamente da concessionária regulada pelo "
        "DAER/RS (sem taxa de intermediação) ou atendido por veículo oficial da SMS, garantindo maior economicidade "
        "sem prejuízo à integração de horários.",
    )

    # 10
    h2(
        doc,
        "10. DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS",
        "Art. 18, § 1º, IX, da Lei Federal nº 14.133/2021 e art. 26, IX, do Decreto Municipal nº 6.799/2023.",
    )
    bullet(
        doc,
        "Habilitação técnica do Município de Pelotas para operacionalizar os espirômetros digitais recebidos por "
        "meio do Novo PAC Saúde, evitando a ociosidade de patrimônio público federal.",
    )
    bullet(
        doc,
        "Implantação da oferta de Tele-espirometria na APS com emissão gratuita de laudos a distância por "
        "pneumologistas do HC-UFMG e acesso a teleconsultoria especializada, sem custo de contratação de médicos "
        "laudadores para o Município.",
    )
    bullet(
        doc,
        "Ampliação da resolubilidade clínica da Atenção Primária no diagnóstico precoce e acompanhamento de asma "
        "e DPOC, reduzindo filas de espera na regulação especializada e internações por condições sensíveis à APS.",
    )
    bullet(
        doc,
        "Efeito multiplicador e economicidade institucional: com investimento estimado em torno de "
        f"{brl(TOTAL_GLOBAL_REF)} para capacitar 01 servidora multiplicadora, evita-se o custeio futuro de "
        "deslocamentos interestaduais de diversos outros profissionais da rede.",
    )
    bullet(
        doc,
        "Cumprimento das metas municipais vinculadas ao Programa SUS Digital (Portarias GM/MS nº 3.691/2024 e "
        "nº 7.495/2025) e ao Termo de Compromisso firmado com o CTS-HC/UFMG.",
    )

    # 11
    h2(
        doc,
        "11. PROVIDÊNCIAS PRÉVIAS AO CONTRATO",
        "Art. 18, § 1º, X, da Lei Federal nº 14.133/2021 e art. 26, X, do Decreto Municipal nº 6.799/2023.",
    )
    bullet(doc, "Concluir o curso teórico EaD no Moodle do CTS-HC/UFMG até 08/12/2026 (Responsável: Servidora designada).")
    bullet(
        doc,
        "Confirmar oficialmente, por e-mail ao CTS-HC/UFMG, o recebimento das orientações e a presença da "
        "profissional em 15/12/2026 (Responsável: Depto. Saúde Digital / Servidora).",
    )
    bullet(
        doc,
        "Providenciar termo de cautela/autorização de saída temporária do espirômetro digital PAC, cabo Micro-USB "
        "e turbina, ou comunicar ao CTS-HC/UFMG caso o equipamento ainda não esteja disponível (Responsável: "
        "Depto. Saúde Digital / Patrimônio SMS).",
    )
    bullet(
        doc,
        "Verificar junto ao setor de Transportes da SMS eventual disponibilidade de veículo oficial para o trecho "
        "Pelotas ↔ Porto Alegre em 14/12 e 16/12/2026 (Responsável: DAF / Transportes SMS).",
    )
    bullet(
        doc,
        "Realizar pesquisa formal de preços (art. 23, Lei nº 14.133/2021) com cotações contemporâneas e verificar "
        "o controle do limite anual de dispensa do art. 75, § 1º (Responsável: Setor de Compras SMS).",
    )
    bullet(
        doc,
        "Confirmar dotação orçamentária (3.3.90.33 e 3.3.90.14), validar o valor vigente da URM não tributária de "
        "novembro/dezembro de 2026 junto à SEFAZ e emitir as Notas de Empenho (Responsável: CCGOF / Contabilidade).",
    )
    bullet(
        doc,
        "Designar formalmente o responsável pelo acompanhamento da emissão dos bilhetes e conferência da prestação "
        "de contas das diárias (Responsável: Diretoria de Atenção Primária / DAF).",
    )

    # 12
    h2(
        doc,
        "12. CONTRATAÇÕES CORRELATAS / INTERDEPENDENTES",
        "Art. 18, § 1º, XI, da Lei Federal nº 14.133/2021 e art. 26, XI, do Decreto Municipal nº 6.799/2023.",
    )
    para(
        doc,
        "Os itens deste processo possuem interdependência direta e simultânea: a passagem rodoviária "
        "(Pelotas ↔ Porto Alegre), a passagem aérea (Porto Alegre ↔ Recife) e a concessão das diárias civis "
        "compõem um único arranjo logístico indissociável para viabilizar a presença da servidora no dia 15/12/2026. "
        "O Termo de Compromisso com o CTS-HC/UFMG e o recebimento do espirômetro via Novo PAC constituem ajustes "
        "institucionais correlatos, sem ônus financeiro de inscrição. Recomenda-se, adicionalmente, que o "
        "Departamento de Saúde Digital e a Diretoria de Atenção Primária planejem oportunamente, em processo "
        "próprio, a eventual aquisição contínua de insumos descartáveis (como filtros/bocais/turbinas) necessários "
        "à operação rotineira dos exames após a capacitação.",
    )

    # 13
    h2(
        doc,
        "13. POSSÍVEIS IMPACTOS AMBIENTAIS E TRATAMENTOS",
        "Art. 18, § 1º, XII, da Lei Federal nº 14.133/2021 e art. 26, XII, do Decreto Municipal nº 6.799/2023.",
    )
    para(
        doc,
        "Os principais impactos ambientais associados decorrem das emissões de gases de efeito estufa inerentes "
        "ao transporte aéreo e rodoviário e do consumo de papel em documentos de viagem. Como medidas mitigadoras, "
        "adotam-se: (i) emissão e conferência 100% digitais dos bilhetes (e-ticket), comprovantes e relatórios; "
        "(ii) priorização de voos diretos ou com menor número de escalas no trecho POA ↔ REC (reduzindo queima de "
        "combustível em decolagens adicionais); (iii) utilização de transporte coletivo rodoviário no trecho "
        "Pelotas ↔ Porto Alegre; e (iv) racionalização do número de viajantes mediante estratégia de profissional "
        "multiplicadora (01 deslocamento em vez de múltiplas viagens interestaduais).",
    )

    # 14
    h2(
        doc,
        "14. ANÁLISE DE RISCOS",
        "Art. 18, X, da Lei Federal nº 14.133/2021 e art. 22, IV, do Decreto Municipal nº 6.799/2023.",
    )
    risk_card(
        doc,
        titulo="1. Elevação abrupta da tarifa aérea ou esgotamento de assentos (alta temporada de dezembro)",
        prob="Alta",
        impacto="Alto",
        situacao="Demora na instrução/empenho resulta em tarifas muito acima da estimativa ou indisponibilidade de voos compatíveis com o evento em 15/12/2026.",
        preventiva="Tramitação em caráter prioritário após juntada deste DFD/ETP; cotação imediata com reserva tarifária (bloqueio) e emissão célere logo após o empenho.",
        resp_prev="Depto. Saúde Digital / Setor de Compras / CCGOF",
        contingencia="Ajuste da estimativa mediante pesquisa atualizada (dentro do limite do art. 75, II) ou escolha de voo alternativo com escala em horário compatível.",
        resp_cont="Setor de Compras / DAF",
    )
    risk_card(
        doc,
        titulo="2. Perda de conexão entre transporte rodoviário (Pelotas–POA) e voo (POA–REC), ou atraso/cancelamento de voo",
        prob="Média",
        impacto="Alto",
        situacao="Atraso na BR-116 ou alteração de malha aérea impede o embarque ou compromete a chegada da servidora até a véspera do curso.",
        preventiva="Seleção de horários com margem mínima de 2h30 a 3h entre chegada em POA e embarque aéreo; preferência por voo direto POA–REC.",
        resp_prev="Setor de Compras / Servidora designada",
        contingencia="Acionamento imediato da companhia aérea/agência para reacomodação (Resolução ANAC nº 400/2016) e comunicação tempestiva à coordenação do CTS-HC/UFMG.",
        resp_cont="Agência/Cia. Aérea / Servidora / Depto. Saúde Digital",
    )
    risk_card(
        doc,
        titulo="3. Não conclusão do módulo EaD prévio ou indisponibilidade/avaria do espirômetro PAC",
        prob="Baixa",
        impacto="Alto",
        situacao="Servidora fica impedida de realizar a prática presencial ou equipamento sofre dano no transporte aéreo.",
        preventiva="Acompanhamento semanal da conclusão do EaD até 08/12/2026; transporte do espirômetro PAC + cabo Micro-USB + turbina exclusivamente em bagagem de mão protegida, com termo de cautela.",
        resp_prev="Servidora designada / Depto. Saúde Digital",
        contingencia="Comunicação prévia por e-mail ao CTS-HC/UFMG caso haja impedimento técnico com o equipamento, conforme previsto nas orientações oficiais do evento.",
        resp_cont="Depto. Saúde Digital",
    )
    risk_card(
        doc,
        titulo="4. Atraso no crédito das diárias antes da viagem",
        prob="Média",
        impacto="Médio",
        situacao="Servidora embarca sem o recurso indenizatório disponível para custear hospedagem, alimentação e deslocamentos urbanos em Recife.",
        preventiva="Tramitação paralela do empenho de diárias (3.3.90.14) com antecedência superior às 72h mínimas do art. 6º da Lei nº 6.621/2018 e conferência prévia dos dados bancários no SIS.",
        resp_prev="DAF / Contabilidade / Tesouraria",
        contingencia="Priorização da liquidação e pagamento bancário ou ressarcimento regulamentar mediante justificativa formal nos termos da Lei nº 6.621/2018.",
        resp_cont="DAF / Financeiro SMS",
    )
    risk_card(
        doc,
        titulo="5. Descumprimento do prazo ou incompletude da prestação de contas das diárias",
        prob="Baixa",
        impacto="Médio",
        situacao="Apontamento pelo Controle Interno/TCE-RS ou obrigação de devolução de valores recebidos (art. 4º e art. 8º da Lei nº 6.621/2018).",
        preventiva="Orientação formal prévia à servidora sobre guarda de notas fiscais/comprovantes de hospedagem, alimentação, deslocamento interno e obtenção do certificado/declaração de participação.",
        resp_prev="Depto. Saúde Digital / Gestor da viagem",
        contingencia="Notificação imediata da servidora para saneamento no prazo de 5 dias úteis após o retorno (até 23/12/2026) ou restituição proporcional aos cofres públicos.",
        resp_cont="DAF / Controle Interno",
    )

    # 15
    h2(
        doc,
        "15. DECLARAÇÃO DE VIABILIDADE DA CONTRATAÇÃO",
        "Art. 18, § 1º, XIII, da Lei Federal nº 14.133/2021 e art. 26, XIII, do Decreto Municipal nº 6.799/2023.",
    )
    para(
        doc,
        f"Diante do estudo técnico realizado, constata-se que: (i) a necessidade pública é concreta, urgente e "
        f"amparada em política nacional estruturante (Novo PAC Saúde, Programa SUS Digital e Termo de Compromisso "
        f"com o CTS-HC/UFMG); (ii) o custo global preliminar estimado ({brl(SUBTOTAL_PASSAGENS)} em passagens + "
        f"10 URMs não tributárias em diárias) é módico, proporcional e altamente vantajoso frente ao retorno "
        f"assistencial e ao efeito multiplicador na rede municipal; (iii) o enquadramento das passagens em "
        f"Dispensa de Licitação por valor (art. 75, inciso II, da Lei nº 14.133/2021) e das diárias na Lei "
        f"Municipal nº 6.621/2018 atende integralmente à legalidade; e (iv) os riscos identificados são "
        f"administráveis mediante as ações preventivas estabelecidas.",
    )
    para(
        doc,
        "Assim, a equipe de planejamento DECLARA A VIABILIDADE técnica, operacional, jurídica e orçamentária da "
        "contratação das passagens e da concessão concomitante das diárias civis, recomendando o prosseguimento "
        "célere do Processo SIS nº 15912/2026.",
        bold_prefix="Conclusão: ",
    )

    # 16
    h2(doc, "16. EQUIPE DE PLANEJAMENTO")
    bloco_assinaturas(doc, incluir_autorizacao_secretaria=True)

    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    print(f"OK {path}")


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    gerar_dfd(OUT / "TeleEspirometria_DFD_Diarias_Passagens.docx")
    gerar_etp(OUT / "TeleEspirometria_ETP_Diarias_Passagens.docx")
    readme = OUT / "README.md"
    readme.write_text(
        f"""# Projeto Tele-espirometria — DFD e ETP (Diárias e Passagens)

Pacote documental no padrão oficial SARH/SMS Pelotas (Lei Federal nº 14.133/2021, Decreto Municipal nº 6.799/2023 e Lei Municipal nº 6.621/2018) para instrução do **Processo SIS nº {SIS_NUM}**.

## Peças geradas

| Ordem | Documento | Arquivo |
|------:|-----------|---------|
| 1 | **DFD** | `TeleEspirometria_DFD_Diarias_Passagens.docx` |
| 2 | **ETP** | `TeleEspirometria_ETP_Diarias_Passagens.docx` |

## Síntese do processo

- **Expediente:** Processo SIS nº {SIS_NUM} (abertura em {SIS_ABERTURA}, acordo da Secretária em 09/11/2026)
- **Unidade Demandante:** {UNIDADE}
- **Servidora Designada:** {SERVIDORA} ({CARGO_SERVIDORA}, Matrícula {MATRICULA_SERVIDORA})
- **Evento:** Treinamento Prático Presencial — Projeto Tele-espirometria / TeleRESP (CTS-HC/UFMG + SAPS/MS + Novo PAC)
- **Data e Local:** {EVENTO_DATA}, {EVENTO_HORARIO} — {EVENTO_LOCAL}
- **Deslocamento:** {SAIDA_DATA} ({SAIDA_HORA}) a {RETORNO_DATA} ({RETORNO_HORA}) — 60 horas de afastamento
- **Passagens (Lei 14.133/2021 — art. 75, II — Elemento 3.3.90.33):** estimativa preliminar de {brl(SUBTOTAL_PASSAGENS)} (aérea POA↔REC ref. {brl(VAL_AEREA)} + 2 trechos rodoviários Pelotas↔POA ref. {brl(VAL_RODOV_TOTAL)})
- **Diárias (Lei Municipal nº 6.621/2018 — Elemento 3.3.90.14):** 2 diárias integrais (8 URMs) + 1 diária estada (2 URMs) = **10 URMs não tributárias** (ref. mínima {brl(SUBTOTAL_DIARIAS_REF)} pela URM de R$ 111,27 — atualizar pela tabela SEFAZ de nov–dez/2026)

## Pendências operacionais antes do protocolo final

1. Atribuir numeração interna ao DFD (`nº [___]/2026`) e informar matrículas dos signatários da equipe de planejamento.
2. Confirmar Projeto/Atividade e Fonte de Recurso com a CCGOF/Contabilidade (elementos `3.3.90.33` e `3.3.90.14`).
3. Confirmar junto à SEFAZ/Contabilidade o valor vigente da **URM não tributária** de novembro/dezembro de 2026 para o cálculo definitivo das 10 URMs.
4. Juntar mapa de pesquisa formal de preços (mínimo 3 cotações contemporâneas — art. 23, Lei nº 14.133/2021) para emissão imediata das passagens.
5. Verificar disponibilidade de veículo oficial da SMS para o trecho Pelotas ↔ Porto Alegre (14/12 e 16/12/2026).
6. Providenciar conclusão do módulo EaD (até 08/12/2026), e-mail de confirmação de presença ao CTS-HC/UFMG e termo de cautela do espirômetro digital do Novo PAC.

## Regenerar

```bash
pip install python-docx
python3 scripts/gerar_tele_espirometria.py
```
""",
        encoding="utf-8",
    )
    print(f"OK {readme}")


if __name__ == "__main__":
    main()
