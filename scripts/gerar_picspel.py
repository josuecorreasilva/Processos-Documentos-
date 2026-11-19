#!/usr/bin/env python3
"""Gera ETP, TR e DFD para aquisição de materiais PICSPel (Emenda 43530013).

Padrão documental RP02–RP13 / SMS Pelotas (brasão, Arial, Página X de Y).
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "docs" / "05-PICSPEL-MATERIAIS"
BRASAO = ROOT / "docs" / "assets" / "brasao_rp_header.png"
if not BRASAO.exists():
    BRASAO = ROOT / "docs" / "assets" / "brasao_pelotas.png"

GESTOR = ("Dariane Lima Portela", "44151", "Gestor")
FISCAL = ("Josué Corrêa da Silva", "43351", "Fiscal")
SECRETARIA = ("Angela Moreira Vitória", "Secretária Municipal de Saúde")
DIRETORA = ("Cândida Silveira Rodrigues", "Diretora de Atenção Primária")

# Fonte: Lista de Materiais PICSPel — Emenda Parlamentar nº 43530013
ITENS = [
    # nº, descrição, unid, qtd, vu, vt, natureza, destino
    (1, "Bacia para escalda-pés, material plástico resistente, capacidade adequada a imersão dos pés, superfície lisa e de fácil higienização", "UN", 120, 35.00, 4200.00, "consumo", "eMulti (15 un. × 8 equipes)"),
    (2, "Protetor descartável para bacias de escalda-pés (pacote com 100 unidades), material plástico/filme adequado ao contato com água e pele", "PCT", 80, 40.00, 3200.00, "consumo", "eMulti (10 pct. × 8 equipes)"),
    (3, "Placas para auriculoterapia (caixa com 30 placas), padrão clínico para aplicação em pontos auriculares", "CX", 50, 190.00, 9500.00, "consumo", "UBS / rede APS"),
    (4, "Caixa de som portátil, com bateria recarregável, entrada Bluetooth/USB/AUX, potência adequada a grupos terapêuticos em ambiente de UBS/eMulti", "UN", 8, 600.00, 4800.00, "permanente", "eMulti (1 un. × 8 equipes)"),
    (5, "Pinça Adson auricular 12 cm, reta, aço inoxidável, ponta adequada a manipulação de sementes/cristais em auriculoterapia", "UN", 100, 35.00, 3500.00, "consumo", "UBS / rede APS"),
    (6, "Óleo essencial 10 ml — aromas: lavanda, peppermint (hortelã-pimenta), melaleuca e laranja doce; frasco com tampa conta-gotas; grau terapêutico/aromaterapia", "FR", 250, 60.00, 15000.00, "consumo", "UBS (kits bimestrais / 4 aromas)"),
    (7, "Óleo carreador 100 ml — óleo de coco, frasco adequado a diluição de óleos essenciais em aromaterapia", "UN", 50, 60.00, 3000.00, "consumo", "UBS / eMulti"),
    (8, "Frasco âmbar 5 ml com gotejador, para acondicionamento e diluição de blend de óleos essenciais", "UN", 400, 5.00, 2000.00, "consumo", "UBS / eMulti (50 un. × 8)"),
    (9, "Inalador nasal para aromaterapia, reutilizável ou de uso individual conforme especificação de mercado, compatível com óleos essenciais", "UN", 240, 7.00, 1680.00, "consumo", "eMulti (30 un. × 8 equipes)"),
    (10, "Garrafa térmica, capacidade mínima 500 ml, manutenção de temperatura, uso em preparo/conservação de água para práticas de escalda-pés e chás terapêuticos", "UN", 24, 150.00, 3600.00, "permanente", "eMulti (3 un. × 8 equipes)"),
    (11, "Maleta organizadora rígida, com divisórias, para transporte e acondicionamento de materiais de PICS nas equipes eMulti", "UN", 8, 150.00, 1200.00, "permanente", "eMulti (1 un. × 8 equipes)"),
    (12, "Jarra elétrica / chaleira elétrica, capacidade adequada a aquecimento de água para escalda-pés e preparações, com desligamento automático", "UN", 8, 100.00, 800.00, "permanente", "eMulti (1 un. × 8 equipes)"),
    (13, "Tapete de yoga / colchonete para práticas corporais e meditação, antiderrapante, espessura adequada a uso em solo", "UN", 20, 80.00, 1600.00, "consumo", "cursos / práticas coletivas"),
    (14, "Apalpador com mola modelo Nogier Fava, para localização de pontos em auriculoterapia", "UN", 100, 40.00, 3200.00, "consumo", "UBS / rede APS e cursos"),
    (15, "Plastificadora A4, para plastificação de materiais educativos e fichas de pontos de PICS", "UN", 1, 400.00, 400.00, "permanente", "coordenação PICSPel"),
    (16, "Plástico para plastificação A4 (pacote com 100 unidades), espessura compatível com a plastificadora", "PCT", 2, 160.00, 320.00, "consumo", "coordenação PICSPel"),
]

VALOR_TOTAL = sum(i[5] for i in ITENS)  # 58000.00
VALOR_CONSUMO = sum(i[5] for i in ITENS if i[6] == "consumo")
VALOR_PERMANENTE = sum(i[5] for i in ITENS if i[6] == "permanente")

EMENDA = "43530013"
RECURSO_TOTAL = 100_000.00
RECURSO_CONSUMO = 60_000.00
RECURSO_SERVICOS = 40_000.00

EQUIPES = [
    "Centro/Porto",
    "Fragata 1",
    "Fragata 2",
    "Três Vendas 1",
    "Três Vendas 2",
    "Areal",
    "Colônia",
    "Praias",
]


def _font(run, *, name="Arial", size=11, bold=False):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold


def _shade(cell, fill="EDEDED"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    shd.set(qn("w:val"), "clear")
    tcPr.append(shd)


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
    sec.top_margin = Cm(2.5)
    sec.bottom_margin = Cm(2.0)
    sec.left_margin = Cm(2.0)
    sec.right_margin = Cm(2.0)
    sec.header_distance = Cm(1.25)

    style = doc.styles["Normal"]
    style.font.name = "Arial"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
    style.font.size = Pt(11)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    style.paragraph_format.space_after = Pt(6)

    header = sec.header
    header.is_linked_to_previous = False
    header.paragraphs[0].clear()
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.space_after = Pt(0)
    if BRASAO.exists():
        run = hp.add_run()
        run.add_picture(str(BRASAO), width=Cm(1.7))
    for _ in range(5):
        p = header.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
    for text in ["ESTADO DO RIO GRANDE DO SUL", "MUNICÍPIO DE PELOTAS", "SECRETARIA MUNICIPAL DE SAÚDE"]:
        p = header.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(text)
        _font(r, size=10, bold=False)

    footer = sec.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.clear()
    _page_fields(fp)
    return doc


def title(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run(text)
    _font(r, size=12, bold=True)
    return p


def h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    _font(r, size=11, bold=True)
    return p


def p(doc, text, *, bold=False, justify=True):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if justify else WD_ALIGN_PARAGRAPH.LEFT
    if bold and ":" in text:
        label, rest = text.split(":", 1)
        r1 = para.add_run(label + ":")
        _font(r1, bold=True)
        r2 = para.add_run(rest)
        _font(r2, bold=False)
    else:
        r = para.add_run(text)
        _font(r, bold=bold)
    return para


def bullet_line(doc, text):
    para = doc.add_paragraph()
    para.paragraph_format.left_indent = Cm(0.5)
    r = para.add_run("• " + text)
    _font(r)
    return para


def checkbox(doc, text, checked=False):
    mark = "X" if checked else " "
    para = doc.add_paragraph()
    para.paragraph_format.space_after = Pt(2)
    r = para.add_run(f"( {mark} ) {text}")
    _font(r)
    return para


def table(doc, headers, rows, *, header_shade=True):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = ""
        r = cell.paragraphs[0].add_run(h)
        _font(r, size=8, bold=True)
        if header_shade:
            _shade(cell)
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            cell = t.rows[ri].cells[ci]
            cell.text = ""
            r = cell.paragraphs[0].add_run(str(val))
            _font(r, size=8)
    doc.add_paragraph()
    return t


def brl(v: float) -> str:
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def assinaturas_equipe(doc):
    h2(doc, "EQUIPE DE PLANEJAMENTO / FISCALIZAÇÃO")
    p(
        doc,
        "Certificamos que somos responsáveis pela elaboração do presente instrumento e que o mesmo está em consonância com a legislação vigente.",
    )
    t = doc.add_table(rows=1, cols=2)
    for i, pessoa in enumerate((GESTOR, FISCAL)):
        cell = t.rows[0].cells[i]
        cell.text = ""
        for j, line in enumerate(
            [
                "__________________________",
                pessoa[0],
                f"Matrícula: {pessoa[1]}",
                pessoa[2],
            ]
        ):
            para = cell.paragraphs[0] if j == 0 else cell.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_after = Pt(0)
            r = para.add_run(line)
            _font(r, size=9)
    doc.add_paragraph()
    p(doc, "Pelotas/RS, _____ de ______________ de 2026.", justify=False)
    doc.add_paragraph()
    for block in [
        [DIRETORA[0], DIRETORA[1]],
        [SECRETARIA[0], SECRETARIA[1] + " — Autorizo"],
    ]:
        line = doc.add_paragraph()
        line.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = line.add_run("_" * 42)
        _font(r)
        for text in block:
            para = doc.add_paragraph()
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            para.paragraph_format.space_after = Pt(0)
            rr = para.add_run(text)
            _font(rr, size=10)
        doc.add_paragraph()


def tabela_quantitativos(doc):
    rows = []
    for n, desc, unid, qtd, vu, vt, nat, dest in ITENS:
        short = desc if len(desc) <= 90 else desc[:87] + "..."
        rows.append(
            [
                f"{n:02d}",
                short,
                unid,
                f"{qtd}",
                brl(vu),
                brl(vt),
            ]
        )
    rows.append(["", "VALOR TOTAL ESTIMADO", "", "", "", brl(VALOR_TOTAL)])
    table(
        doc,
        ["Item", "Descrição resumida", "Unid.", "Qtd.", "Valor unit.", "Valor total"],
        rows,
    )


def tabela_natureza(doc):
    rows = []
    for n, desc, unid, qtd, vu, vt, nat, dest in ITENS:
        rows.append(
            [
                f"{n:02d}",
                "Consumo" if nat == "consumo" else "Permanente",
                dest,
                brl(vt),
            ]
        )
    rows.append(["", "Subtotal consumo", "", brl(VALOR_CONSUMO)])
    rows.append(["", "Subtotal permanente", "", brl(VALOR_PERMANENTE)])
    table(doc, ["Item", "Natureza", "Destinação", "Valor"], rows)


# ---------------------------------------------------------------------------
# 1) ETP
# ---------------------------------------------------------------------------
def gerar_etp(path: Path):
    doc = new_doc()
    title(doc, "ESTUDO TÉCNICO PRELIMINAR – PICSPel (MATERIAIS)")

    h2(doc, "1 IDENTIFICAÇÃO DO PROCESSO")
    p(
        doc,
        "1.1 Unidade responsável pela contratação: Secretaria Municipal de Saúde – Diretoria de Atenção Primária – Programa Municipal de Práticas Integrativas e Complementares em Saúde (PICSPel).",
    )
    p(
        doc,
        "1.2 Objeto da contratação: Aquisição de materiais de consumo e de apoio (incluindo itens de natureza permanente de pequeno porte) destinados à implementação e ao fortalecimento das Práticas Integrativas e Complementares em Saúde (PICS) na rede de Atenção Primária à Saúde do Município de Pelotas/RS, compreendendo 16 (dezesseis) itens, para uso pelas equipes eMulti e Unidades Básicas de Saúde (UBS), bem como para suporte aos cursos de formação previstos no âmbito da Emenda Parlamentar nº "
        + EMENDA
        + ".",
    )
    p(
        doc,
        "Os materiais destinam-se a práticas de auriculoterapia, aromaterapia, escalda-pés, práticas corporais/meditação e atividades educativas correlatas, contribuindo para a ampliação da oferta de PICS, a humanização do cuidado e a resolutividade da Atenção Primária.",
    )
    p(
        doc,
        "1.3 Número do processo: Memorando nº [______]/2026 – DFD nº [___]/2026 – Emenda Parlamentar nº "
        + EMENDA
        + ".",
    )
    p(
        doc,
        "1.4 Fonte de recurso: Emenda Parlamentar nº "
        + EMENDA
        + ", no valor global de "
        + brl(RECURSO_TOTAL)
        + ", sendo "
        + brl(RECURSO_CONSUMO)
        + " destinados a material de consumo e "
        + brl(RECURSO_SERVICOS)
        + " a outros serviços de terceiros – pessoa jurídica (cursos/capacitação). O presente ETP abrange exclusivamente a parcela de materiais (estimativa de "
        + brl(VALOR_TOTAL)
        + "), residual da cota de consumo. A contratação dos serviços de capacitação será objeto de processo próprio.",
    )

    h2(doc, "2 DESCRIÇÃO DA NECESSIDADE")
    p(
        doc,
        "A Política Nacional de Práticas Integrativas e Complementares (PNPIC) e a organização local do PICSPel demandam insumos específicos para a oferta segura e regular de PICS na Atenção Primária. A Secretaria Municipal de Saúde necessita adquirir o conjunto de materiais descritos na Lista de Materiais PICSPel, de modo a municiar as equipes eMulti e as UBS de referência, bem como viabilizar a realização de cursos de formação com 50 (cinquenta) vagas (dois grupos de 25 participantes).",
    )
    p(
        doc,
        "A ausência desses materiais impede ou limita a oferta de auriculoterapia, aromaterapia, escalda-pés e práticas corporais, reduzindo a resolutividade e a equidade do cuidado. A demanda foi dimensionada a partir do planejamento do PICSPel, considerando 08 (oito) equipes eMulti já contempladas e a necessidade de estoque operacional nas UBS e na coordenação do programa.",
    )
    p(doc, "Equipes eMulti previstas (quantitativos por equipe multiplicados por 8):")
    for eq in EQUIPES:
        bullet_line(doc, eq)
    bullet_line(doc, "Eventual inclusão de até 4 equipes adicionais a definir — sem alteração automática dos quantitativos deste ETP; eventual suplementação dependerá de novo planejamento.")
    p(
        doc,
        "Após a chegada dos materiais, será realizada reunião com as equipes eMulti e com as equipes da APS de referência, para organizar armazenamento, controle e fluxos de utilização, conforme encaminhamento da Lista de Materiais PICSPel.",
    )

    h2(doc, "3 DA CONTRATAÇÃO E DO PLANEJAMENTO")
    p(
        doc,
        "A contratação tem por finalidade a aquisição pontual dos materiais necessários à execução do PICSPel com recursos da Emenda Parlamentar nº "
        + EMENDA
        + ". Os itens possuem padrões de desempenho e qualidade objetivamente definíveis por especificações usuais de mercado, enquadrando-se como bens comuns, nos termos da Lei Federal nº 14.133/2021.",
    )
    p(
        doc,
        "Forma de contratação sugerida: Pregão Eletrônico, critério de menor preço por item, com julgamento e adjudicação por item, observando-se os princípios da legalidade, isonomia, competitividade, economicidade, planejamento e seleção da proposta mais vantajosa.",
    )
    p(
        doc,
        "Justifica-se a aquisição pontual (e não o Sistema de Registro de Preços) porque: (i) os quantitativos estão vinculados a recurso de emenda com objeto e prazo definidos; (ii) a demanda foi previamente dimensionada para as equipes e cursos; e (iii) não se trata de consumo contínuo indeterminado a longo prazo, e sim de implantação/fortalecimento com quantitativo certo.",
    )
    p(
        doc,
        "Previsão no Plano de Contratações Anual: [informar se a contratação consta do PCA e o respectivo número do item — se não constar, justificar inclusão excepcional vinculada à emenda parlamentar].",
    )
    p(
        doc,
        "Serão observadas as etapas da Lei nº 14.133/2021, incluindo o presente Estudo Técnico Preliminar, a pesquisa de preços (art. 23), o Termo de Referência, a designação de gestor e fiscal e a publicação no PNCP.",
    )

    h2(doc, "4 REQUISITOS DA CONTRATAÇÃO")
    p(
        doc,
        "Os produtos deverão atender integralmente às especificações técnicas de cada item, ser novos, de primeiro uso, livres de vícios ou defeitos de fabricação, entregues em embalagem íntegra, com identificação em português do fabricante/fornecedor, e, quando aplicável, com lote, data de fabricação e validade.",
    )
    p(doc, "Requisitos específicos por família de produtos:")
    bullet_line(
        doc,
        "Auriculoterapia: placas, pinças Adson e apalpadores Nogier deverão ser adequados ao uso clínico em APS, em material compatível com higienização e reprocessamento quando cabível; pinças em aço inoxidável.",
    )
    bullet_line(
        doc,
        "Aromaterapia: óleos essenciais e carreador com rotulagem clara (nome do óleo, volume, lote e validade); frascos âmbar com gotejador íntegros; inaladores nasais seguros para uso individual.",
    )
    bullet_line(
        doc,
        "Escalda-pés: bacias de fácil higienização; protetores descartáveis em quantidade suficiente; jarra elétrica com dispositivo de segurança; garrafas térmicas com vedação adequada.",
    )
    bullet_line(
        doc,
        "Práticas corporais: tapetes antiderrapantes, em condições de uso coletivo sob supervisão.",
    )
    bullet_line(
        doc,
        "Apoio educativo/logístico: caixa de som portátil funcional; maletas organizadoras; plastificadora A4 e plásticos compatíveis.",
    )
    p(
        doc,
        "Produtos sujeitos a regulamentação sanitária (quando aplicável a cosméticos/óleos) deverão observar as normas da ANVISA. Não haverá exigência de marca. A contratada responderá pela substituição, sem ônus, dos produtos defeituosos ou em desconformidade, no prazo máximo de 5 (cinco) dias úteis do recebimento da notificação.",
    )
    p(
        doc,
        "Habilitação: jurídica; regularidade fiscal e trabalhista; qualificação econômico-financeira proporcional; e, se pertinente ao item, comprovação de regularidade sanitária. Exigências serão proporcionais ao valor e à complexidade de cada item, vedadas exigências desarrazoadas.",
    )

    h2(doc, "5 ESTIMATIVA DAS QUANTIDADES")
    p(
        doc,
        "As quantidades foram estimadas com base na Lista de Materiais para Aquisição do PICSPel, considerando: (a) 08 equipes eMulti, com quantitativos unitários por equipe multiplicados por 8; (b) demanda das UBS para auriculoterapia e aromaterapia (incluindo 4 kits bimestrais de óleos); (c) suporte a cursos com 50 vagas; e (d) itens de coordenação (plastificadora e plásticos).",
    )
    p(
        doc,
        "Por se tratar de aquisição pontual com quantitativos certos, a Administração adquirirá os totais adjudicados, observados o empenho e a disponibilidade orçamentária da emenda.",
    )
    tabela_quantitativos(doc)
    p(doc, "Classificação por natureza da despesa (para fins de empenho e patrimônio):")
    tabela_natureza(doc)
    p(
        doc,
        "Observação: a Contabilidade / CCGOF deverá validar a classificação orçamentária dos itens de natureza permanente (caixa de som, garrafas térmicas, maletas, jarra elétrica e plastificadora), ainda que custeados no âmbito da cota de materiais da emenda. Itens permanentes deverão ser tombados quando exigido pela legislação patrimonial municipal.",
    )

    h2(doc, "6 LEVANTAMENTO DE MERCADO")
    p(
        doc,
        "Foi realizado levantamento preliminar de mercado com base na Lista de Materiais PICSPel, que já apresenta valores unitários de referência totalizando "
        + brl(VALOR_TOTAL)
        + ". Verificou-se a existência de fornecedores locais e regionais (comércio de artigos de saúde, aromaterapia, esportivos, eletroportáteis e papelaria) aptos a fornecer os itens, isoladamente ou em conjunto.",
    )
    p(doc, "Alternativas consideradas:")
    bullet_line(doc, "(i) licitação própria por Pregão Eletrônico com julgamento por item;")
    bullet_line(doc, "(ii) adesão a ata de registro de preços de outro órgão ou do Consórcio Público do Extremo Sul;")
    bullet_line(doc, "(iii) aquisição por dispensa de licitação, se o valor atualizado e a legislação vigente autorizarem o enquadramento.")
    p(
        doc,
        "Conclui-se, preliminarmente, que o Pregão Eletrônico com julgamento por item é a solução mais adequada, por ampliar a competitividade (permitindo que empresas especializadas disputem apenas os itens de seu ramo) e por garantir isonomia e publicidade no PNCP. A adesão a atas existentes deverá ser verificada antes da abertura do certame; caso se mostre mais vantajosa para item específico, o item poderá ser excluído do edital com a devida justificativa.",
    )
    p(
        doc,
        "A pesquisa de preços formal será consolidada nos termos do art. 23 da Lei nº 14.133/2021 (painel de preços, contratações similares, mídia especializada e/ou cotações de fornecedores), adotando-se critério estatístico justificado (mediana/média), com mapa anexo ao processo.",
    )

    h2(doc, "7 ESTIMATIVA DO VALOR DA CONTRATAÇÃO")
    p(
        doc,
        "Valor total estimado com base na Lista de Materiais PICSPel: "
        + brl(VALOR_TOTAL)
        + " ("
        + "cinquenta e oito mil reais"
        + "), compatível com a parcela de "
        + brl(RECURSO_CONSUMO)
        + " da Emenda Parlamentar nº "
        + EMENDA
        + " destinada a materiais.",
    )
    p(
        doc,
        "Detalhamento preliminar: material de consumo ≈ "
        + brl(VALOR_CONSUMO)
        + "; material permanente de apoio ≈ "
        + brl(VALOR_PERMANENTE)
        + ". Os valores unitários são referenciais e serão atualizados/validados pela pesquisa de preços oficial antes da publicação do edital.",
    )

    h2(doc, "8 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO")
    p(
        doc,
        "A solução consiste na aquisição, mediante Pregão Eletrônico, dos 16 itens necessários ao PICSPel, com entrega única ou parcelada em até 2 (duas) remessas no Almoxarifado da SMS ou em local indicado pela Diretoria de Atenção Primária / coordenação PICSPel, seguida de distribuição às equipes eMulti e UBS, conforme plano de entrega interno.",
    )
    p(
        doc,
        "O ciclo de vida do objeto compreende: planejamento e formalização da demanda; aquisição; recebimento provisório e definitivo; eventual tombamento de permanentes; armazenamento; distribuição; utilização assistencial/educativa; higienização/reposição de descartáveis; e descarte ambientalmente adequado de embalagens e resíduos.",
    )

    h2(doc, "9 JUSTIFICATIVA PARA PARCELAMENTO")
    p(
        doc,
        "A contratação será realizada com parcelamento por itens e julgamento por menor preço por item, nos termos do art. 40 da Lei nº 14.133/2021, por ser tecnicamente viável e economicamente vantajoso: os itens possuem naturezas mercadológicas distintas (descartáveis clínicos, óleos, eletroportáteis, artigos esportivos e de escritório), e a divisão permite a participação de fornecedores especializados, ampliando a competitividade. A reunião dos itens em um mesmo processo decorre da unidade de finalidade (fortalecimento do PICSPel com recurso da mesma emenda) e da racionalização administrativa.",
    )

    h2(doc, "10 DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS")
    p(doc, "Com a contratação, pretende-se:")
    bullet_line(doc, "viabilizar a oferta regular de PICS (auriculoterapia, aromaterapia, escalda-pés e práticas corporais) nas equipes eMulti e UBS;")
    bullet_line(doc, "capacitar 50 profissionais/interessados nos cursos previstos, com infraestrutura material adequada;")
    bullet_line(doc, "padronizar kits e fluxos de utilização entre as 8 equipes eMulti contempladas;")
    bullet_line(doc, "fortalecer a atenção integral e a humanização do cuidado na APS;")
    bullet_line(doc, "aplicar de forma eficiente e transparente os recursos da Emenda Parlamentar nº " + EMENDA + ";")
    bullet_line(doc, "obter preços vantajosos por meio da competição no pregão eletrônico.")

    h2(doc, "11 PROVIDÊNCIAS PRÉVIAS AO CONTRATO")
    bullet_line(doc, "Formalização do DFD e aprovação da autoridade competente;")
    bullet_line(doc, "elaboração do Termo de Referência com especificações e cronograma de entrega;")
    bullet_line(doc, "pesquisa de preços formal (art. 23) e mapa comparativo;")
    bullet_line(doc, "verificação de atas vigentes (Consórcio / outros órgãos) para eventual adesão vantajosa;")
    bullet_line(doc, "confirmação da dotação, fonte e classificação da despesa com a Contabilidade;")
    bullet_line(doc, "designação de gestor e fiscal do contrato;")
    bullet_line(doc, "preparação do Almoxarifado / locais de armazenamento nas eMulti e UBS;")
    bullet_line(doc, "planejamento da reunião de distribuição e controle pós-entrega.")

    h2(doc, "12 CONTRATAÇÕES CORRELATAS/INTERDEPENDENTES")
    p(
        doc,
        "A parcela de "
        + brl(RECURSO_SERVICOS)
        + " da mesma emenda, destinada a outros serviços de terceiros – pessoa jurídica (cursos/capacitação para 50 vagas), constitui contratação correlata, porém autônoma, a ser formalizada em processo próprio. Não há interdependência técnica que impeça a aquisição dos materiais antes ou em paralelo aos serviços, recomenda-se, contudo, alinhamento de cronogramas para que os insumos estejam disponíveis preferencialmente até a realização dos cursos.",
    )

    h2(doc, "13 POSSÍVEIS IMPACTOS AMBIENTAIS")
    p(
        doc,
        "Os impactos decorrem principalmente da geração de resíduos de embalagens e de protetores descartáveis de bacias. Deverão ser observados a Política Nacional de Resíduos Sólidos, as diretrizes municipais de coleta seletiva e, quando houver contato com fluidos corporais, o Plano de Gerenciamento de Resíduos de Serviços de Saúde das unidades. Óleos e frascos deverão ser armazenados de forma a evitar vazamentos. Equipamentos elétricos deverão observar eficiência energética básica e descarte adequado ao fim da vida útil (logística reversa quando aplicável).",
    )

    h2(doc, "14 ANÁLISE DE RISCOS")
    p(doc, "Risco 1: Atraso na entrega")
    p(doc, "a) Probabilidade: Média. b) Impacto: Médio.")
    p(doc, "c) Dano: adiamento da oferta de PICS e dos cursos.")
    p(doc, "d) Ação preventiva: prazo de entrega claro (até 30 dias corridos da Nota de Empenho, salvo prazo específico no TR); sanções contratuais; acompanhamento pelo fiscal.")
    p(doc, "e) Ação de contingência: notificação, aplicação de sanções (Decreto Municipal nº 6.799/2023, no que couber) e, se necessário, remanejamento de cronograma dos cursos.")

    p(doc, "Risco 2: Entrega em desacordo com especificações")
    p(doc, "a) Probabilidade: Média. b) Impacto: Alto.")
    p(doc, "c) Dano: inutilização do item e prejuízo à prática clínica/educativa.")
    p(doc, "d) Ação preventiva: especificações claras no TR; recebimento provisório/definitivo; amostragem quando justificada.")
    p(doc, "e) Ação de contingência: recusa, substituição em 5 dias úteis e sanções cabíveis.")

    p(doc, "Risco 3: Itens desertos ou fracassados / variação de preços")
    p(doc, "a) Probabilidade: Média. b) Impacto: Médio.")
    p(doc, "c) Dano: lacuna em kit de equipe ou necessidade de novo procedimento.")
    p(doc, "d) Ação preventiva: julgamento por item; pesquisa de preços atualizada; divulgação adequada no PNCP.")
    p(doc, "e) Ação de contingência: republicação do item, dispensa cabível ou realocação interna de materiais.")

    p(doc, "Risco 4: Controle patrimonial e armazenamento inadequado")
    p(doc, "a) Probabilidade: Média. b) Impacto: Médio.")
    p(doc, "c) Dano: extravio de permanentes ou deterioração de óleos/descartáveis.")
    p(doc, "d) Ação preventiva: reunião pós-entrega; fichas de controle por equipe; tombamento dos permanentes.")
    p(doc, "e) Ação de contingência: apuração de responsabilidade e reposição.")

    h2(doc, "15 DECLARAÇÃO DE VIABILIDADE DA CONTRATAÇÃO")
    p(
        doc,
        "A solução proposta mostra-se tecnicamente adequada, operacionalmente viável e economicamente compatível com o recurso da Emenda Parlamentar nº "
        + EMENDA
        + ". Há fornecedores no mercado, os quantitativos estão fundamentados no planejamento do PICSPel e os riscos identificados são administráveis. Declara-se a viabilidade técnica, operacional, econômica e administrativa da contratação, recomendando-se o prosseguimento com a elaboração do Termo de Referência e a instauração de Pregão Eletrônico para aquisição dos materiais.",
    )

    h2(doc, "16 EQUIPE DE PLANEJAMENTO")
    p(
        doc,
        "Equipe de planejamento – Programa PICSPel, Diretoria de Atenção Primária e área administrativa da SMS, com apoio do Almoxarifado para recebimento e distribuição.",
    )
    assinaturas_equipe(doc)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    print(f"OK {path}")


# ---------------------------------------------------------------------------
# 2) TR
# ---------------------------------------------------------------------------
def gerar_tr(path: Path):
    doc = new_doc()
    title(doc, "TERMO DE REFERÊNCIA – PICSPel (MATERIAIS)")

    h2(doc, "1 DO OBJETO, VALOR E CONDIÇÕES GERAIS DA CONTRATAÇÃO")
    p(
        doc,
        "1.1 Aquisição de materiais de consumo e de apoio destinados ao Programa Municipal de Práticas Integrativas e Complementares em Saúde (PICSPel), compreendendo 16 (dezesseis) itens, conforme especificações e quantitativos deste Termo de Referência, com recursos da Emenda Parlamentar nº "
        + EMENDA
        + ".",
    )
    p(doc, "1.2 Os bens são comuns, não de luxo, com padrões de qualidade objetivamente definíveis.")
    p(
        doc,
        "1.3 A entrega deverá ocorrer em até 30 (trinta) dias corridos contados do recebimento da Nota de Empenho / ordem de fornecimento, podendo ser fracionada em até 2 (duas) remessas, desde que o prazo final seja respeitado.",
    )
    p(doc, "1.4 Valor total estimado: " + brl(VALOR_TOTAL) + ", conforme mapa de preços a ser consolidado (art. 23 da Lei nº 14.133/2021).")
    p(doc, "1.5 Critério de julgamento: menor preço por item, Pregão Eletrônico.")
    tabela_quantitativos(doc)

    h2(doc, "2 DA FUNDAMENTAÇÃO DA CONTRATAÇÃO")
    p(
        doc,
        "2.1 A contratação fundamenta-se na Lei Federal nº 14.133/2021, na Política Nacional de Práticas Integrativas e Complementares e no planejamento do PICSPel, com lastro orçamentário na Emenda Parlamentar nº "
        + EMENDA
        + " ("
        + brl(RECURSO_CONSUMO)
        + " para materiais).",
    )
    p(
        doc,
        "2.2 A necessidade e a solução escolhida estão detalhadas no Estudo Técnico Preliminar – PICSPel (Materiais), do qual este TR é correlato.",
    )

    h2(doc, "3 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO CONSIDERADO O CICLO DE VIDA DO OBJETO")
    p(
        doc,
        "Fornecimento dos itens adjudicados, transporte até o local de entrega, recebimento e conferência, eventual tombamento de bens permanentes, armazenamento, distribuição às equipes eMulti/UBS, utilização assistencial e educativa, e descarte adequado de embalagens/resíduos.",
    )

    h2(doc, "4 GESTÃO/ANÁLISE DE RISCOS")
    p(
        doc,
        "Os riscos de atraso, desconformidade, itens desertos e falhas de armazenamento foram tratados no ETP. Mitigações principais: prazo e sanções claros; recebimento em duas etapas; julgamento por item; controle interno pós-entrega com as equipes eMulti.",
    )

    h2(doc, "5 REQUISITOS DA CONTRATAÇÃO")
    p(doc, "5.1 Especificações técnicas:")
    for n, desc, unid, qtd, vu, vt, nat, dest in ITENS:
        p(doc, f"Item {n:02d} ({unid} — qtd. {qtd} — {dest}): {desc}.")
    p(
        doc,
        "5.2 Todos os bens deverão ser novos, de primeiro uso, com embalagem íntegra e identificação em português. Óleos essenciais e carreador deverão apresentar rotulagem com identificação do produto, volume, lote e validade mínima de 12 (doze) meses na data da entrega, salvo prazo diverso justificado no mapa de especificação.",
    )
    p(
        doc,
        "5.3 Habilitação jurídica, regularidade fiscal e trabalhista, e qualificação econômico-financeira proporcional. Documentação sanitária quando exigível ao item. Vedadas exigências desproporcionais.",
    )
    p(
        doc,
        "5.4 Garantia: no mínimo 90 (noventa) dias contra defeitos de fabricação para materiais de consumo duráveis e de 12 (doze) meses para equipamentos/eletroportáteis (itens 04, 10, 11, 12 e 15), contados do recebimento definitivo, sem prejuízo do Código de Defesa do Consumidor.",
    )

    h2(doc, "6 DA EXECUÇÃO DO OBJETO")
    p(
        doc,
        "6.1 Local de entrega: Almoxarifado da Secretaria Municipal de Saúde de Pelotas/RS e/ou outro endereço indicado na ordem de fornecimento pela Diretoria de Atenção Primária / PICSPel, em dias úteis, das 8h às 17h, mediante agendamento prévio de no mínimo 24 horas.",
    )
    p(
        doc,
        "6.2 A contratada deverá discriminar na nota fiscal o número do empenho, o item, a quantidade e a descrição compatível com o edital.",
    )
    p(
        doc,
        "6.3 Após o recebimento institucional, a SMS promoverá a distribuição interna às equipes eMulti e UBS, com registro de controle, conforme reunião de organização prevista no planejamento do PICSPel.",
    )

    h2(doc, "7 DA GESTÃO DA CONTRATAÇÃO")
    p(doc, f"Gestor: {GESTOR[0]} — Matrícula {GESTOR[1]}.")
    p(doc, f"Fiscal: {FISCAL[0]} — Matrícula {FISCAL[1]}.")
    p(doc, "A gestão e a fiscalização observarão a Lei nº 14.133/2021 e a regulamentação municipal vigente.")

    h2(doc, "8 RECEBIMENTO DO OBJETO")
    p(
        doc,
        "8.1 Recebimento provisório: no ato da entrega, mediante verificação de quantidades e integridade das embalagens.",
    )
    p(
        doc,
        "8.2 Recebimento definitivo: em até 5 (cinco) dias úteis após o provisório, mediante verificação de conformidade com as especificações. Produtos em desacordo serão recusados e deverão ser substituídos em até 5 (cinco) dias úteis, sem ônus para a Administração.",
    )

    h2(doc, "9 DOS CRITÉRIOS DE MEDIÇÃO E PAGAMENTO")
    p(
        doc,
        "O pagamento será efetuado após o recebimento definitivo e o atesto da nota fiscal, no prazo e condições da legislação municipal e do edital, por item efetivamente entregue e aceito. Não haverá pagamento por estimativa.",
    )

    h2(doc, "10 DA FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR")
    p(
        doc,
        "Pregão Eletrônico — menor preço por item. Poderá haver adjudicação a múltiplos fornecedores, conforme o resultado por item. Aplicam-se as regras de desempate e sanções da Lei nº 14.133/2021 e do edital.",
    )

    h2(doc, "11 DAS OBRIGAÇÕES DO CONTRATANTE")
    bullet_line(doc, "emitir empenho/ordem de fornecimento e indicar local e horário de entrega;")
    bullet_line(doc, "receber, conferir e atestar as notas fiscais;")
    bullet_line(doc, "efetuar o pagamento nas condições pactuadas;")
    bullet_line(doc, "fiscalizar a execução e notificar irregularidades;")
    bullet_line(doc, "promover a distribuição interna e o controle patrimonial quando cabível.")

    h2(doc, "12 DAS OBRIGAÇÕES DO CONTRATADO")
    bullet_line(doc, "fornecer os itens conforme especificações, prazos e preços adjudicados;")
    bullet_line(doc, "substituir produtos desconformes no prazo estipulado;")
    bullet_line(doc, "manter, durante a execução, as condições de habilitação;")
    bullet_line(doc, "responsabilizar-se por transporte, carga, descarga e encargos trabalhistas, previdenciários, fiscais e comerciais;")
    bullet_line(doc, "cumprir as normas de segurança e higiene no ato da entrega.")

    h2(doc, "13 DAS SANÇÕES")
    p(
        doc,
        "Pelo descumprimento total ou parcial das obrigações, aplicam-se as sanções dos arts. 155 a 163 da Lei nº 14.133/2021 e a regulamentação municipal (incluindo o Decreto Municipal nº 6.799/2023, no que couber), assegurados o contraditório e a ampla defesa.",
    )

    h2(doc, "14 DA ADEQUAÇÃO ORÇAMENTÁRIA")
    p(doc, "Emenda Parlamentar nº " + EMENDA + ".")
    p(doc, "Valor estimado: " + brl(VALOR_TOTAL) + ".")
    p(doc, "Projeto/Atividade: [informar código — Atenção Primária / PICSPel].")
    p(doc, "Fonte de recurso: [informar — vinculada à emenda parlamentar].")
    p(doc, "Classificação da despesa: 3.3.90.30 (material de consumo) e 4.4.90.52 (equipamentos e material permanente), conforme validação da Contabilidade.")
    p(doc, "Elementos a confirmar com a Contabilidade/CCGOF antes do empenho.")

    assinaturas_equipe(doc)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    print(f"OK {path}")


# ---------------------------------------------------------------------------
# 3) DFD
# ---------------------------------------------------------------------------
def gerar_dfd(path: Path):
    doc = new_doc()
    title(doc, "DOCUMENTO DE FORMALIZAÇÃO DA DEMANDA nº [___]/2026 – PICSPel (MATERIAIS)")

    p(doc, "UNIDADE DEMANDANTE: Secretaria Municipal de Saúde – SMS / Diretoria de Atenção Primária / Programa PICSPel.", bold=True)

    h2(doc, "1 OBJETO")
    p(
        doc,
        "Aquisição de materiais destinados às Práticas Integrativas e Complementares em Saúde (PICSPel), compreendendo 16 itens (bacias e protetores para escalda-pés, insumos de auriculoterapia e aromaterapia, equipamentos de apoio a práticas coletivas e materiais educativos), para abastecimento de 08 equipes eMulti, UBS da rede e suporte a cursos com 50 vagas, com recursos da Emenda Parlamentar nº "
        + EMENDA
        + ".",
    )

    h2(doc, "1.1 TIPO DE OBJETO")
    checkbox(doc, "Material de Consumo", checked=True)
    checkbox(doc, "Material Permanente / Equipamento", checked=True)
    checkbox(doc, "Serviço continuado com dedicação exclusiva de mão de obra")
    checkbox(doc, "Serviço continuado sem dedicação exclusiva de mão de obra")
    checkbox(doc, "Serviço não continuado")
    p(
        doc,
        "Observação: o objeto é predominantemente material de consumo, com itens de natureza permanente de apoio (caixa de som, garrafas térmicas, maletas, jarra elétrica e plastificadora). A parcela de serviços de capacitação da mesma emenda ("
        + brl(RECURSO_SERVICOS)
        + ") será formalizada em DFD/processo próprio.",
    )

    h2(doc, "2 FORMA DE CONTRATAÇÃO SUGERIDA")
    checkbox(doc, "Pregão Eletrônico", checked=True)
    checkbox(doc, "Pregão Eletrônico com Sistema/Ata de Registro de Preços")
    checkbox(doc, "Dispensa com Disputa Eletrônica")
    checkbox(doc, "Adesão à Ata do Consórcio Público do Extremo Sul")
    checkbox(doc, "Adesão à Ata de Registro de Preços de Outro(s) Órgão(s)")
    p(doc, "Critério de julgamento sugerido: menor preço por item.")

    h2(doc, "3 JUSTIFICATIVA DA CONTRATAÇÃO")
    p(
        doc,
        "A formalização desta demanda destina-se a viabilizar a oferta de Práticas Integrativas e Complementares na Atenção Primária, em conformidade com a PNPIC e com o planejamento do PICSPel, utilizando recursos da Emenda Parlamentar nº "
        + EMENDA
        + " (cota de "
        + brl(RECURSO_CONSUMO)
        + " para materiais). Os insumos são indispensáveis às práticas de auriculoterapia, aromaterapia, escalda-pés e práticas corporais, bem como à realização de cursos de formação (50 vagas). A aquisição por pregão eletrônico com julgamento por item assegura competitividade, publicidade e economicidade.",
    )

    h2(doc, "4 DESCRIÇÃO DO OBJETO")
    rows = []
    for n, desc, unid, qtd, vu, vt, nat, dest in ITENS:
        short = desc if len(desc) <= 70 else desc[:67] + "..."
        rows.append([f"{n:02d}", short, unid, str(qtd)])
    table(doc, ["Item", "Especificação", "Unidade", "Qtd."], rows)
    p(doc, "Valor total estimado: " + brl(VALOR_TOTAL) + " (referência da Lista de Materiais PICSPel; a confirmar por pesquisa de preços).")

    h2(doc, "5 DOTAÇÃO ORÇAMENTÁRIA")
    table(
        doc,
        ["Campo", "Informação"],
        [
            ["Emenda Parlamentar", EMENDA],
            ["Projeto/Atividade", "[informar — Atenção Primária / PICSPel]"],
            ["Fonte de recurso", "[informar — vinculada à emenda]"],
            ["Classificação da despesa", "3.3.90.30 e 4.4.90.52 (validar Contabilidade)"],
            ["Valor estimado", brl(VALOR_TOTAL)],
            ["Teto da cota de materiais na emenda", brl(RECURSO_CONSUMO)],
        ],
    )

    h2(doc, "6 LOCAL DE ENTREGA")
    p(
        doc,
        "Almoxarifado da Secretaria Municipal de Saúde e/ou local indicado pela Diretoria de Atenção Primária / coordenação PICSPel, com posterior distribuição às equipes eMulti (Centro/Porto, Fragata 1 e 2, Três Vendas 1 e 2, Areal, Colônia, Praias) e UBS de referência.",
    )

    h2(doc, "7 PREVISÃO DE INÍCIO")
    p(
        doc,
        "Entrega em até 30 (trinta) dias corridos da Nota de Empenho / ordem de fornecimento. Após a chegada, reunião com eMulti e APS de referência para organização do armazenamento e controle.",
    )

    assinaturas_equipe(doc)
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    print(f"OK {path}")


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    gerar_etp(OUT / "PICSPel_ETP_Materiais.docx")
    gerar_tr(OUT / "PICSPel_TR_Materiais.docx")
    gerar_dfd(OUT / "PICSPel_DFD_Materiais.docx")
    readme = OUT / "README.md"
    readme.write_text(
        """# PICSPel — Aquisição de Materiais (Emenda 43530013)

Pacote documental no padrão SMS Pelotas / RP para aquisição dos materiais das Práticas Integrativas e Complementares em Saúde.

## Peças

| Ordem | Arquivo | Conteúdo |
|------:|---------|----------|
| 1 | `PICSPel_ETP_Materiais.docx` | Estudo Técnico Preliminar |
| 2 | `PICSPel_TR_Materiais.docx` | Termo de Referência |
| 3 | `PICSPel_DFD_Materiais.docx` | Documento de Formalização da Demanda |

## Parâmetros

- **Recurso:** Emenda Parlamentar nº 43530013 (R$ 100.000,00)
- **Materiais (este processo):** estimativa R$ 58.000,00 (cota de consumo R$ 60.000,00)
- **Serviços/cursos (processo próprio):** R$ 40.000,00
- **Modalidade sugerida:** Pregão Eletrônico — menor preço por item
- **Abrangência:** 8 equipes eMulti + UBS + suporte a cursos (50 vagas)

## Regenerar

```bash
pip install python-docx
python3 scripts/gerar_picspel.py
```

## Pendências antes de protocolar

- Número do memorando / DFD / processo
- Dotação, fonte e classificação (Contabilidade)
- Mapa de pesquisa de preços (art. 23)
- Confirmação de PCA
- Tombamento dos itens permanentes
""",
        encoding="utf-8",
    )
    print(f"OK {readme}")


if __name__ == "__main__":
    main()
