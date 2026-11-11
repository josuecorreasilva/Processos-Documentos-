#!/usr/bin/env python3
"""Gera os Processos O2 (Dispensa + Pregão) no padrão documental dos RP02–RP13."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "docs" / "04-PROCESSOS-O2"
P1 = OUT / "O2_P1_Dispensa_Emergencial"
P2 = OUT / "O2_P2_Pregao_Paralelo"
BRASAO = ROOT / "docs" / "assets" / "brasao_rp_header.png"
if not BRASAO.exists():
    BRASAO = ROOT / "docs" / "assets" / "brasao_pelotas.png"

GESTOR = ("Dariane Lima Portela", "44151", "Gestor")
FISCAL = ("Josué Corrêa da Silva", "43351", "Fiscal")

D = {
    "processo_origem": "1242/2026",
    "sl": "00710/2026",
    "memo": "003746/2026",
    "memo_ccgof": "005720/2026",
    "data_deserto": "29/10/2026",
    "prazo": "180 (cento e oitenta) dias",
    "prazo_curto": "180 dias",
    "projeto": "10.301.0105.2167.00 — Atenção Primária em Saúde",
    "fonte": "1600.000",
    "consumo_mes": "206,0 m³",
    "consumo_180": "1.236,0 m³",
    "consumo_ano": "2.472,0 m³",
    "urbano_mes": "182,0 m³",
    "rural_mes": "24,0 m³",
    "urbano_ano": "2.184,0 m³",
    "rural_ano": "288,0 m³",
}


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
    """Insere 'Página X de Y' no rodapé."""
    def _fld(instr):
        run = paragraph.add_run()
        _font(run, size=9)
        r = run._r
        fc1 = OxmlElement("w:fldChar"); fc1.set(qn("w:fldCharType"), "begin"); r.append(fc1)
        run2 = paragraph.add_run()
        _font(run2, size=9)
        it = OxmlElement("w:instrText"); it.set(qn("xml:space"), "preserve"); it.text = instr
        run2._r.append(it)
        run3 = paragraph.add_run()
        _font(run3, size=9)
        fc2 = OxmlElement("w:fldChar"); fc2.set(qn("w:fldCharType"), "end"); run3._r.append(fc2)

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
    # clear default
    header.paragraphs[0].clear()
    hp = header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    hp.paragraph_format.space_after = Pt(0)
    if BRASAO.exists():
        run = hp.add_run()
        run.add_picture(str(BRASAO), width=Cm(1.7))
    # blank lines similar to RP spacing
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
    # support "Label: rest" bold label
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
    r = para.add_run(text)
    _font(r)
    return para


def table(doc, headers, rows, *, header_shade=True):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = t.rows[0].cells[i]
        cell.text = ""
        r = cell.paragraphs[0].add_run(h)
        _font(r, size=9, bold=True)
        if header_shade:
            _shade(cell)
    for ri, row in enumerate(rows, start=1):
        for ci, val in enumerate(row):
            cell = t.rows[ri].cells[ci]
            cell.text = ""
            r = cell.paragraphs[0].add_run(str(val))
            _font(r, size=9)
    doc.add_paragraph()
    return t


def assinaturas_equipe(doc):
    h2(doc, "EQUIPE DE PLANEJAMENTO / FISCALIZAÇÃO")
    t = doc.add_table(rows=1, cols=2)
    t.style = "Table Grid"
    for i, pessoa in enumerate((GESTOR, FISCAL)):
        cell = t.rows[0].cells[i]
        cell.text = ""
        lines = [
            "__________________________",
            pessoa[0],
            f"Matrícula: {pessoa[1]}",
            pessoa[2],
        ]
        first = True
        for line in lines:
            para = cell.paragraphs[0] if first else cell.add_paragraph()
            first = False
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = para.add_run(line)
            _font(r, size=10)
    doc.add_paragraph()
    p(doc, "Pelotas/RS, _____ de ______________ de 2026.", justify=False)


def save(doc, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    print("OK", path.relative_to(ROOT))


# ---------------------------------------------------------------------------
# VOLUME 0
# ---------------------------------------------------------------------------

def o2_indice():
    doc = new_doc()
    title(doc, "ÍNDICE DOS AUTOS — PROCESSOS O2 (OXIGÊNIO MEDICINAL)")
    p(doc, "Pacote elaborado no padrão documental dos Registros de Preços RP02 a RP13 "
           "(DFD / ETP / TR), com cabeçalho oficial da SMS e equipe de planejamento padronizada. "
           "Os dois processos tramitam simultaneamente.")
    h2(doc, "1 PROCESSO 1 — DISPENSA EMERGENCIAL (Art. 75, VIII)")
    table(doc, ["Peça", "Arquivo", "Conteúdo"], [
        ["1.1", "O2_P1_DFD.docx", "Documento de Formalização da Demanda emergencial"],
        ["1.2", "O2_P1_ETP.docx", "Estudo Técnico Preliminar simplificado/completo para emergência"],
        ["1.3", "O2_P1_TR.docx", "Termo de Referência emergencial"],
        ["1.4", "O2_P1_Justificativa_Emergencia.docx", "Justificativa da emergência e enquadramento legal"],
        ["1.5", "O2_P1_Mapa_Precos.docx", "Mapa de pesquisa de preços"],
        ["1.6", "O2_P1_Razao_Escolha.docx", "Razão da escolha do fornecedor"],
        ["1.7", "O2_P1_Minuta_Contrato.docx", "Minuta do contrato emergencial"],
        ["1.8", "O2_P1_Despacho_Autorizacao.docx", "Despacho de autorização e extrato PNCP"],
    ])
    h2(doc, "2 PROCESSO 2 — PREGÃO ELETRÔNICO EM PARALELO")
    table(doc, ["Peça", "Arquivo", "Conteúdo"], [
        ["2.1", "O2_P2_DFD.docx", "DFD revisado para novo pregão"],
        ["2.2", "O2_P2_ETP.docx", "ETP completo"],
        ["2.3", "O2_P2_TR.docx", "Termo de Referência revisado"],
        ["2.4", "O2_P2_Pesquisa_Precos.docx", "Pesquisa de preços atualizada"],
        ["2.5", "O2_P2_Minuta_Edital.docx", "Minuta de edital"],
        ["2.6", "O2_P2_Minuta_Contrato.docx", "Minuta de contrato definitivo"],
        ["2.7", "O2_P2_Oficio_Divulgacao.docx", "Ofício de divulgação ampliada"],
    ])
    h2(doc, "3 PEÇAS COMUNS")
    table(doc, ["Peça", "Arquivo", "Conteúdo"], [
        ["0.1", "O2_Despacho_Encaminhamento_Paralelo.docx", "Adoção simultânea das Opções A e C"],
        ["0.2", "O2_Cronograma_Paralelo.docx", "Cronograma semanal integrado"],
        ["0.3", "O2_Indice.docx", "Este índice"],
    ])
    assinaturas_equipe(doc)
    save(doc, OUT / "O2_Indice.docx")


def o2_despacho():
    doc = new_doc()
    title(doc, "DESPACHO DE ENCAMINHAMENTO — PROCESSOS O2 EM PARALELO")
    p(doc, f"A/C Setor de Compras e Setor de Licitações da SMS — Processo nº {D['processo_origem']}.")
    h2(doc, "1 DOS FATOS")
    p(doc, f"O Processo nº {D['processo_origem']} (SL nº {D['sl']}) restou deserto, conforme comunicação "
           f"de {D['data_deserto']} (Memorando nº {D['memo']}; Memorando CCGOF nº {D['memo_ccgof']}).")
    h2(doc, "2 DA DECISÃO")
    p(doc, "Adotar SIMULTANEAMENTE:")
    bullet_line(doc, "Opção C — Dispensa emergencial (Art. 75, VIII, Lei nº 14.133/2021), prazo "
                     f"{D['prazo_curto']} ou até homologação do pregão;")
    bullet_line(doc, "Opção A — Novo Pregão Eletrônico, com DFD/ETP/TR revisados, pesquisa atualizada "
                     "e divulgação ampliada.")
    h2(doc, "3 DOS DOCUMENTOS")
    p(doc, "Seguem, no padrão documental dos RP (DFD, ETP e TR), as peças dos Volumes 1 e 2, "
           "para tramitação prioritária e concomitante.")
    assinaturas_equipe(doc)
    save(doc, OUT / "O2_Despacho_Encaminhamento_Paralelo.docx")


def o2_cronograma():
    doc = new_doc()
    title(doc, "CRONOGRAMA PARALELO — PROCESSOS O2")
    table(doc,
          ["Semana", "Processo 1 — Dispensa", "Processo 2 — Pregão"],
          [
              ["1", "Cotações + DFD/ETP/TR + justificativa + jurídico", "Revisão DFD/ETP/TR + nova pesquisa"],
              ["2", "Autorização + PNCP + assinatura + início", "Publicação do edital + divulgação ampliada"],
              ["3", "Fornecimento + fiscalização", "Sessão pública / propostas"],
              ["4", "Medições", "Habilitação / homologação"],
              ["Após homologar", "Rescisão automática com transição", "Assinatura do contrato definitivo"],
          ])
    assinaturas_equipe(doc)
    save(doc, OUT / "O2_Cronograma_Paralelo.docx")


# ---------------------------------------------------------------------------
# PROCESSO 1
# ---------------------------------------------------------------------------

def p1_dfd():
    doc = new_doc()
    title(doc, f"DOCUMENTO DE FORMALIZAÇÃO DA DEMANDA nº [___]/2026 – O2 PROCESSO 1")
    p(doc, "UNIDADE DEMANDANTE: Secretaria Municipal de Saúde – SMS Diretoria de Atenção Primária.", bold=True)

    h2(doc, "1 OBJETO")
    p(doc, "Objeto: Contratação emergencial de empresa especializada no fornecimento contínuo, sob demanda, "
           "de Oxigênio Medicinal gasoso, com cessão de cilindros em regime de comodato, manutenção e logística "
           "de entrega, para UBSs e UBAIs das zonas urbana e rural de Pelotas/RS, pelo prazo estritamente "
           f"necessário de {D['prazo']}, ou até a conclusão da licitação regular instaurada em paralelo, "
           "o que ocorrer primeiro.", bold=True)
    p(doc, "O objeto compreende o abastecimento por metro cúbico (m³) efetivamente consumido, sem obrigação "
           "de consumo mínimo, com responsabilidade integral da contratada pela logística, troca de cilindros, "
           "manutenção, testes hidrostáticos e conformidade ANVISA/ABNT.")

    h2(doc, "1.1 TIPO DE OBJETO")
    for line in [
        "(   ) Material de Consumo",
        "(   ) Material Permanente/ Equipamento",
        "(   ) Serviço continuado com dedicação exclusiva de mão de obra",
        "( X ) Serviço continuado sem dedicação exclusiva de mão de obra",
        "(   ) Serviço não continuado",
    ]:
        bullet_line(doc, line)

    h2(doc, "2 FORMA DE CONTRATAÇÃO SUGERIDA")
    for line in [
        "(   ) Pregão Eletrônico",
        "(   ) Pregão Eletrônico com Sistema/Ata de Registro de Preços",
        "(   ) Dispensa com Disputa Eletrônica",
        "( X ) Dispensa de Licitação — Art. 75, VIII, Lei nº 14.133/2021 (emergencial)",
        "(   ) Adesão à Ata do Consórcio Público do Extremo Sul",
        "(   ) Adesão à Ata de Registro de Preços de Outro(s) Órgão(s)",
    ]:
        bullet_line(doc, line)
    p(doc, "Critério de escolha do fornecedor: menor preço compatível com pesquisa de mercado e capacidade "
           "de atendimento imediato (24h / 12h em urgência).")

    h2(doc, "3 JUSTIFICATIVA DA CONTRATAÇÃO")
    p(doc, f"A licitação do Processo nº {D['processo_origem']} restou deserta em {D['data_deserto']}. "
           "O oxigênio medicinal é insumo essencial ao suporte à vida nas UBSs e UBAIs. A tramitação média "
           "de novo certame (30 a 45 dias) é incompatível com o estoque disponível. A situação não decorre "
           "de falta de planejamento, mas de ausência de interessados no certame regular.")
    p(doc, "Paralelamente, instaurar-se-á o Processo 2 (Pregão Eletrônico revisado), para solução definitiva, "
           "com rescisão automática desta contratação emergencial quando homologado o pregão.")

    h2(doc, "4 DESCRIÇÃO DO OBJETO")
    p(doc, "Quantitativos referenciais para o período emergencial:")
    table(doc, ["Item", "Especificação", "Unidade", "Qtd./mês", "Qtd./180 dias"], [
        ["01", "Oxigênio medicinal gasoso — Zona Urbana (UBSs + UBAIs)", "m³", "182,0", "1.092,0"],
        ["02", "Oxigênio medicinal gasoso — Zona Rural (12 UBSs)", "m³", "24,0", "144,0"],
        ["", "TOTAL ESTIMADO", "m³", "206,0", "1.236,0"],
    ])
    p(doc, "Especificação mínima: pureza ≥ 99,5%; cilindros em comodato (7 m³ / 10 m³ ou equivalente); "
           "entrega em até 24 horas (12 horas em urgência clínica); assistência 24h; laudo de pureza quando solicitado.")

    h2(doc, "5 DOTAÇÃO ORÇAMENTÁRIA")
    table(doc, ["Campo", "Informação"], [
        ["Projeto/Atividade", D["projeto"]],
        ["Fonte de recurso", D["fonte"]],
        ["Classificação da despesa", "[3.3.90.__.__]"],
        ["Valor estimado", "R$ [________] (mapa de preços)"],
    ])

    h2(doc, "6 LOCAL DE ENTREGA")
    p(doc, "Entrega direta nas UBSs e UBAIs das zonas urbana e rural, conforme relação de endereços anexa "
           "aos autos (41 pontos urbanos e 12 rurais), mediante solicitação das unidades/Diretoria.")

    h2(doc, "7 PREVISÃO DE INÍCIO")
    p(doc, "Imediato após assinatura do contrato / Nota de Empenho, com primeira entrega em até 24 horas.")
    assinaturas_equipe(doc)
    save(doc, P1 / "O2_P1_DFD.docx")


def p1_etp():
    doc = new_doc()
    title(doc, "ESTUDO TÉCNICO PRELIMINAR – O2 PROCESSO 1 (DISPENSA EMERGENCIAL)")
    h2(doc, "1 IDENTIFICAÇÃO DO PROCESSO")
    p(doc, "1.1 Unidade responsável pela contratação: Secretaria Municipal de Saúde – Diretoria de Atenção Primária")
    p(doc, "1.2 Objeto da contratação: fornecimento emergencial de oxigênio medicinal gasoso com comodato de cilindros "
           "para UBSs e UBAIs (zonas urbana e rural).")
    p(doc, f"1.3 Número do processo: referência Processo nº {D['processo_origem']} (licitação deserta) — "
           "DFD nº [___]/2026 – O2 Processo 1")

    h2(doc, "2 DESCRIÇÃO DA NECESSIDADE")
    p(doc, "Assegurar abastecimento ininterrupto de oxigênio medicinal após a desertificação do certame regular, "
           "evitando risco à vida e interrupção de procedimentos assistenciais (inalações, nebulizações e suporte clínico).")

    h2(doc, "3 DA CONTRATAÇÃO E DO PLANEJAMENTO")
    p(doc, "Forma: Dispensa de licitação, Art. 75, VIII, Lei nº 14.133/2021, pelo prazo estritamente necessário "
           f"({D['prazo_curto']}), em paralelo a novo Pregão Eletrônico.")
    p(doc, "Previsão no Plano de Contratações Anual: [informar item do PCA]. A emergência resulta do resultado "
           "do certame (deserto), não da ausência de planejamento da demanda.")

    h2(doc, "4 REQUISITOS DA CONTRATAÇÃO")
    for item in [
        "Oxigênio medicinal gasoso, pureza mínima 99,5% (v/v);",
        "Cilindros em comodato, com teste hidrostático válido e válvulas em bom estado;",
        "Entrega em até 24h (12h em urgência); assistência 24h;",
        "Conformidade ANVISA/ABNT; troca imediata de recipiente irregular;",
        "Medição e faturamento exclusivamente pelo volume entregue (m³).",
    ]:
        bullet_line(doc, item)

    h2(doc, "5 ESTIMATIVA DAS QUANTIDADES")
    table(doc, ["Área", "Consumo médio mensal", "Estimativa 180 dias", "Consumo histórico"], [
        ["Zona Urbana", "182,0 m³", "1.092,0 m³", "[____]"],
        ["Zona Rural", "24,0 m³", "144,0 m³", "[____]"],
        ["Total", "206,0 m³", "1.236,0 m³", "[____]"],
    ])

    h2(doc, "6 LEVANTAMENTO DE MERCADO")
    p(doc, "Cotações diretas a empresas de gases medicinais da região, painéis oficiais e contratações similares. "
           "Alternativas avaliadas: (i) repetição isolada do pregão; (ii) dispensa emergencial; (iii) ambas em paralelo. "
           "Concluiu-se pela opção (iii).")

    h2(doc, "7 ESTIMATIVA DO VALOR DA CONTRATAÇÃO")
    table(doc, ["Item", "Descrição", "Unid.", "Qtd. 180 dias", "Valor unit. (R$)", "Valor total (R$)"], [
        ["01", "Oxigênio medicinal — Urbano", "m³", "1.092,0", "[____]", "[____]"],
        ["02", "Oxigênio medicinal — Rural", "m³", "144,0", "[____]", "[____]"],
        ["", "TOTAL", "", "1.236,0", "", "R$ [________]"],
    ])

    h2(doc, "8 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO")
    p(doc, "Fornecimento sob demanda com comodato e logística inclusos no preço do m³, vigência limitada à emergência, "
           "com transição para o contrato do pregão paralelo.")

    h2(doc, "9 JUSTIFICATIVA PARA PARCELAMENTO")
    p(doc, "Na emergência, admite-se contratação única abrangendo zonas urbana e rural para garantir cobertura imediata. "
           "No pregão paralelo, mantém-se a divisão em lotes.")

    h2(doc, "10 DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS")
    p(doc, "Continuidade assistencial, ausência de desabastecimento e contratação definitiva via pregão.")

    h2(doc, "11 PROVIDÊNCIAS PRÉVIAS AO CONTRATO")
    p(doc, "Relatório de estoque; cotações; indicação de gestor/fiscal; parecer jurídico; publicação no PNCP.")

    h2(doc, "12 CONTRATAÇÕES CORRELATAS/INTERDEPENDENTES")
    p(doc, "Processo 2 — Pregão Eletrônico em paralelo, com rescisão automática desta dispensa na homologação.")

    h2(doc, "13 POSSÍVEIS IMPACTOS AMBIENTAIS")
    p(doc, "Baixo impacto; responsabilidade da contratada sobre cilindros, transporte e destinação conforme normas.")

    h2(doc, "14 ANÁLISE DE RISCOS")
    table(doc, ["Risco", "Prob.", "Impacto", "Tratamento"], [
        ["Desabastecimento antes da assinatura", "Alta", "Alto", "Priorizar cotações/autorização; entrega 24h/12h"],
        ["Preço acima do mercado", "Média", "Médio", "Mapa com ≥3 fontes; justificativa formal"],
        ["Falha logística rural", "Média", "Alto", "Exigir cobertura territorial na proposta"],
        ["Atraso do pregão paralelo", "Média", "Médio", "Cronograma integrado; vedar prorrogação indevida"],
    ])

    h2(doc, "15 DECLARAÇÃO DE VIABILIDADE DA CONTRATAÇÃO")
    p(doc, "Declaramos a viabilidade técnica, administrativa e orçamentária da contratação emergencial proposta.")
    assinaturas_equipe(doc)
    save(doc, P1 / "O2_P1_ETP.docx")


def p1_tr():
    doc = new_doc()
    title(doc, "TERMO DE REFERÊNCIA – O2 PROCESSO 1 (DISPENSA EMERGENCIAL)")
    h2(doc, "1 DO OBJETO, VALOR E CONDIÇÕES GERAIS DA CONTRATAÇÃO")
    p(doc, "1.1 O presente Termo de Referência tem por objeto a contratação emergencial de fornecimento contínuo, "
           "sob demanda, de oxigênio medicinal gasoso, com comodato de cilindros, para UBSs e UBAIs de Pelotas/RS.")
    p(doc, "1.2 Os objetos deste instrumento não se enquadram como bens de luxo.")
    p(doc, "1.3 O objeto é caracterizado como comum, com padrões de desempenho e qualidade objetivamente definidos.")
    p(doc, f"1.4 O prazo de vigência será de {D['prazo']}, ou até a homologação/assinatura do contrato do Pregão "
           "paralelo, o que ocorrer primeiro, sem prorrogação ordinária.")
    p(doc, "1.5 A entrega será parcelada sob demanda, mediante solicitação das unidades/Diretoria.")
    p(doc, "1.6 O valor total estimado é de R$ [________], conforme mapa de preços.")

    h2(doc, "2 DA FUNDAMENTAÇÃO DA CONTRATAÇÃO")
    p(doc, f"2.1 Fundamentação: Art. 75, VIII, Lei nº 14.133/2021; licitação deserta do Processo nº {D['processo_origem']}.")
    p(doc, "2.2 Justificativa: risco de desabastecimento de oxigênio medicinal e prejuízo à continuidade do serviço de saúde.")

    h2(doc, "3 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO CONSIDERADO O CICLO DE VIDA DO OBJETO")
    p(doc, "Fornecimento, comodato, manutenção, logística, medição mensal e transição para o contrato definitivo.")

    h2(doc, "4 GESTÃO/ANÁLISE DE RISCOS")
    p(doc, "Principais riscos: atraso de entrega, irregularidade sanitária e falha na transição contratual. "
           "Tratamento: prazos rígidos, fiscalização, cláusulas de sanção e transição de até 10 dias úteis.")

    h2(doc, "5 REQUISITOS DA CONTRATAÇÃO")
    p(doc, "5.1 Oxigênio medicinal gasoso, pureza ≥ 99,5%; cilindros em comodato; ANVISA/ABNT.")
    p(doc, "5.2 Entrega em até 24 horas; 12 horas em urgência; assistência 24h; troca imediata de cilindro irregular.")
    p(doc, "5.3 Regularização sanitária: registro/autorização ANVISA do produto e licença sanitária da empresa.")

    h2(doc, "6 DA EXECUÇÃO DO OBJETO")
    table(doc, ["Item", "Descrição", "Unid.", "Qtd. 180 dias", "Valor unit. (R$)", "Valor total (R$)"], [
        ["01", "Oxigênio medicinal gasoso — Zona Urbana", "m³", "1.092,0", "[____]", "[____]"],
        ["02", "Oxigênio medicinal gasoso — Zona Rural", "m³", "144,0", "[____]", "[____]"],
    ])
    p(doc, "Locais: UBSs/UBAIs urbanas e rurais. Não há consumo mínimo.")

    h2(doc, "7 DA GESTÃO DA CONTRATAÇÃO")
    p(doc, f"Gestor: {GESTOR[0]} (Matrícula {GESTOR[1]}). Fiscal: {FISCAL[0]} (Matrícula {FISCAL[1]}).")

    h2(doc, "8 RECEBIMENTO DO OBJETO")
    p(doc, "Recebimento provisório na entrega e definitivo após verificação de quantidade, integridade e conformidade.")

    h2(doc, "9 DOS CRITÉRIOS DE MEDIÇÃO E PAGAMENTO")
    p(doc, "Medição mensal do volume efetivamente fornecido (m³), com NF e atesto. Vedado pagamento por volume não entregue.")

    h2(doc, "10 DA FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR")
    p(doc, "Contratação direta emergencial, com pesquisa de preços e justificativa do fornecedor escolhido.")

    h2(doc, "11 DAS OBRIGAÇÕES DO CONTRATANTE")
    p(doc, "Pagar as medições; indicar fiscal/gestor; assegurar acesso às unidades; comunicar demandas.")

    h2(doc, "12 DAS OBRIGAÇÕES DO CONTRATADO")
    p(doc, "Cumprir o TR; manter habilitação; estoque operacional; comodato/manutenção; prazos; documentação sanitária.")

    h2(doc, "13 DA ADEQUAÇÃO ORÇAMENTÁRIA")
    p(doc, f"Dotação: {D['projeto']} | Fonte: {D['fonte']} | Empenho nº [________].")

    h2(doc, "14 DA RESCISÃO AUTOMÁTICA")
    p(doc, "O contrato extingue-se automaticamente com a homologação do Pregão Eletrônico paralelo e assinatura "
           "do contrato definitivo, assegurada transição de até 10 (dez) dias úteis sem desabastecimento.")
    assinaturas_equipe(doc)
    save(doc, P1 / "O2_P1_TR.docx")


def p1_justificativa():
    doc = new_doc()
    title(doc, "JUSTIFICATIVA DA EMERGÊNCIA E DO ENQUADRAMENTO LEGAL – O2 PROCESSO 1")
    h2(doc, "1 DOS FATOS")
    p(doc, f"A licitação do Processo nº {D['processo_origem']} (SL nº {D['sl']}) restou deserta em "
           f"{D['data_deserto']}, conforme Memorando nº {D['memo']} e Memorando CCGOF nº {D['memo_ccgof']}.")
    p(doc, "O estoque atual de oxigênio medicinal é de aproximadamente _____ (_____) dias, insuficiente para "
           "aguardar novo certame (30 a 45 dias).")
    h2(doc, "2 DO DIREITO")
    p(doc, "Art. 75, VIII, da Lei nº 14.133/2021 autoriza a dispensa em situações de emergência que possam "
           "comprometer a segurança de pessoas e a continuidade de serviços públicos essenciais.")
    h2(doc, "3 DO PRAZO E DA ESTRATÉGIA")
    p(doc, f"Prazo: {D['prazo']}, ou até a conclusão do Pregão paralelo. Adota-se simultaneamente a Opção A "
           "(repetição do certame com TR/pesquisa revisados).")
    h2(doc, "4 DO PREÇO")
    p(doc, "O valor será justificado por mapa de preços com fontes de mercado local/regional.")
    h2(doc, "5 CONCLUSÃO")
    p(doc, "Requer-se a autorização da dispensa emergencial e a tramitação prioritária do pregão em paralelo.")
    assinaturas_equipe(doc)
    save(doc, P1 / "O2_P1_Justificativa_Emergencia.docx")


def p1_mapa():
    doc = new_doc()
    title(doc, "MAPA DE PESQUISA DE PREÇOS – O2 PROCESSO 1")
    p(doc, "Pesquisa nos termos do art. 23 da Lei nº 14.133/2021. Mínimo de 3 fontes; número inferior exige justificativa.")
    table(doc, ["Fonte / Empresa", "CNPJ", "R$/m³ Urbano", "R$/m³ Rural", "Validade", "Obs."], [
        ["Cotação 1 — _____________", "", "", "", "", ""],
        ["Cotação 2 — _____________", "", "", "", "", ""],
        ["Cotação 3 — _____________", "", "", "", "", ""],
        ["Painel / contrato similar", "", "", "", "", ""],
    ])
    h2(doc, "CONSOLIDAÇÃO")
    p(doc, "Metodologia: ( ) média  ( ) mediana  ( ) menor preço compatível  ( ) outra: ______")
    p(doc, "Preço de referência Urbano: R$ ________ / m³")
    p(doc, "Preço de referência Rural: R$ ________ / m³")
    p(doc, "Valor estimado 180 dias: R$ _______________")
    assinaturas_equipe(doc)
    save(doc, P1 / "O2_P1_Mapa_Precos.docx")


def p1_razao():
    doc = new_doc()
    title(doc, "RAZÃO DA ESCOLHA DO FORNECEDOR E JUSTIFICATIVA DO PREÇO – O2 PROCESSO 1")
    p(doc, "Razão social: _________________________________ CNPJ: ____________________")
    h2(doc, "1 RAZÃO DA ESCOLHA")
    for line in [
        "(   ) Menor preço compatível com a pesquisa",
        "(   ) Única com entrega em 24h e estoque imediato",
        "(   ) Única com cobertura urbana/rural no prazo emergencial",
        "(   ) Outra: ________________________________",
    ]:
        bullet_line(doc, line)
    h2(doc, "2 COMPARATIVO")
    p(doc, "Preço ofertado: R$ ________ / m³ — ( ) abaixo / ( ) na média / ( ) até _____% acima da referência.")
    h2(doc, "3 CONCLUSÃO")
    p(doc, "Recomenda-se a contratação direta da empresa indicada, com fulcro no art. 75, VIII, Lei nº 14.133/2021.")
    assinaturas_equipe(doc)
    save(doc, P1 / "O2_P1_Razao_Escolha.docx")


def p1_contrato():
    doc = new_doc()
    title(doc, "MINUTA DE CONTRATO EMERGENCIAL Nº _____/2026 – O2 PROCESSO 1")
    p(doc, "Pelo presente instrumento, o MUNICÍPIO DE PELOTAS, CNPJ 87.455.531/0001-57, por intermédio da "
           "Secretaria Municipal de Saúde, e a empresa _____________, CNPJ _____________, celebram o presente "
           "contrato de fornecimento emergencial de oxigênio medicinal, mediante as cláusulas seguintes:")
    for t, b in [
        ("CLÁUSULA PRIMEIRA — DO OBJETO",
         "Fornecimento contínuo, sob demanda, de oxigênio medicinal gasoso (pureza mínima 99,5%), com comodato "
         "de cilindros, conforme Termo de Referência O2 Processo 1."),
        ("CLÁUSULA SEGUNDA — DA FUNDAMENTAÇÃO",
         f"Dispensa de licitação, art. 75, VIII, Lei nº 14.133/2021, em razão da licitação deserta do Processo nº {D['processo_origem']}."),
        ("CLÁUSULA TERCEIRA — DO PREÇO E PAGAMENTO",
         "Preço unitário de R$ ________ / m³. Pagamento mensal por medição atestada."),
        ("CLÁUSULA QUARTA — DA VIGÊNCIA",
         f"Vigência de {D['prazo']}, ou até a conclusão do pregão paralelo, o que ocorrer primeiro."),
        ("CLÁUSULA QUINTA — DA RESCISÃO AUTOMÁTICA",
         "Extinção automática na homologação do Pregão paralelo, com transição de até 10 dias úteis."),
        ("CLÁUSULA SEXTA — DAS OBRIGAÇÕES",
         "A CONTRATADA cumprirá o TR (prazos 24h/12h, assistência 24h, ANVISA). A CONTRATANTE pagará e fiscalizará."),
        ("CLÁUSULA SÉTIMA — DA FISCALIZAÇÃO",
         f"Gestor: {GESTOR[0]}. Fiscal: {FISCAL[0]}."),
        ("CLÁUSULA OITAVA — DAS SANÇÕES",
         "Arts. 155 a 163 da Lei nº 14.133/2021."),
        ("CLÁUSULA NONA — DO FORO",
         "Comarca de Pelotas/RS."),
    ]:
        h2(doc, t)
        p(doc, b)
    assinaturas_equipe(doc)
    save(doc, P1 / "O2_P1_Minuta_Contrato.docx")


def p1_autorizacao():
    doc = new_doc()
    title(doc, "DESPACHO DE AUTORIZAÇÃO DA DISPENSA EMERGENCIAL – O2 PROCESSO 1")
    p(doc, f"À vista do DFD, ETP, TR, justificativa, mapa de preços e minuta, referentes ao oxigênio medicinal "
           f"após a licitação deserta do Processo nº {D['processo_origem']},")
    p(doc, "AUTORIZO a dispensa de licitação (Art. 75, VIII, Lei nº 14.133/2021) e a contratação emergencial "
           f"pelo prazo de {D['prazo']}, ou até a conclusão do Pregão paralelo, no valor estimado de "
           "R$ _______________, com a empresa _____________ — CNPJ _____________.")
    h2(doc, "DETERMINAÇÕES")
    for item in [
        "Publicação no PNCP e Portal da Transparência;",
        "Emissão de empenho e assinatura do contrato;",
        "Início imediato do fornecimento sob fiscalização;",
        "Prosseguimento prioritário do Processo 2 (Pregão).",
    ]:
        bullet_line(doc, item)
    h2(doc, "EXTRATO PARA PNCP")
    p(doc, f"DISPENSA — Art. 75, VIII. Órgão: SMS Pelotas/RS. Processo: _____/2026 (origem {D['processo_origem']}). "
           f"Objeto: oxigênio medicinal com comodato. Vigência: {D['prazo_curto']} ou até pregão. "
           "Valor: R$ _________. Contratada: _________.")
    p(doc, "Angela Moreira Vitória — Secretária Municipal de Saúde", justify=False)
    assinaturas_equipe(doc)
    save(doc, P1 / "O2_P1_Despacho_Autorizacao.docx")


# ---------------------------------------------------------------------------
# PROCESSO 2
# ---------------------------------------------------------------------------

def p2_dfd():
    doc = new_doc()
    title(doc, "DOCUMENTO DE FORMALIZAÇÃO DA DEMANDA nº [___]/2026 – O2 PROCESSO 2")
    p(doc, "UNIDADE DEMANDANTE: Secretaria Municipal de Saúde – SMS Diretoria de Atenção Primária.", bold=True)
    h2(doc, "1 OBJETO")
    p(doc, "Objeto: Contratação de empresa(s) especializada(s) para fornecimento contínuo, sob demanda, de "
           "Oxigênio Medicinal gasoso, com comodato de cilindros, manutenção e logística, para UBSs e UBAIs "
           "das zonas urbana e rural de Pelotas/RS, em 2 (dois) lotes, pelo prazo de 12 (doze) meses, "
           "prorrogável nos termos da lei.", bold=True)
    h2(doc, "1.1 TIPO DE OBJETO")
    for line in [
        "(   ) Material de Consumo",
        "(   ) Material Permanente/ Equipamento",
        "(   ) Serviço continuado com dedicação exclusiva de mão de obra",
        "( X ) Serviço continuado sem dedicação exclusiva de mão de obra",
        "(   ) Serviço não continuado",
    ]:
        bullet_line(doc, line)
    h2(doc, "2 FORMA DE CONTRATAÇÃO SUGERIDA")
    for line in [
        "( X ) Pregão Eletrônico",
        "(   ) Pregão Eletrônico com Sistema/Ata de Registro de Preços",
        "(   ) Dispensa com Disputa Eletrônica",
        "(   ) Adesão à Ata do Consórcio Público do Extremo Sul",
        "(   ) Adesão à Ata de Registro de Preços de Outro(s) Órgão(s)",
    ]:
        bullet_line(doc, line)
    p(doc, "Critério de julgamento sugerido: menor preço por lote.")
    h2(doc, "3 JUSTIFICATIVA DA CONTRATAÇÃO")
    p(doc, f"Assegurar abastecimento contínuo de oxigênio medicinal. O certame do Processo nº {D['processo_origem']} "
           "restou deserto. Este DFD revisado corrige atratividade: pesquisa atualizada, prazos realistas, "
           "habilitação proporcional e divulgação ampliada. Mantém-se a divisão em dois lotes.")
    h2(doc, "4 DESCRIÇÃO DO OBJETO")
    table(doc, ["Item", "Especificação", "Unidade", "Qtd. anual"], [
        ["01 — Lote 1", "Oxigênio medicinal gasoso — Zona Urbana (38 UBSs + 3 UBAIs)", "m³", "2.184,0"],
        ["02 — Lote 2", "Oxigênio medicinal gasoso — Zona Rural (12 UBSs)", "m³", "288,0"],
    ])
    h2(doc, "5 DOTAÇÃO ORÇAMENTÁRIA")
    table(doc, ["Campo", "Informação"], [
        ["Projeto/Atividade", D["projeto"]],
        ["Fonte de recurso", D["fonte"]],
        ["Classificação da despesa", "[3.3.90.__.__]"],
        ["Valor estimado", "R$ [________]"],
    ])
    h2(doc, "6 LOCAL DE ENTREGA")
    p(doc, "UBSs/UBAIs urbanas e rurais, conforme anexo de endereços.")
    h2(doc, "7 PREVISÃO DE INÍCIO")
    p(doc, "Até 10 (dez) dias úteis da Nota de Empenho / assinatura do contrato.")
    assinaturas_equipe(doc)
    save(doc, P2 / "O2_P2_DFD.docx")


def p2_etp():
    doc = new_doc()
    title(doc, "ESTUDO TÉCNICO PRELIMINAR – O2 PROCESSO 2 (PREGÃO ELETRÔNICO)")
    secoes = [
        ("1 IDENTIFICAÇÃO DO PROCESSO",
         f"Unidade: SMS/DAP. Objeto: fornecimento contínuo de oxigênio medicinal com comodato. "
         f"Referência: Processo nº {D['processo_origem']} (deserto)."),
        ("2 DESCRIÇÃO DA NECESSIDADE",
         "Abastecimento contínuo nas UBSs/UBAIs após desertificação do certame anterior."),
        ("3 DA CONTRATAÇÃO E DO PLANEJAMENTO",
         "Pregão Eletrônico, menor preço por lote, Lei nº 14.133/2021. PCA: [informar]."),
        ("4 REQUISITOS DA CONTRATAÇÃO",
         "Pureza ≥ 99,5%; comodato; logística urbana/rural; ANVISA/ABNT; assistência 24h; medição por m³."),
        ("5 ESTIMATIVA DAS QUANTIDADES",
         f"Anual: {D['consumo_ano']} (Lote 1: {D['urbano_ano']}; Lote 2: {D['rural_ano']})."),
        ("6 LEVANTAMENTO DE MERCADO",
         "Nova pesquisa com painéis, municípios similares e cotações. Registrar defasagem do preço anterior."),
        ("7 ESTIMATIVA DO VALOR DA CONTRATAÇÃO",
         "R$ [________] — mapa atualizado."),
        ("8 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO",
         "Pregão em 2 lotes; contrato 12 meses prorrogável; pagamento por consumo; comodato incluso."),
        ("9 JUSTIFICATIVA PARA PARCELAMENTO",
         "Dois lotes ampliam competitividade e refletem logística rural."),
        ("10 DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS",
         "Segurança assistencial, economicidade e ampliação de interessados."),
        ("11 PROVIDÊNCIAS PRÉVIAS AO CONTRATO",
         "Revisão de TR; ofícios a empresas; PNCP; prazo mínimo de 8 dias úteis."),
        ("12 CONTRATAÇÕES CORRELATAS/INTERDEPENDENTES",
         "Dispensa emergencial (O2 Processo 1) em paralelo, com rescisão automática na homologação."),
        ("13 POSSÍVEIS IMPACTOS AMBIENTAIS",
         "Impactos mínimos; responsabilidade da contratada sobre cilindros/transporte."),
        ("14 ANÁLISE DE RISCOS",
         "Nova desertificação, preço defasado, exigências excessivas — tratados por pesquisa nova, "
         "habilitação proporcional e divulgação ampliada."),
        ("15 DECLARAÇÃO DE VIABILIDADE DA CONTRATAÇÃO",
         "Declaramos a viabilidade da contratação por Pregão Eletrônico com os ajustes deste ETP/TR."),
    ]
    for t, b in secoes:
        h2(doc, t)
        p(doc, b)
    table(doc, ["Lote", "Abrangência", "Qtde anual (m³)", "Valor unit. (R$)", "Valor total (R$)"], [
        ["1", "Zona Urbana", "2.184,0", "[____]", "[____]"],
        ["2", "Zona Rural", "288,0", "[____]", "[____]"],
    ])
    assinaturas_equipe(doc)
    save(doc, P2 / "O2_P2_ETP.docx")


def p2_tr():
    doc = new_doc()
    title(doc, "TERMO DE REFERÊNCIA – O2 PROCESSO 2 (PREGÃO ELETRÔNICO REVISADO)")
    h2(doc, "1 DO OBJETO, VALOR E CONDIÇÕES GERAIS DA CONTRATAÇÃO")
    p(doc, "1.1 Registro/contrato para fornecimento contínuo de oxigênio medicinal gasoso com comodato, em 2 lotes.")
    p(doc, "1.2 Objetos comuns, não de luxo.")
    p(doc, "1.3 Vigência de 12 (doze) meses, prorrogável na forma da Lei nº 14.133/2021.")
    p(doc, "1.4 Entrega parcelada sob demanda; início em até 10 dias úteis.")
    p(doc, "1.5 Valor total estimado: R$ [________].")
    h2(doc, "2 DA FUNDAMENTAÇÃO DA CONTRATAÇÃO")
    p(doc, f"2.1 Lei nº 14.133/2021; revisão motivada pela desertificação do Processo nº {D['processo_origem']}.")
    p(doc, "2.2 Pontos revisados: pesquisa de preços nova; prazos realistas; habilitação proporcional; "
           "comodato/manutenção expressos; divulgação ampliada.")
    h2(doc, "3 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO CONSIDERADO O CICLO DE VIDA DO OBJETO")
    p(doc, "Fornecimento, comodato, manutenção, logística, medição e eventual prorrogação contratual.")
    h2(doc, "4 GESTÃO/ANÁLISE DE RISCOS")
    p(doc, "Riscos de nova desertificação e atraso — mitigados por preço atualizado, exigências proporcionais e ofícios ao mercado.")
    h2(doc, "5 REQUISITOS DA CONTRATAÇÃO")
    p(doc, "5.1 Pureza ≥ 99,5%; comodato; ANVISA; assistência 24h; troca imediata de cilindro irregular.")
    p(doc, "5.2 Habilitação jurídica, fiscal, trabalhista, sanitária e atestado técnico compatível, sem exigências desarrazoadas.")
    h2(doc, "6 DA EXECUÇÃO DO OBJETO")
    table(doc, ["Lote", "Descrição", "Unid.", "Qtd. anual", "Valor unit. (R$)", "Valor total (R$)"], [
        ["1", "Oxigênio medicinal — Zona Urbana", "m³", "2.184,0", "[____]", "[____]"],
        ["2", "Oxigênio medicinal — Zona Rural", "m³", "288,0", "[____]", "[____]"],
    ])
    h2(doc, "7 DA GESTÃO DA ATA E DA CONTRATAÇÃO")
    p(doc, f"Gestor: {GESTOR[0]}. Fiscal: {FISCAL[0]}.")
    h2(doc, "8 RECEBIMENTO DO OBJETO")
    p(doc, "Provisório na entrega e definitivo após verificação de conformidade.")
    h2(doc, "9 DOS CRITÉRIOS DE MEDIÇÃO E PAGAMENTO")
    p(doc, "Mensal, por m³ efetivamente fornecido, com atesto.")
    h2(doc, "10 DA FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR")
    p(doc, "Pregão Eletrônico — menor preço por lote.")
    h2(doc, "11 DAS OBRIGAÇÕES DO CONTRATANTE")
    p(doc, "Pagar, fiscalizar e disponibilizar acesso às unidades.")
    h2(doc, "12 DAS OBRIGAÇÕES DO CONTRATADO")
    p(doc, "Cumprir o TR, manter habilitação e assegurar continuidade do abastecimento.")
    h2(doc, "13 DA ADEQUAÇÃO ORÇAMENTÁRIA")
    p(doc, f"Dotação: {D['projeto']} | Fonte: {D['fonte']}.")
    assinaturas_equipe(doc)
    save(doc, P2 / "O2_P2_TR.docx")


def p2_pesquisa():
    doc = new_doc()
    title(doc, "PESQUISA DE PREÇOS ATUALIZADA – O2 PROCESSO 2")
    p(doc, "Atualização motivada pela licitação deserta. Fontes: Painel de Preços, municípios similares e cotações.")
    table(doc, ["Fonte", "Data", "R$/m³ Lote 1", "R$/m³ Lote 2", "Documento", "Obs."], [
        ["Painel de Preços", "", "", "", "", ""],
        ["Município similar 1", "", "", "", "", ""],
        ["Município similar 2", "", "", "", "", ""],
        ["Cotação empresa A", "", "", "", "", ""],
        ["Cotação empresa B", "", "", "", "", ""],
        ["Cotação empresa C", "", "", "", "", ""],
    ])
    h2(doc, "CONSOLIDAÇÃO")
    p(doc, "Preço de referência Lote 1: R$ ________ / m³")
    p(doc, "Preço de referência Lote 2: R$ ________ / m³")
    p(doc, "Valor anual estimado: R$ _______________")
    p(doc, "Justificativa de diferença em relação ao edital anterior: ________________")
    assinaturas_equipe(doc)
    save(doc, P2 / "O2_P2_Pesquisa_Precos.docx")


def p2_edital():
    doc = new_doc()
    title(doc, "MINUTA DE EDITAL — PREGÃO ELETRÔNICO Nº _____/2026 – O2 PROCESSO 2")
    p(doc, f"Processo administrativo nº _____/2026 (ref. {D['processo_origem']})")
    h2(doc, "1 PREÂMBULO")
    p(doc, "O Município de Pelotas, por intermédio da Secretaria Municipal de Saúde, torna público o Pregão "
           "Eletrônico para fornecimento de oxigênio medicinal com comodato de cilindros, na forma da Lei nº 14.133/2021.")
    h2(doc, "2 OBJETO")
    p(doc, "Fornecimento contínuo em 2 lotes (zona urbana e zona rural), conforme Anexo I — Termo de Referência.")
    h2(doc, "3 DATAS")
    p(doc, "Início de acolhimento: ___/___/2026 às __:__")
    p(doc, "Limite de propostas: ___/___/2026 às __:__")
    p(doc, "Sessão pública: ___/___/2026 às __:__")
    p(doc, "Plataforma: ________________")
    h2(doc, "4 CRITÉRIO / PUBLICIDADE")
    p(doc, "Menor preço por lote. Prazo mínimo de 8 (oito) dias úteis. Divulgação no PNCP, Portal e ofícios a empresas do ramo.")
    h2(doc, "5 ANEXOS")
    for item in [
        "Anexo I — Termo de Referência;",
        "Anexo II — Minuta do Contrato;",
        "Anexo III — Modelo de Proposta;",
        "Anexo IV — Locais de Entrega.",
    ]:
        bullet_line(doc, item)
    assinaturas_equipe(doc)
    save(doc, P2 / "O2_P2_Minuta_Edital.docx")


def p2_contrato():
    doc = new_doc()
    title(doc, "MINUTA DE CONTRATO Nº _____/2026 – O2 PROCESSO 2")
    p(doc, "Contrato entre o Município de Pelotas / SMS e a empresa _____________, vencedora do Pregão "
           "Eletrônico nº _____/2026, Lote(s) _____.")
    for t, b in [
        ("CLÁUSULA 1 — OBJETO", "Fornecimento contínuo de oxigênio medicinal com comodato, conforme TR e proposta."),
        ("CLÁUSULA 2 — VIGÊNCIA", "12 meses, prorrogável na forma da Lei nº 14.133/2021."),
        ("CLÁUSULA 3 — PREÇO E PAGAMENTO", "R$ ______ / m³ (Lote ____). Pagamento mensal por medição."),
        ("CLÁUSULA 4 — OBRIGAÇÕES", "Cumprimento do TR; pagamento e fiscalização pela CONTRATANTE."),
        ("CLÁUSULA 5 — FISCALIZAÇÃO", f"Gestor: {GESTOR[0]}. Fiscal: {FISCAL[0]}."),
        ("CLÁUSULA 6 — SANÇÕES", "Arts. 155 a 163 da Lei nº 14.133/2021."),
        ("CLÁUSULA 7 — RESCISÃO", "Nos casos legais, com contraditório e ampla defesa."),
        ("CLÁUSULA 8 — FORO", "Comarca de Pelotas/RS."),
    ]:
        h2(doc, t)
        p(doc, b)
    assinaturas_equipe(doc)
    save(doc, P2 / "O2_P2_Minuta_Contrato.docx")


def p2_oficio():
    doc = new_doc()
    title(doc, "OFÍCIO-CIRCULAR — DIVULGAÇÃO AMPLIADA DO PREGÃO O2")
    p(doc, "Ofício nº _____/2026 — SMS/Licitações", justify=False)
    p(doc, "À empresa: _______________________________", justify=False)
    p(doc, "Informamos a publicação/publicação iminente de Pregão Eletrônico para fornecimento de oxigênio "
           f"medicinal com comodato. O certame anterior (Processo nº {D['processo_origem']}) restou deserto. "
           "Convidamos à participação.")
    p(doc, "Link PNCP / edital: ________________________________")
    p(doc, "Data da sessão: ___/___/2026")
    table(doc, ["Empresa", "E-mail/WhatsApp", "Data envio", "Comprovante"], [
        ["", "", "", "[ ]"], ["", "", "", "[ ]"], ["", "", "", "[ ]"], ["", "", "", "[ ]"], ["", "", "", "[ ]"],
    ])
    assinaturas_equipe(doc)
    save(doc, P2 / "O2_P2_Oficio_Divulgacao.docx")


def write_readme():
    (OUT / "README.md").write_text(
        """# Processos O2 — Oxigênio Medicinal (padrão RP)

Documentos novos dos **dois processos simultâneos** (Dispensa emergencial + Pregão),
elaborados no **mesmo padrão** dos Registros de Preços RP02–RP13 do painel
`painel-compras-empenhos` (cabeçalho com brasão, seções DFD/ETP/TR, equipe
Gestor/Fiscal padronizada).

> O conteúdo do ZIP/painel do Drive **não foi alterado**.

## Pastas

- `O2_P1_Dispensa_Emergencial/` — Opção C (Art. 75, VIII)
- `O2_P2_Pregao_Paralelo/` — Opção A (Pregão revisado)
- Peças comuns na raiz desta pasta (`O2_Indice`, despacho, cronograma)

## Regenerar

```bash
python3 scripts/gerar_o2_padrao_rp.py
```
""",
        encoding="utf-8",
    )


def main():
    for d in (OUT, P1, P2):
        d.mkdir(parents=True, exist_ok=True)
    write_readme()
    o2_indice()
    o2_despacho()
    o2_cronograma()
    p1_dfd(); p1_etp(); p1_tr(); p1_justificativa(); p1_mapa(); p1_razao(); p1_contrato(); p1_autorizacao()
    p2_dfd(); p2_etp(); p2_tr(); p2_pesquisa(); p2_edital(); p2_contrato(); p2_oficio()
    print("\nProcessos O2 gerados no padrão RP.")


if __name__ == "__main__":
    main()
