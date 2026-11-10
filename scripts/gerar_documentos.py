#!/usr/bin/env python3
"""Gera o pacote documental dos Processos 1 (Dispensa Emergencial) e 2 (Pregão Regular)."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

ROOT = Path(__file__).resolve().parents[1]
OUT1 = ROOT / "docs" / "processo-1-dispensa-emergencial"
OUT2 = ROOT / "docs" / "processo-2-pregao-regular"
OUTA = ROOT / "docs" / "anexos"

# Dados consolidados dos autos / DFD / Projeto O2
DADOS = {
    "municipio": "Município de Pelotas",
    "estado": "Estado do Rio Grande do Sul",
    "sms": "Secretaria Municipal de Saúde",
    "unidade": "Diretoria de Atenção Primária",
    "processo_origem": "1242/2026",
    "sl": "00710/2026",
    "memo": "003746/2026",
    "memo_ccgof": "005720/2026",
    "abertura": "27/01/2026",
    "data_deserto": "29/10/2026",
    "projeto_atividade": "10.301.0105.2167.00 — Atenção Primária em Saúde",
    "fonte": "1600.000",
    "conta": "Agência 2820 / C/C 2200/575835306-6 — Caixa Econômica Federal",
    "consumo_ubs_urbana_mes": "152,0 m³",
    "consumo_ubai_mes": "30,0 m³",
    "consumo_rural_mes": "24,0 m³",
    "consumo_total_mes": "206,0 m³",
    "consumo_urbano_ano": "2.184,0 m³",
    "consumo_rural_ano": "288,0 m³",
    "consumo_total_ano": "2.472,0 m³",
    "prazo_emergencial": "180 (cento e oitenta) dias",
    "prazo_emergencial_curto": "180 dias",
    "entrega_urgencia": "24 (vinte e quatro) horas",
    "entrega_critica": "12 (doze) horas",
    "contato": "Carol — (53) 98462-4786",
    "endereco_sms": "Rua Lobo da Costa, nº 1764, Pelotas/RS",
    "diretora": "Cândida Silveira Rodrigues",
    "secretaria": "Angela Moreira Vitória",
    "solicitante": "Josué Silva",
}


UNIDADES_URBANAS = [
    ("1", "Arco Íris", "Rua Pery Ribas, nº 523 — Três Vendas", "3277-6203"),
    ("2", "Areal Fundos (UFPEL)", "Av. Domingos de Almeida, nº 4265 — Areal", "3284-4336"),
    ("3", "Areal I", "Rua Apolinário Porto Alegre, nº 290 — Areal", "3278-3531"),
    ("4", "Balsa", "Rua João Thomaz Munhoz, nº 270 — Balsa", "3199-0623"),
    ("5", "Barro Duro", "Praça Aratiba, nº 12", "3199-0729"),
    ("6", "Bom Jesus", "Av. Itália, nº 350 — Jardim Europa/Areal", "3199-0803"),
    ("7", "CAIC Pestano (UCPEL)", "Rua Leopoldo Brod, nº 3220 — Pestano", "3199-0575"),
    ("8", "Cohab Fragata", "Rua Paulo Simões Lopes, nº 230 — Fragata", "3271-9425"),
    ("9", "Cohab Guabiroba", "Rua Dr. Arnaldo da Silva Ferreira, nº 352 — Fragata", "3199-0541"),
    ("10", "Cohab Lindóia", "Rua Ernani Osmar Blass, nº 344 — Três Vendas", "3199-0812"),
    ("11", "Cohab Pestano", "Rua Leopoldo Brod, nº 2297 — Pestano", "3199-0652"),
    ("12", "CSU Areal (UFPEL)", "Rua Guararapes, nº 50 A — Areal", "3284-4334"),
    ("13", "CSU Cruzeiro do Sul", "Rua Barão de Itamaracá, nº 690 — Cruzeiro", "3199-0547"),
    ("14", "Dom Pedro I", "Rua Ulisses Batinga, nº 749 — Fragata", "3271-3383"),
    ("15", "Dunas", "Av. Ulysses Silveira Guimarães, nº 1988 — Dunas", "3228-4666"),
    ("16", "Fátima", "Rua Baldomero Trápaga, nº 480 — Fátima", "3199-0845"),
    ("17", "Fraget", "Rua Jorn. Carlos Andrade, nº 241 — Fragata", "3199-0690"),
    ("18", "Getúlio Vargas", "Rua Sete, Lot. G. Vargas, nº 184", "3199-0894"),
    ("19", "Jardim de Allah", "Av. Fernando Osório, nº 7430 / Rua Dois, nº 1 — Três Vendas", "3199-0859"),
    ("20", "Laranjal", "Rua São Borja, nº 683 — Laranjal", "3226-4488"),
    ("21", "Leocádia", "Rua David Canabarro, nº 890 — Areal", "3282-2476"),
    ("22", "Navegantes", "Rua Darci Vargas, nº 212 — Navegantes", "3199-0742"),
    ("23", "Obelisco", "Rua Francisco R. da Silva, nº 435 — Obelisco", "3282-2477"),
    ("24", "Osório", "Rua Quatro, Lot. Osório, nº 14 / Barão de Mauá, nº 229", "3199-0806"),
    ("25", "PAM Fragata", "Av. Pinheiro Machado, nº 168 — Fragata", "3221-0362"),
    ("26", "Posto Branco", "Estrada Porto Alegre — Antiga Granja Retiro", "—"),
    ("27", "Puericultura", "Rua João Pessoa, nº 220 — Centro", "3199-0830"),
    ("28", "Py Crespo", "Rua Marquês de Olinda, nº 1291 — Py Crespo", "3199-0736"),
    ("29", "Salgado Filho", "Av. Salgado Filho, nº 912", "3199-0914"),
    ("30", "Sanga Funda", "Av. Ildefonso Simões Lopes, nº 5225 — Três Vendas", "3199-0926"),
    ("31", "SANSCA", "Rua Dr. Amarante, nº 919 — Centro", "3199-0517"),
    ("32", "Santa Terezinha", "Rua São Miguel, nº 5 — Santa Teresinha", "3199-0523"),
    ("33", "Simões Lopes", "Rua Visconde da Graça, nº 107 — Simões Lopes", "3199-0923"),
    ("34", "Sítio Floresta", "Rua Ignácio Teixeira Machado, nº 299 — ao lado da BR-116", "3278-9188"),
    ("35", "UBAI Navegantes", "Rua Darci Vargas, nº 212 B — Navegantes", "3279-4546"),
    ("36", "União de Bairros", "Av. Fernando Osório, nº 4434 — Três Vendas", "3199-0668"),
    ("37", "UBAI Lindóia", "Rua Ernani Osmar Blass, nº 344 — Três Vendas", "3199-0812"),
    ("38", "UBAI Fraget", "Rua Jorn. Carlos Andrade, nº 241 — Fragata", "3199-0690"),
    ("39", "Vila Municipal", "Rua Luciano Galletti, nº 600 — Três Vendas", "3199-0907"),
    ("40", "Vila Princesa", "Av. Quatro, Vila Princesa, nº 3205", "3278-1564"),
    ("41", "Virgílio Costa", "Rua Epitácio Pessoa, nº 1291 — Fragata", "3199-0897"),
]

UNIDADES_RURAIS = [
    ("1", "Colônia Cascata", "Est. da Cascatinha, nº 4656 — 5º Distrito / BR-392", "22,8 km", "3227-5108"),
    ("2", "Colônia Cerrito Alegre", "Est. Cerrito Alegre, nº 5780 — 3º Distrito", "28,7 km", "3227-7054"),
    ("3", "Colônia Cordeiro de Farias", "Ponte Cordeiro de Farias, nº 4455 — 5º Distrito", "33,5 km", "3224-7309"),
    ("4", "Colônia Grupelli", "Est. Grupelli, nº 4525 — 7º Distrito", "50,2 km", "3224-5011"),
    ("5", "Colônia Maciel", "Colônia Maciel, nº 7240 — 8º Distrito", "39,6 km", "3199-0774"),
    ("6", "UBS Pedreiras (Monte Bonito)", "Estrada Monte Bonito — 9º Distrito", "17,3 km", "3277-3272"),
    ("7", "UBS Monte Bonito", "Estrada Monte Bonito — 9º Distrito", "23,3 km", "3277-3273"),
    ("8", "Colônia Osório", "Colônia Osório — 4º Distrito", "36,7 km", "3277-4172"),
    ("9", "Colônia Santa Silvana", "Colônia Santa Silvana, nº 6010", "54,1 km", "3277-8043"),
    ("10", "Colônia Triunfo", "Colônia Triunfo, nº 4320 — 8º Distrito", "59,8 km", "3199-0709"),
    ("11", "Colônia Vila Nova", "Colônia Vila Nova — 7º Distrito", "36,8 km", "3224-7275"),
    ("12", "Corrientes", "BR-116 — Corrientes, nº 6035 — Ponte do Retiro", "42,7 km", "3274-6605"),
]


def set_doc_defaults(doc: Document) -> None:
    section = doc.sections[0]
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(3)
    section.right_margin = Cm(2.5)
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    style.font.size = Pt(12)
    pf = style.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    pf.space_after = Pt(6)


def add_center(doc: Document, text: str, *, bold: bool = False, size: int = 12, space_after: int = 6):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    run.font.name = "Times New Roman"
    p.paragraph_format.space_after = Pt(space_after)
    return p


def add_just(doc: Document, text: str, *, bold: bool = False, first_line: float | None = 1.25):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run(text)
    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    if first_line is not None:
        p.paragraph_format.first_line_indent = Cm(first_line)
    return p


def add_left(doc: Document, text: str, *, bold: bool = False, indent: float = 0):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(text)
    run.bold = bold
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)
    if indent:
        p.paragraph_format.left_indent = Cm(indent)
    return p


def add_heading_local(doc: Document, text: str):
    return add_left(doc, text, bold=True)


def cabecalho(doc: Document, titulo: str, subtitulo: str | None = None):
    add_center(doc, DADOS["estado"].upper(), bold=True, size=12)
    add_center(doc, DADOS["municipio"].upper(), bold=True, size=12)
    add_center(doc, DADOS["sms"].upper(), bold=True, size=12, space_after=18)
    add_center(doc, titulo.upper(), bold=True, size=13, space_after=6)
    if subtitulo:
        add_center(doc, subtitulo, bold=False, size=12, space_after=18)
    else:
        doc.add_paragraph()


def assinaturas(doc: Document, linhas: list[str]):
    doc.add_paragraph()
    add_center(doc, "Pelotas/RS, _____ de ______________ de 2026.")
    doc.add_paragraph()
    for bloco in linhas:
        add_center(doc, "_" * 45, space_after=0)
        for linha in bloco:
            add_center(doc, linha, space_after=0)
        doc.add_paragraph()


def add_table(doc: Document, headers: list[str], rows: list[list[str]]):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        for p in cell.paragraphs:
            for r in p.runs:
                r.bold = True
                r.font.name = "Times New Roman"
                r.font.size = Pt(10)
    for r_idx, row in enumerate(rows, start=1):
        for c_idx, value in enumerate(row):
            cell = table.rows[r_idx].cells[c_idx]
            cell.text = value
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.name = "Times New Roman"
                    run.font.size = Pt(9)
    doc.add_paragraph()
    return table


def save(doc: Document, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    print(f"OK {path.relative_to(ROOT)}")


# ---------------------------------------------------------------------------
# PROCESSO 1
# ---------------------------------------------------------------------------

def p1_capa():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Capa do Processo", "Processo 1 — Dispensa Emergencial (Art. 75, VIII, Lei nº 14.133/2021)")
    add_left(doc, f"Processo administrativo nº: _______________ / 2026 (vinculado ao Processo nº {DADOS['processo_origem']})")
    add_left(doc, f"Unidade demandante: {DADOS['sms']} — {DADOS['unidade']}")
    add_left(doc, f"Solicitante: {DADOS['solicitante']}")
    add_left(doc, f"Contato: {DADOS['contato']}")
    add_left(doc, f"Endereço: {DADOS['endereco_sms']}")
    doc.add_paragraph()
    add_heading_local(doc, "1. Objeto")
    add_just(
        doc,
        "Contratação emergencial de empresa especializada no fornecimento contínuo, sob demanda, "
        "de Oxigênio Medicinal gasoso, com cessão de cilindros em regime de comodato, para atender "
        "as Unidades Básicas de Saúde (UBSs) e Unidades Básicas de Atendimento Imediato (UBAIs) "
        "das zonas urbana e rural do Município de Pelotas/RS, pelo prazo estritamente necessário "
        f"à conclusão da licitação regular, limitado a {DADOS['prazo_emergencial']}.",
    )
    add_heading_local(doc, "2. Fundamentação legal")
    add_just(
        doc,
        "Art. 75, inciso VIII, da Lei nº 14.133/2021 (dispensa em razão de licitação deserta e/ou "
        "situação emergencial que possa ocasionar prejuízo ou comprometer a segurança de pessoas, "
        "obras, serviços, equipamentos e outros bens, públicos ou particulares).",
    )
    add_heading_local(doc, "3. Origem / fato gerador")
    add_just(
        doc,
        f"O Processo nº {DADOS['processo_origem']}, destinado à aquisição de oxigênio medicinal "
        f"para UBSs e UBAIs, teve o certame restado deserto, conforme tramitação de {DADOS['data_deserto']} "
        f"(SL nº {DADOS['sl']}; Memorando nº {DADOS['memo']}; Memorando CCGOF nº {DADOS['memo_ccgof']}). "
        "A continuidade do abastecimento é indispensável à assistência à saúde.",
    )
    add_heading_local(doc, "4. Quantitativo de referência (período emergencial)")
    add_just(
        doc,
        f"Consumo médio mensal consolidado: {DADOS['consumo_total_mes']} "
        f"(zona urbana {DADOS['consumo_ubs_urbana_mes']} em UBSs + {DADOS['consumo_ubai_mes']} em UBAIs; "
        f"zona rural {DADOS['consumo_rural_mes']}). "
        f"Para {DADOS['prazo_emergencial_curto']}: aproximadamente "
        f"{round(206 * 6, 1):.1f} m³ (estimativa referencial, sem obrigação de consumo mínimo).",
    )
    add_heading_local(doc, "5. Valor estimado")
    add_just(doc, "R$ _______________ (conforme mapa de preços anexo — a preencher após cotações).")
    add_heading_local(doc, "6. Prazo")
    add_just(
        doc,
        f"{DADOS['prazo_emergencial']}, contados da assinatura do contrato, ou até a homologação/"
        "assinatura do contrato decorrente do Pregão Eletrônico instaurado em paralelo, o que ocorrer primeiro. "
        "Vedada a prorrogação além do prazo legalmente admitido para a emergência.",
    )
    add_heading_local(doc, "7. Dotação orçamentária")
    add_left(doc, f"Projeto/Atividade: {DADOS['projeto_atividade']}")
    add_left(doc, f"Fonte: {DADOS['fonte']}")
    add_left(doc, f"Conta: {DADOS['conta']}")
    add_left(doc, "Natureza da despesa: _______________")
    assinaturas(
        doc,
        [
            [DADOS["diretora"], "Diretora de Atenção Primária"],
            [DADOS["secretaria"], "Secretária Municipal de Saúde"],
        ],
    )
    save(doc, OUT1 / "01-capa-processo-dispensa-emergencial.docx")


def p1_dfd():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Documento de Formalização da Demanda — DFD", "Versão emergencial — Dispensa Art. 75, VIII")
    add_left(doc, f"Unidade demandante: {DADOS['sms']} — {DADOS['unidade']}")
    add_left(doc, f"Processo de origem: nº {DADOS['processo_origem']}")
    doc.add_paragraph()
    add_heading_local(doc, "1. Objeto")
    add_just(
        doc,
        "Contratação emergencial de empresa especializada para fornecimento contínuo, sob demanda, "
        "de Gases Medicinais (Oxigênio Medicinal), com cessão de cilindros em regime de comodato, "
        "destinada a UBSs e UBAIs das zonas urbana e rural de Pelotas/RS, pelo prazo estritamente "
        f"necessário à conclusão da licitação regular ({DADOS['prazo_emergencial_curto']}).",
    )
    add_left(doc, "Tipo do objeto: ( x ) Serviço continuado sem dedicação exclusiva de mão de obra")
    add_heading_local(doc, "2. Forma de contratação sugerida")
    add_just(
        doc,
        "Dispensa de licitação, com fulcro no art. 75, inciso VIII, da Lei nº 14.133/2021, "
        "em razão da licitação deserta do Processo nº "
        f"{DADOS['processo_origem']} e do risco de desabastecimento de oxigênio medicinal.",
    )
    add_heading_local(doc, "3. Justificativa da necessidade")
    add_just(
        doc,
        "O oxigênio medicinal é insumo essencial ao suporte à vida, à estabilização clínica e à "
        "realização de procedimentos assistenciais (inalações e nebulizações) na Atenção Primária. "
        f"O certame do Processo nº {DADOS['processo_origem']} restou deserto. A tramitação média "
        "de novo pregão (30 a 45 dias) é incompatível com a urgência do abastecimento. "
        "A situação não decorre de falta de planejamento, mas de ausência de interessados no certame regular. "
        "Paralelamente, instaurar-se-á novo Pregão Eletrônico para solução definitiva.",
    )
    add_heading_local(doc, "4. Quantidade e descrição")
    add_just(
        doc,
        "Fornecimento sob demanda, faturado pelo volume efetivamente entregue (m³). "
        "Cilindros em comodato, sem ônus de locação. Responsabilidade da contratada pela logística, "
        "troca, manutenção, testes hidrostáticos e conformidade ANVISA/ABNT.",
    )
    add_table(
        doc,
        ["Item", "Especificação", "Unidade", "Qtde. mensal", "Qtde. 180 dias"],
        [
            [
                "1",
                "Oxigênio medicinal gasoso — Zona Urbana (UBSs + UBAIs)",
                "m³",
                "182,0",
                "1.092,0",
            ],
            [
                "2",
                "Oxigênio medicinal gasoso — Zona Rural (12 UBSs)",
                "m³",
                "24,0",
                "144,0",
            ],
            ["", "TOTAL ESTIMADO", "m³", "206,0", "1.236,0"],
        ],
    )
    add_just(
        doc,
        "Anexar: ata/termo da licitação deserta; relatório de estoque atual (em dias); "
        "DFD originário do Processo nº 1242/2026.",
    )
    add_heading_local(doc, "5. Estimativa de valor")
    add_just(doc, "A preencher conforme mapa de preços (mínimo de 3 cotações ou justificativa fundamentada).")
    add_heading_local(doc, "6. Dotação")
    add_left(doc, f"Projeto/Atividade: {DADOS['projeto_atividade']}")
    add_left(doc, f"Fonte: {DADOS['fonte']}")
    add_left(doc, "Valor estimado: R$ _______________")
    add_heading_local(doc, "7. Locais de entrega")
    add_just(doc, "Conforme Anexo de Locais de Entrega (41 pontos urbanos e 12 rurais), integrantes deste DFD.")
    add_heading_local(doc, "8. Início da prestação")
    add_just(
        doc,
        f"Imediato após a assinatura do contrato / emissão da Nota de Empenho, com primeira entrega "
        f"em até {DADOS['entrega_urgencia']}, e em até {DADOS['entrega_critica']} em situações críticas "
        "comunicadas pela unidade.",
    )
    assinaturas(
        doc,
        [
            [DADOS["diretora"], "Diretora de Atenção Primária"],
            [DADOS["secretaria"], "Secretária Municipal de Saúde — Autorização"],
        ],
    )
    save(doc, OUT1 / "02-dfd-emergencial.docx")


def p1_etp():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Estudo Técnico Preliminar — ETP Simplificado", "Contratação emergencial de oxigênio medicinal")
    add_heading_local(doc, "1. Necessidade")
    add_just(
        doc,
        "Assegurar o abastecimento ininterrupto de oxigênio medicinal nas UBSs e UBAIs municipais "
        f"após a licitação deserta do Processo nº {DADOS['processo_origem']}, enquanto se conclui "
        "novo Pregão Eletrônico instaurado em paralelo.",
    )
    add_heading_local(doc, "2. Área requisitante")
    add_just(doc, f"{DADOS['sms']} / {DADOS['unidade']}.")
    add_heading_local(doc, "3. Descrição da solução")
    add_just(
        doc,
        "Fornecimento de oxigênio medicinal gasoso (pureza mínima 99,5%), em cilindros cedidos em "
        "comodato, com logística de entrega/retirada, manutenção e assistência técnica 24h, "
        "conforme normas da ANVISA e ABNT.",
    )
    add_heading_local(doc, "4. Estimativa de consumo e estoque")
    add_just(
        doc,
        f"Consumo médio mensal: {DADOS['consumo_total_mes']} "
        f"(urbano {DADOS['consumo_ubs_urbana_mes']} + UBAIs {DADOS['consumo_ubai_mes']}; "
        f"rural {DADOS['consumo_rural_mes']}). "
        "Estoque atual estimado: _____ dias de autonomia (preencher com relatório anexo). "
        "Consequência da falta: risco à vida e interrupção de procedimentos essenciais.",
    )
    add_heading_local(doc, "5. Requisitos da contratação")
    for item in [
        "Oxigênio medicinal gasoso, pureza ≥ 99,5%;",
        "Cilindros em comodato (capacidades usuais 7 m³ / 10 m³ ou equivalente);",
        f"Entrega em até {DADOS['entrega_urgencia']} e {DADOS['entrega_critica']} em urgência clínica;",
        "Laudo de pureza / conformidade sanitária;",
        "Troca imediata de cilindro defeituoso;",
        "Atendimento 24h para contingências;",
        "Cobertura de todas as unidades listadas (urbana e rural).",
    ]:
        add_left(doc, f"• {item}", indent=0.5)
    add_heading_local(doc, "6. Estimativa de quantidades para o período emergencial")
    add_just(
        doc,
        f"Prazo: {DADOS['prazo_emergencial_curto']}. Volume estimado: 1.236,0 m³ "
        "(sem consumo mínimo obrigatório; pagamento por medição).",
    )
    add_heading_local(doc, "7. Estimativa de valor")
    add_just(doc, "R$ _______________ — a consolidar no mapa de preços.")
    add_heading_local(doc, "8. Justificativa da escolha da solução")
    add_just(
        doc,
        "A dispensa emergencial é a única medida capaz de evitar desabastecimento imediato. "
        "A solução definitiva permanece a licitação regular (Pregão Eletrônico) em curso paralelo, "
        "com TR revisado, pesquisa atualizada e divulgação ampliada.",
    )
    add_heading_local(doc, "9. Providências prévias / alinhamentos")
    add_just(
        doc,
        "Articulação com Compras e Licitações; abertura simultânea do Processo 2 (Pregão); "
        "coleta de cotações; relatório de estoque; indicação de fiscal e gestor do contrato.",
    )
    add_heading_local(doc, "10. Conclusão")
    add_just(
        doc,
        "Recomenda-se a contratação direta emergencial pelo prazo estritamente necessário, "
        "com rescisão automática quando homologado o certame regular.",
    )
    assinaturas(doc, [[DADOS["diretora"], "Diretora de Atenção Primária — Elaboração do ETP"]])
    save(doc, OUT1 / "03-etp-simplificado.docx")


def p1_tr():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Termo de Referência Emergencial", "Fornecimento de oxigênio medicinal — Dispensa Art. 75, VIII")
    add_heading_local(doc, "1. Objeto")
    add_just(
        doc,
        "Contratação emergencial de empresa especializada no fornecimento contínuo, sob demanda, "
        "de Oxigênio Medicinal gasoso, com comodato de cilindros, manutenção e logística de entrega "
        f"nas UBSs e UBAIs de Pelotas/RS, pelo prazo de {DADOS['prazo_emergencial']}, ou até a "
        "conclusão da licitação regular, o que ocorrer primeiro.",
    )
    add_heading_local(doc, "2. Fundamentação")
    add_just(
        doc,
        f"Art. 75, VIII, Lei nº 14.133/2021; Processo nº {DADOS['processo_origem']} (licitação deserta); "
        "DFD e ETP deste expediente.",
    )
    add_heading_local(doc, "3. Especificação técnica")
    for item in [
        "Oxigênio medicinal gasoso, pureza mínima de 99,5% (v/v);",
        "Conformidade com normas ANVISA e ABNT aplicáveis a gases medicinais;",
        "Fornecimento em cilindros de 7 m³, 10 m³ ou outro volume usual do mercado, em comodato;",
        "Identificação, validade de teste hidrostático e válvulas em perfeito estado;",
        "Laudo de análise/qualidade quando solicitado pela fiscalização.",
    ]:
        add_left(doc, f"• {item}", indent=0.5)
    add_heading_local(doc, "4. Quantitativos estimados")
    add_table(
        doc,
        ["Lote/Área", "Consumo médio mensal", "Estimativa 180 dias"],
        [
            ["Zona Urbana (UBSs + UBAIs)", "182,0 m³", "1.092,0 m³"],
            ["Zona Rural (12 UBSs)", "24,0 m³", "144,0 m³"],
            ["Total", "206,0 m³", "1.236,0 m³"],
        ],
    )
    add_just(doc, "Os quantitativos são referenciais. Não há obrigação de consumo mínimo.")
    add_heading_local(doc, "5. Forma de fornecimento e prazos")
    add_just(
        doc,
        f"Entrega sob demanda, mediante solicitação das unidades/Diretoria. Prazo ordinário: até "
        f"{DADOS['entrega_urgencia']}. Situações críticas: até {DADOS['entrega_critica']}. "
        "Substituição imediata de cilindro com vazamento, válvula defeituosa ou irregularidade sanitária.",
    )
    add_heading_local(doc, "6. Locais de entrega")
    add_just(
        doc,
        "Conforme relação anexa (41 unidades urbanas e 12 rurais). A contratada deverá cobrir "
        "ambas as zonas, salvo contratação parcial justificada por cotações.",
    )
    add_heading_local(doc, "7. Obrigações da contratada")
    for item in [
        "Manter estoque operacional compatível com a demanda municipal;",
        "Realizar comodato, manutenção preventiva/corretiva e testes hidrostáticos;",
        "Disponibilizar canal de atendimento 24 horas;",
        "Apresentar documentação sanitária e registro ANVISA do produto;",
        "Emitir notas fiscais conforme medições mensais;",
        "Cumprir legislação trabalhista, tributária e ambiental aplicável.",
    ]:
        add_left(doc, f"• {item}", indent=0.5)
    add_heading_local(doc, "8. Obrigações da contratante")
    add_just(
        doc,
        "Efetuar o pagamento das medições; indicar fiscal e gestor; disponibilizar acesso às unidades; "
        "comunicar demanda com a antecedência compatível com a operação, ressalvadas urgências.",
    )
    add_heading_local(doc, "9. Fiscalização e gestão")
    add_left(doc, "Gestor do contrato: _______________________________ / Matrícula: __________")
    add_left(doc, "Fiscal do contrato: ________________________________ / Matrícula: __________")
    add_heading_local(doc, "10. Pagamento")
    add_just(
        doc,
        "Mensal, mediante medição do volume efetivamente fornecido (m³), apresentação de NF "
        "e atesto da fiscalização, no prazo da legislação municipal.",
    )
    add_heading_local(doc, "11. Vigência e rescisão")
    add_just(
        doc,
        f"Vigência de {DADOS['prazo_emergencial_curto']}, sem prorrogação ordinária. "
        "Rescisão automática quando homologado/assinado o contrato do Pregão regular em paralelo, "
        "assegurada a transição operacional sem desabastecimento.",
    )
    add_heading_local(doc, "12. Sanções")
    add_just(
        doc,
        "Aplicam-se as sanções da Lei nº 14.133/2021 e do contrato, especialmente em caso de "
        "atraso injustificado, desabastecimento imputável à contratada ou irregularidade sanitária.",
    )
    assinaturas(doc, [[DADOS["diretora"], "Diretora de Atenção Primária"]])
    save(doc, OUT1 / "04-termo-de-referencia-emergencial.docx")


def p1_justificativa():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Justificativa da Emergência", f"Processo nº {DADOS['processo_origem']} — licitação deserta")
    add_just(
        doc,
        f"A licitação referente ao Processo nº {DADOS['processo_origem']} (SL nº {DADOS['sl']}) "
        f"restou deserta, conforme comunicação oficial de {DADOS['data_deserto']} e tramitações "
        f"precedentes (Memorando nº {DADOS['memo']}; Memorando CCGOF nº {DADOS['memo_ccgof']}), "
        "conforme ata/documentos anexos.",
    )
    add_just(
        doc,
        "O estoque atual de oxigênio medicinal nas unidades da Atenção Primária é de aproximadamente "
        "_____ (_____) dias, insuficiente para aguardar a conclusão de novo certame, cuja tramitação "
        "média observada é de 30 a 45 dias, o que colocaria em risco a vida de pacientes e a "
        "continuidade do serviço público de saúde.",
    )
    add_just(
        doc,
        "A situação não decorre de falta de planejamento da Administração, mas de ausência de "
        "interessados no certame regular. O oxigênio medicinal é insumo crítico, sem substituto "
        "equivalente imediato no âmbito das UBSs e UBAIs.",
    )
    add_just(
        doc,
        "Diante do risco concreto de desabastecimento, justifica-se a contratação direta emergencial "
        "com base no art. 75, inciso VIII, da Lei nº 14.133/2021, pelo prazo estritamente necessário "
        f"({DADOS['prazo_emergencial_curto']}), enquanto se promove, em paralelo, a repetição do "
        "certame (Opção A — Pregão Eletrônico), com revisão de Termo de Referência, atualização de "
        "pesquisa de preços e ampliação da divulgação a empresas do ramo.",
    )
    add_just(
        doc,
        "O preço a ser contratado será justificado mediante pesquisa de mercado (mapa anexo) e "
        "permanecerá compatível com os valores praticados na região. As condições essenciais do "
        "objeto (especificação do oxigênio, comodato e obrigações sanitárias) serão mantidas, "
        "ajustando-se apenas o que for indispensável à viabilidade emergencial e à atratividade "
        "do novo pregão.",
    )
    add_heading_local(doc, "Anexos obrigatórios")
    for item in [
        "Ata / termo / despacho da licitação deserta;",
        "Relatório / print do estoque atual (em dias);",
        "DFD e ETP;",
        "Mapa de preços e cotações;",
        "Minuta contratual.",
    ]:
        add_left(doc, f"• {item}", indent=0.5)
    assinaturas(
        doc,
        [
            [DADOS["diretora"], "Diretora de Atenção Primária"],
            [DADOS["secretaria"], "Secretária Municipal de Saúde"],
        ],
    )
    save(doc, OUT1 / "05-justificativa-da-emergencia.docx")


def p1_pesquisa_precos():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Mapa de Pesquisa de Preços — Emergencial", "Oxigênio medicinal (m³) com comodato de cilindros")
    add_just(
        doc,
        "Pesquisa elaborada nos termos do art. 23 da Lei nº 14.133/2021. Preferência por cotações "
        "de empresas de gases medicinais da região, painéis oficiais e contratações similares. "
        "Mínimo de 3 (três) fontes; se houver menos respostas, justificar.",
    )
    add_table(
        doc,
        ["Fonte / Empresa", "CNPJ", "Valor unit. (R$/m³) Urbano", "Valor unit. (R$/m³) Rural", "Validade", "Obs."],
        [
            ["Cotação 1 — _____________", "", "", "", "", ""],
            ["Cotação 2 — _____________", "", "", "", "", ""],
            ["Cotação 3 — _____________", "", "", "", "", ""],
            ["Painel de Preços / similar", "", "", "", "", ""],
        ],
    )
    add_heading_local(doc, "Consolidação")
    add_left(doc, "Média / mediana adotada — Zona Urbana: R$ ________ / m³")
    add_left(doc, "Média / mediana adotada — Zona Rural: R$ ________ / m³")
    add_left(doc, "Valor estimado 180 dias (1.236 m³, se preço único): R$ _______________")
    add_left(doc, "Metodologia adotada: ( ) média  ( ) mediana  ( ) menor preço compatível  ( ) outra: ______")
    add_heading_local(doc, "Justificativa do preço")
    add_just(
        doc,
        "O valor escolhido mostra-se compatível com o mercado local/regional, considerando logística "
        "de comodato, distância das unidades rurais e urgência do abastecimento. "
        "Detalhar: _______________________________________________________________.",
    )
    assinaturas(doc, [["Responsável pela pesquisa", "Setor de Licitações / Compras — SMS"]])
    save(doc, OUT1 / "06-mapa-pesquisa-de-precos.docx")


def p1_razao_escolha():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Razão da Escolha do Fornecedor e Justificativa do Preço")
    add_left(doc, "Empresa selecionada: _______________________________________________")
    add_left(doc, "CNPJ: _____________________________")
    add_left(doc, "Endereço / contato: _______________________________________________")
    doc.add_paragraph()
    add_heading_local(doc, "1. Razão da escolha")
    add_just(
        doc,
        "A empresa acima foi selecionada porque: ( ) apresentou o menor preço compatível com a "
        "pesquisa de mercado; ( ) é a única que assegura entrega em até 24h com estoque imediato; "
        "( ) é a única com habilitação sanitária e capacidade logística para zona urbana e rural "
        "no prazo emergencial; ( ) outra: ________________________________.",
    )
    add_heading_local(doc, "2. Comparativo com a pesquisa")
    add_just(
        doc,
        "O preço unitário ofertado de R$ ________ / m³ situa-se "
        "( ) abaixo / ( ) na média / ( ) até _____% acima da média/mediana do mapa de preços, "
        "permanecendo justificado em face da urgência, do comodato e da cobertura territorial.",
    )
    add_heading_local(doc, "3. Conclusão")
    add_just(
        doc,
        "Face ao art. 75, VIII, da Lei nº 14.133/2021, e à comprovação de compatibilidade de preços, "
        "recomenda-se a contratação direta da empresa indicada pelo prazo emergencial.",
    )
    assinaturas(doc, [["Responsável técnico-administrativo", "SMS / Compras / Licitações"]])
    save(doc, OUT1 / "07-razao-da-escolha-do-fornecedor.docx")


def p1_habilitacao():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Checklist de Habilitação da Empresa", "Dispensa emergencial — oxigênio medicinal")
    add_left(doc, "Razão social: ________________________________")
    add_left(doc, "CNPJ: ____________________")
    doc.add_paragraph()
    rows = [
        ["Contrato social / ato constitutório atualizado", "[ ]"],
        ["Cartão CNPJ", "[ ]"],
        ["Certidão Federal (Receita/PGFN)", "[ ]"],
        ["Certidão Estadual", "[ ]"],
        ["Certidão Municipal", "[ ]"],
        ["Certificado FGTS", "[ ]"],
        ["CNDT (trabalhista)", "[ ]"],
        ["Licença sanitária / alvará de vigilância", "[ ]"],
        ["Registro/autorização ANVISA do produto (oxigênio medicinal)", "[ ]"],
        ["Atestado de capacidade técnica (se exigido/apresentado)", "[ ]"],
        ["Procuração / poderes do signatário", "[ ]"],
        ["Proposta / cotação assinada", "[ ]"],
    ]
    add_table(doc, ["Documento", "Situação"], rows)
    add_just(doc, "Observações: _______________________________________________________________")
    assinaturas(doc, [["Servidor responsável pela conferência", "Matrícula / Setor"]])
    save(doc, OUT1 / "08-checklist-habilitacao.docx")


def p1_minuta_contrato():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(
        doc,
        "Minuta de Contrato Emergencial",
        "Fornecimento de oxigênio medicinal — Art. 75, VIII, Lei nº 14.133/2021",
    )
    add_just(
        doc,
        f"CONTRATO Nº _____/2026 que entre si celebram o {DADOS['municipio']}, por intermédio da "
        f"{DADOS['sms']}, e a empresa ________________________, para fornecimento emergencial de "
        "oxigênio medicinal com comodato de cilindros.",
    )
    add_heading_local(doc, "CLÁUSULA PRIMEIRA — DO OBJETO")
    add_just(
        doc,
        "Fornecimento contínuo, sob demanda, de oxigênio medicinal gasoso (pureza mínima 99,5%), "
        "com cessão de cilindros em comodato, manutenção e logística de entrega nas UBSs e UBAIs "
        "do Município de Pelotas/RS, conforme Termo de Referência Emergencial anexo.",
    )
    add_heading_local(doc, "CLÁUSULA SEGUNDA — DO REGIME E DA FUNDAMENTAÇÃO")
    add_just(
        doc,
        f"Contratação por dispensa de licitação, art. 75, VIII, Lei nº 14.133/2021, em razão da "
        f"licitação deserta do Processo nº {DADOS['processo_origem']} e do risco de desabastecimento.",
    )
    add_heading_local(doc, "CLÁUSULA TERCEIRA — DO PREÇO E DO PAGAMENTO")
    add_just(
        doc,
        "O preço unitário é de R$ ________ por m³ efetivamente fornecido. "
        "Pagamento mensal por medição, após atesto, vedado pagamento por volume não entregue.",
    )
    add_heading_local(doc, "CLÁUSULA QUARTA — DA VIGÊNCIA")
    add_just(
        doc,
        f"O prazo de vigência é de {DADOS['prazo_emergencial']}, contados da assinatura, "
        "ou até a conclusão da licitação regular instaurada em paralelo, o que ocorrer primeiro. "
        "Não haverá prorrogação ordinária.",
    )
    add_heading_local(doc, "CLÁUSULA QUINTA — DA RESCISÃO AUTOMÁTICA")
    add_just(
        doc,
        "Este contrato extingue-se automaticamente com a homologação do Pregão Eletrônico e "
        "assinatura do contrato definitivo correspondente, assegurada transição sem interrupção "
        "do abastecimento, no prazo de até 10 (dez) dias úteis, salvo ajuste diverso entre as partes "
        "e a fiscalização.",
    )
    add_heading_local(doc, "CLÁUSULA SEXTA — DAS OBRIGAÇÕES")
    add_just(
        doc,
        "A CONTRATADA obriga-se a cumprir o Termo de Referência, manter habilitação, garantir "
        f"entrega em até {DADOS['entrega_urgencia']} (e {DADOS['entrega_critica']} em urgência), "
        "assistência 24h e conformidade sanitária. A CONTRATANTE obriga-se a pagar as medições "
        "e a exercer a fiscalização.",
    )
    add_heading_local(doc, "CLÁUSULA SÉTIMA — DA FISCALIZAÇÃO")
    add_just(
        doc,
        "Gestor: ________________. Fiscal: ________________. A fiscalização não exclui a "
        "responsabilidade da CONTRATADA.",
    )
    add_heading_local(doc, "CLÁUSULA OITAVA — DAS SANÇÕES")
    add_just(doc, "Aplicam-se as sanções dos arts. 155 a 163 da Lei nº 14.133/2021 e normas municipais.")
    add_heading_local(doc, "CLÁUSULA NONA — DO FORO")
    add_just(doc, "Fica eleito o foro da Comarca de Pelotas/RS.")
    doc.add_paragraph()
    add_just(doc, "E por estarem de acordo, assinam o presente.")
    assinaturas(
        doc,
        [
            [DADOS["municipio"] + " / " + DADOS["sms"], "CONTRATANTE"],
            ["Empresa _______________________", "CONTRATADA"],
            ["Testemunha 1", "CPF"],
            ["Testemunha 2", "CPF"],
        ],
    )
    save(doc, OUT1 / "09-minuta-contrato-emergencial.docx")


def p1_autorizacao():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Despacho de Autorização da Dispensa Emergencial")
    add_just(
        doc,
        f"Trata-se de pedido de contratação emergencial de oxigênio medicinal, em face da licitação "
        f"deserta do Processo nº {DADOS['processo_origem']}, com risco de desabastecimento das "
        "UBSs e UBAIs municipais.",
    )
    add_just(
        doc,
        "À vista do DFD, ETP simplificado, Termo de Referência, justificativa da emergência, "
        "pesquisa de preços, razão de escolha e minuta contratual, e considerando o art. 75, VIII, "
        "da Lei nº 14.133/2021,",
    )
    add_center(doc, "AUTORIZO", bold=True, size=14)
    add_just(
        doc,
        f"a dispensa de licitação e a contratação emergencial pelo prazo de {DADOS['prazo_emergencial']}, "
        "ou até a conclusão do Pregão Eletrônico paralelo, o que ocorrer primeiro, no valor estimado "
        "de R$ _______________, com publicação no PNCP e no Portal da Transparência, "
        "encaminhando-se à Procuradoria/Assessoria Jurídica para parecer, e aos setores competentes "
        "para empenho, assinatura e execução.",
    )
    assinaturas(doc, [[DADOS["secretaria"], "Secretária Municipal de Saúde"]])
    doc.add_paragraph()
    add_heading_local(doc, "Minuta de extrato para publicação (PNCP / Transparência)")
    add_just(
        doc,
        f"EXTRATO DE DISPENSA DE LICITAÇÃO — Art. 75, VIII, Lei nº 14.133/2021. "
        f"Órgão: {DADOS['sms']} — {DADOS['municipio']}/RS. "
        f"Processo: _____/2026 (origem {DADOS['processo_origem']}). "
        "Objeto: fornecimento emergencial de oxigênio medicinal com comodato de cilindros para UBSs/UBAIs. "
        f"Contratada: _____________ — CNPJ _____________. Vigência: {DADOS['prazo_emergencial_curto']} "
        "ou até homologação do pregão regular. Valor: R$ _____________. "
        "Data: ___/___/2026.",
    )
    save(doc, OUT1 / "10-despacho-autorizacao-e-extrato.docx")


def p1_parecer_juridico_roteiro():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Roteiro para Parecer Jurídico", "Elementos mínimos a verificar na dispensa emergencial")
    for item in [
        "Competência da autoridade para autorizar a dispensa;",
        "Enquadramento no art. 75, VIII, Lei nº 14.133/2021;",
        "Comprovação da licitação deserta e do risco de prejuízo/desabastecimento;",
        "Prazo estritamente necessário e vedação de prorrogação indevida;",
        "Pesquisa de preços e justificativa do valor;",
        "Habilitação da contratada (jurídica, fiscal, sanitária);",
        "Cláusulas essenciais da minuta (objeto, preço, vigência, rescisão automática, sanções);",
        "Publicidade no PNCP;",
        "Existência de processo paralelo de pregão regular (boa prática / mitigação).",
    ]:
        add_left(doc, f"• {item}", indent=0.5)
    add_just(
        doc,
        "Espaço para parecer formal da Assessoria Jurídica / Procuradoria: "
        "_______________________________________________________________ "
        "_______________________________________________________________.",
    )
    save(doc, OUT1 / "11-roteiro-parecer-juridico.docx")


# ---------------------------------------------------------------------------
# PROCESSO 2
# ---------------------------------------------------------------------------

def p2_capa():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Capa do Processo", "Processo 2 — Pregão Eletrônico (solução definitiva em paralelo)")
    add_left(doc, "Processo administrativo nº: _______________ / 2026")
    add_left(doc, f"Processo de origem / referência: nº {DADOS['processo_origem']} (licitação deserta)")
    add_left(doc, f"Unidade demandante: {DADOS['sms']} — {DADOS['unidade']}")
    add_left(doc, "Modalidade: Pregão Eletrônico — Lei nº 14.133/2021")
    add_left(doc, "Critério de julgamento: Menor preço por lote")
    doc.add_paragraph()
    add_heading_local(doc, "1. Objeto")
    add_just(
        doc,
        "Registro de preços ou contrato para fornecimento contínuo, sob demanda, de Oxigênio Medicinal "
        "gasoso, com comodato de cilindros, manutenção e logística, para UBSs e UBAIs das zonas "
        "urbana e rural do Município de Pelotas/RS, pelo prazo de 12 (doze) meses, prorrogável "
        "nos termos da lei.",
    )
    add_heading_local(doc, "2. Estratégia")
    add_just(
        doc,
        "Repetição do certame (Opção A), em paralelo à dispensa emergencial (Opção C), com revisão "
        "do Termo de Referência, atualização da pesquisa de preços, racionalização de exigências "
        "de habilitação e ampliação da divulgação junto a empresas de gases medicinais.",
    )
    add_heading_local(doc, "3. Quantitativo anual de referência")
    add_just(
        doc,
        f"Zona urbana: {DADOS['consumo_urbano_ano']}; zona rural: {DADOS['consumo_rural_ano']}; "
        f"total: {DADOS['consumo_total_ano']} (estimativa sem consumo mínimo).",
    )
    add_heading_local(doc, "4. Dotação")
    add_left(doc, f"Projeto/Atividade: {DADOS['projeto_atividade']}")
    add_left(doc, f"Fonte: {DADOS['fonte']}")
    add_left(doc, "Valor estimado: R$ _______________ (pesquisa atualizada)")
    assinaturas(
        doc,
        [
            [DADOS["diretora"], "Diretora de Atenção Primária"],
            [DADOS["secretaria"], "Secretária Municipal de Saúde"],
        ],
    )
    save(doc, OUT2 / "01-capa-processo-pregao.docx")


def p2_dfd():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Documento de Formalização da Demanda — DFD", "Versão revisada para novo Pregão Eletrônico")
    add_left(doc, f"Unidade demandante: {DADOS['sms']} — {DADOS['unidade']}")
    add_left(doc, f"Referência: Processo nº {DADOS['processo_origem']} (deserto) — republicação revisada")
    doc.add_paragraph()
    add_heading_local(doc, "1. Objeto")
    add_just(
        doc,
        "Contratação de empresa especializada para fornecimento contínuo, sob demanda, de Gases "
        "Medicinais (Oxigênio Medicinal), com cessão de cilindros em regime de comodato, para "
        "atender às demandas das UBSs e UBAIs, zonas urbana e rural de Pelotas/RS.",
    )
    add_left(doc, "Tipo do objeto: ( x ) Serviço continuado sem dedicação exclusiva de mão de obra")
    add_heading_local(doc, "2. Forma de contratação sugerida")
    add_just(
        doc,
        "Pregão Eletrônico, critério de MENOR PREÇO por lote, nos termos do art. 6º, XLI, e art. 29, "
        "parágrafo único, da Lei nº 14.133/2021.",
    )
    add_heading_local(doc, "3. Justificativa")
    add_just(
        doc,
        "Assegurar abastecimento contínuo de oxigênio medicinal, insumo essencial à assistência na "
        "Atenção Primária. O certame anterior restou deserto; esta DFD revisada corrige pontos de "
        "atratividade (preço de referência atualizado, prazos realistas, habilitação proporcional e "
        "divulgação ampliada), mantendo a divisão em dois lotes (urbano e rural) para ampliar "
        "competitividade e refletir diferenças logísticas.",
    )
    add_heading_local(doc, "4. Quantidade e descrição")
    add_table(
        doc,
        ["Item", "Especificação", "Unidade", "Quantidade anual"],
        [
            [
                "1.1 — Lote 1",
                "Recarga/abastecimento de oxigênio medicinal gasoso — Zona Urbana "
                "(38 UBSs + 3 UBAIs). Média: 4,0 m³/dia UBSs (152 m³/mês) + 10,0 m³/dia UBAIs (30 m³/mês).",
                "m³",
                "2.184,0",
            ],
            [
                "2.1 — Lote 2",
                "Recarga/abastecimento de oxigênio medicinal gasoso — Zona Rural (12 UBSs). "
                "Média: 2,0 m³/dia (24 m³/mês).",
                "m³",
                "288,0",
            ],
        ],
    )
    add_just(
        doc,
        "Fornecimento sob demanda; faturamento pelo volume efetivamente fornecido; "
        "estimativas referenciais, sem obrigação de consumo mínimo.",
    )
    add_heading_local(doc, "5. Estimativa de valor")
    add_just(doc, "Conforme pesquisa de preços ATUALIZADA (anexo), nos termos do art. 23 da Lei nº 14.133/2021.")
    add_heading_local(doc, "6. Dotação orçamentária")
    add_left(doc, f"Projeto/Atividade: {DADOS['projeto_atividade']}")
    add_left(doc, f"Fonte: {DADOS['fonte']}")
    add_left(doc, "Natureza da despesa: _______________")
    add_left(doc, "Valor estimado: R$ _______________")
    add_heading_local(doc, "7. Locais de entrega")
    add_just(doc, "Conforme Anexo de Locais (Lote 1 — 41 pontos; Lote 2 — 12 pontos rurais).")
    add_heading_local(doc, "8. Previsão de início")
    add_just(
        doc,
        "Início em até 10 (dez) dias úteis da emissão da Nota de Empenho / assinatura do contrato, "
        "prazo mais realista que o anteriormente adotado, visando atratividade do certame.",
    )
    assinaturas(
        doc,
        [
            [DADOS["diretora"], "Diretora de Atenção Primária"],
            [DADOS["secretaria"], "Secretária Municipal de Saúde — Autorização"],
        ],
    )
    save(doc, OUT2 / "02-dfd-revisado-pregao.docx")


def p2_etp():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Estudo Técnico Preliminar — ETP Completo", "Novo Pregão Eletrônico — oxigênio medicinal")
    secoes = [
        (
            "1. Descrição da necessidade",
            "Abastecimento contínuo de oxigênio medicinal nas UBSs/UBAIs urbanas e rurais, "
            f"após desertificação do Processo nº {DADOS['processo_origem']}.",
        ),
        (
            "2. Área requisitante",
            f"{DADOS['sms']} / {DADOS['unidade']}.",
        ),
        (
            "3. Requisitos da contratação",
            "Oxigênio ≥ 99,5%; comodato de cilindros; logística urbana/rural; manutenção; "
            "conformidade ANVISA/ABNT; atendimento emergencial; medição por m³.",
        ),
        (
            "4. Estimativa de quantidades",
            f"Anual: {DADOS['consumo_total_ano']} (Lote 1: {DADOS['consumo_urbano_ano']}; "
            f"Lote 2: {DADOS['consumo_rural_ano']}).",
        ),
        (
            "5. Levantamento de mercado",
            "Nova pesquisa com painéis oficiais, contratações similares de municípios da região "
            "e cotações diretas a distribuidores de gases medicinais. Registrar se o preço anterior "
            "estava defasado.",
        ),
        (
            "6. Descrição da solução como um todo",
            "Pregão eletrônico em dois lotes (urbano/rural), contrato de 12 meses prorrogável, "
            "pagamento por consumo efetivo, com comodato e manutenção inclusos no preço do m³ "
            "ou estrutura claramente discriminada sem ônus de locação para a Administração.",
        ),
        (
            "7. Estimativa de valor",
            "R$ _______________ — mapa atualizado anexo.",
        ),
        (
            "8. Justificativa para o parcelamento",
            "Dois lotes ampliam competitividade e refletem custos logísticos distintos da zona rural.",
        ),
        (
            "9. Contratações correlatas/interdependentes",
            "Dispensa emergencial em paralelo (Processo 1), com rescisão automática quando este "
            "pregão for homologado.",
        ),
        (
            "10. Alinhamento ao planejamento",
            f"Dotação {DADOS['projeto_atividade']}, fonte {DADOS['fonte']}.",
        ),
        (
            "11. Resultados pretendidos",
            "Segurança assistencial, continuidade do serviço, economicidade e ampliação de "
            "interessados no certame.",
        ),
        (
            "12. Providências prévias",
            "Revisão de TR; ofícios a empresas do ramo; publicação PNCP; prazo mínimo de 8 dias "
            "úteis para propostas.",
        ),
        (
            "13. Impactos ambientais",
            "Mínimos e controlados; responsabilidade da contratada quanto a cilindros, transporte "
            "e descarte conforme normas.",
        ),
        (
            "14. Viabilidade da contratação",
            "Solução viável técnica, administrativa e orçamentariamente, essencial e urgente.",
        ),
        (
            "15. Posicionamento conclusivo",
            "Recomenda-se a instauração do Pregão Eletrônico com os ajustes deste ETP/TR.",
        ),
    ]
    for titulo, texto in secoes:
        add_heading_local(doc, titulo)
        add_just(doc, texto)
    assinaturas(doc, [[DADOS["diretora"], "Diretora de Atenção Primária"]])
    save(doc, OUT2 / "03-etp-completo.docx")


def p2_tr():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Termo de Referência Revisado", "Pregão Eletrônico — Oxigênio Medicinal")
    add_heading_local(doc, "1. Objeto")
    add_just(
        doc,
        "Contratação de empresa(s) especializada(s) para fornecimento contínuo, sob demanda, de "
        "Oxigênio Medicinal gasoso, com comodato de cilindros, manutenção e logística de entrega "
        "nas UBSs e UBAIs de Pelotas/RS, em 2 (dois) lotes (zona urbana e zona rural).",
    )
    add_heading_local(doc, "2. Fundamentação")
    add_just(
        doc,
        f"Lei nº 14.133/2021; DFD e ETP deste processo; revisão motivada pela desertificação do "
        f"Processo nº {DADOS['processo_origem']}.",
    )
    add_heading_local(doc, "3. Pontos obrigatoriamente revisados em relação ao certame deserto")
    for item in [
        "Pesquisa de preços NOVA, com valores atuais do mercado local/regional;",
        "Prazo de início/entrega mais realista (até 10 dias úteis para início; 24h para reposição ordinária);",
        "Exigências de habilitação proporcionais — remover exigências excessivas que possam ter afastado empresas;",
        "Previsão expressa de comodato, manutenção e troca de cilindros;",
        "Prazo contratual de 12 meses, prorrogável nos termos da Lei nº 14.133/2021;",
        "Divulgação ampliada a empresas do ramo, além do PNCP.",
    ]:
        add_left(doc, f"• {item}", indent=0.5)
    add_heading_local(doc, "4. Especificação")
    add_just(
        doc,
        "Oxigênio medicinal gasoso, pureza mínima 99,5%; cilindros em comodato; conformidade "
        "ANVISA/ABNT; assistência 24h; troca imediata de recipientes irregulares.",
    )
    add_heading_local(doc, "5. Lotes e quantitativos")
    add_table(
        doc,
        ["Lote", "Abrangência", "Qtde. anual (m³)"],
        [
            ["1", "Zona Urbana — 38 UBSs + 3 UBAIs", "2.184,0"],
            ["2", "Zona Rural — 12 UBSs", "288,0"],
        ],
    )
    add_heading_local(doc, "6. Critério de julgamento")
    add_just(doc, "Menor preço por lote, pregão eletrônico.")
    add_heading_local(doc, "7. Habilitação (proporcional)")
    for item in [
        "Habilitação jurídica;",
        "Regularidade fiscal e trabalhista;",
        "Licença sanitária e registro ANVISA do oxigênio medicinal;",
        "Atestado de capacidade técnica compatível (fornecimento de gases medicinais), sem exigências desarrazoadas de quantitativos históricos;",
        "Não exigir certidões/documentos não previstos em lei ou desproporcionais ao objeto.",
    ]:
        add_left(doc, f"• {item}", indent=0.5)
    add_heading_local(doc, "8. Prazo de vigência")
    add_just(doc, "12 (doze) meses, prorrogável na forma da Lei nº 14.133/2021.")
    add_heading_local(doc, "9. Pagamento")
    add_just(doc, "Mensal, por medição do consumo efetivo (m³), com atesto da fiscalização.")
    add_heading_local(doc, "10. Fiscalização")
    add_left(doc, "Gestor: _____________________  Fiscal: _____________________")
    add_heading_local(doc, "11. Sanções")
    add_just(doc, "Conforme Lei nº 14.133/2021 e edital.")
    assinaturas(doc, [[DADOS["diretora"], "Diretora de Atenção Primária"]])
    save(doc, OUT2 / "04-termo-de-referencia-revisado.docx")


def p2_pesquisa():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Pesquisa de Preços Atualizada", "Novo Pregão — Oxigênio Medicinal")
    add_just(
        doc,
        "Atualização motivada pela licitação deserta. Fontes: Painel de Preços, contratações "
        "similares de municípios, cotações diretas. Justificar eventual defasagem do preço anterior.",
    )
    add_table(
        doc,
        ["Fonte", "Data", "R$/m³ Lote 1", "R$/m³ Lote 2", "Link/doc", "Obs."],
        [
            ["Painel de Preços", "", "", "", "", ""],
            ["Município similar 1", "", "", "", "", ""],
            ["Município similar 2", "", "", "", "", ""],
            ["Cotação empresa A", "", "", "", "", ""],
            ["Cotação empresa B", "", "", "", "", ""],
            ["Cotação empresa C", "", "", "", "", ""],
        ],
    )
    add_left(doc, "Preço de referência Lote 1: R$ ________ / m³")
    add_left(doc, "Preço de referência Lote 2: R$ ________ / m³")
    add_left(doc, "Valor anual estimado total: R$ _______________")
    add_just(
        doc,
        "Justificativa de eventual diferença em relação ao edital anterior: "
        "_______________________________________________________________.",
    )
    assinaturas(doc, [["Responsável pela pesquisa", "Setor de Licitações — SMS"]])
    save(doc, OUT2 / "05-pesquisa-de-precos-atualizada.docx")


def p2_edital():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Minuta de Edital — Pregão Eletrônico", "Fornecimento de oxigênio medicinal")
    add_left(doc, "Pregão Eletrônico nº _____/2026")
    add_left(doc, f"Processo administrativo nº _____/2026 (ref. {DADOS['processo_origem']})")
    doc.add_paragraph()
    add_heading_local(doc, "1. Preâmbulo")
    add_just(
        doc,
        f"O {DADOS['municipio']}, por intermédio da {DADOS['sms']}, torna público que realizará "
        "Pregão Eletrônico, na forma da Lei nº 14.133/2021 e normas municipais, para contratação "
        "do objeto abaixo.",
    )
    add_heading_local(doc, "2. Objeto")
    add_just(
        doc,
        "Fornecimento contínuo de oxigênio medicinal com comodato de cilindros, em 2 lotes "
        "(zona urbana e zona rural), conforme Anexo I — Termo de Referência.",
    )
    add_heading_local(doc, "3. Data e horário")
    add_left(doc, "Início de acolhimento de propostas: ___/___/2026 às __:__")
    add_left(doc, "Limite para recebimento: ___/___/2026 às __:__")
    add_left(doc, "Início da sessão pública: ___/___/2026 às __:__")
    add_left(doc, "Plataforma: ________________ (PNCP / sistema adotado pelo Município)")
    add_heading_local(doc, "4. Critério de julgamento")
    add_just(doc, "Menor preço por lote.")
    add_heading_local(doc, "5. Prazo mínimo de publicidade")
    add_just(doc, "Observar o mínimo legal de 8 (oito) dias úteis para apresentação de propostas.")
    add_heading_local(doc, "6. Anexos")
    for item in [
        "Anexo I — Termo de Referência;",
        "Anexo II — Minuta do Contrato;",
        "Anexo III — Modelo de Proposta;",
        "Anexo IV — Locais de Entrega;",
        "Anexo V — Modelo de Atestado / declarações (se houver).",
    ]:
        add_left(doc, f"• {item}", indent=0.5)
    add_heading_local(doc, "7. Divulgação ampliada")
    add_just(
        doc,
        "Além do PNCP e do Portal do Município, deverão ser encaminhados ofícios/e-mails a "
        "empresas do ramo de gases medicinais da região, com comprovação nos autos.",
    )
    add_heading_local(doc, "8. Disposições finais")
    add_just(
        doc,
        "Integram o edital, independentemente de transcrição, a Lei nº 14.133/2021 e o Termo de "
        "Referência. Esclarecimentos: setor de licitações — e-mail/telefone _______________.",
    )
    assinaturas(doc, [["Autoridade competente / Pregoeiro(a)", DADOS["sms"]]])
    save(doc, OUT2 / "06-minuta-edital-pregao.docx")


def p2_minuta_contrato():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Anexo II — Minuta do Contrato", "Pregão Eletrônico — Oxigênio Medicinal")
    add_just(
        doc,
        f"CONTRATO Nº _____/2026 — {DADOS['municipio']} / {DADOS['sms']} e a empresa "
        "____________________, vencedora do Pregão Eletrônico nº _____/2026, Lote(s) _____.",
    )
    add_heading_local(doc, "CLÁUSULA 1 — OBJETO")
    add_just(
        doc,
        "Fornecimento contínuo, sob demanda, de oxigênio medicinal gasoso com comodato de "
        "cilindros, conforme TR e proposta vencedora.",
    )
    add_heading_local(doc, "CLÁUSULA 2 — VIGÊNCIA")
    add_just(doc, "12 (doze) meses, prorrogável na forma da Lei nº 14.133/2021.")
    add_heading_local(doc, "CLÁUSULA 3 — PREÇO E PAGAMENTO")
    add_just(
        doc,
        "Preço unitário de R$ ______ / m³ (Lote ____). Pagamento mensal por medição efetivamente "
        "executada, após atesto.",
    )
    add_heading_local(doc, "CLÁUSULA 4 — OBRIGAÇÕES")
    add_just(
        doc,
        "A CONTRATADA cumprirá o TR (logística, comodato, manutenção, prazos, ANVISA). "
        "A CONTRATANTE pagará e fiscalizará.",
    )
    add_heading_local(doc, "CLÁUSULA 5 — FISCALIZAÇÃO")
    add_just(doc, "Gestor: ________. Fiscal: ________.")
    add_heading_local(doc, "CLÁUSULA 6 — SANÇÕES")
    add_just(doc, "Arts. 155 a 163 da Lei nº 14.133/2021 e edital.")
    add_heading_local(doc, "CLÁUSULA 7 — RESCISÃO")
    add_just(doc, "Nos casos legais, assegurado contraditório e ampla defesa.")
    add_heading_local(doc, "CLÁUSULA 8 — FORO")
    add_just(doc, "Comarca de Pelotas/RS.")
    assinaturas(
        doc,
        [
            ["CONTRATANTE — SMS", DADOS["municipio"]],
            ["CONTRATADA", "Empresa"],
            ["Testemunhas", ""],
        ],
    )
    save(doc, OUT2 / "07-anexo-ii-minuta-contrato.docx")


def p2_modelo_proposta():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Anexo III — Modelo de Proposta de Preços")
    add_left(doc, "Pregão Eletrônico nº _____/2026")
    add_left(doc, "Razão social: _____________________ CNPJ: ______________")
    add_left(doc, "Endereço: ______________________________________________")
    add_left(doc, "E-mail / telefone: ______________________________________")
    doc.add_paragraph()
    add_table(
        doc,
        ["Lote", "Descrição", "Unidade", "Qtde. anual", "Preço unitário (R$)", "Preço total (R$)"],
        [
            ["1", "Oxigênio medicinal — Zona Urbana", "m³", "2.184,0", "", ""],
            ["2", "Oxigênio medicinal — Zona Rural", "m³", "288,0", "", ""],
        ],
    )
    add_just(
        doc,
        "Declaramos que nos preços estão inclusos comodato de cilindros, manutenção, logística, "
        "tributos e demais encargos. Validade da proposta: 60 (sessenta) dias. "
        "Prazo de início: até 10 dias úteis da ordem de serviço/empenho.",
    )
    assinaturas(doc, [["Representante legal", "Nome / CPF / cargo"]])
    save(doc, OUT2 / "08-anexo-iii-modelo-de-proposta.docx")


def p2_divulgacao():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Ofício-Circular — Divulgação Ampliada do Pregão")
    add_left(doc, "Ofício nº _____/2026 — SMS/Licitações")
    add_left(doc, "Pelotas/RS, _____ de ______________ de 2026.")
    doc.add_paragraph()
    add_left(doc, "À empresa: _______________________________")
    add_left(doc, "E-mail: ___________________________________")
    doc.add_paragraph()
    add_just(doc, "Prezados(as),")
    add_just(
        doc,
        f"Informamos que o {DADOS['municipio']}, por meio da {DADOS['sms']}, publicará/publicou "
        "Pregão Eletrônico para fornecimento de oxigênio medicinal com comodato de cilindros "
        "para UBSs e UBAIs (zonas urbana e rural).",
    )
    add_just(
        doc,
        "O certame anterior (Processo nº "
        f"{DADOS['processo_origem']}) restou deserto. Visando ampliar a competitividade, "
        "convidamos V.Sas. a participar do novo certame, cujos documentos estarão disponíveis "
        "no PNCP e no Portal da Prefeitura.",
    )
    add_left(doc, "Link PNCP / edital: ________________________________")
    add_left(doc, "Data da sessão: ___/___/2026")
    add_left(doc, "Contato: ________________________________")
    add_just(doc, "Atenciosamente,")
    assinaturas(doc, [["Setor de Licitações / Pregoeiro(a)", DADOS["sms"]]])
    doc.add_paragraph()
    add_heading_local(doc, "Controle de envio (anexar comprovantes)")
    add_table(
        doc,
        ["Empresa", "E-mail/WhatsApp", "Data envio", "Comprovante nos autos"],
        [
            ["", "", "", "[ ]"],
            ["", "", "", "[ ]"],
            ["", "", "", "[ ]"],
            ["", "", "", "[ ]"],
            ["", "", "", "[ ]"],
        ],
    )
    save(doc, OUT2 / "09-oficio-divulgacao-ampliada.docx")


def anexos_locais():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(doc, "Anexo — Locais de Entrega", "UBSs e UBAIs — Zona Urbana e Zona Rural")
    add_heading_local(doc, "LOTE 1 — Zona Urbana")
    add_table(
        doc,
        ["Nº", "Unidade", "Endereço", "Telefone"],
        [list(r) for r in UNIDADES_URBANAS],
    )
    add_heading_local(doc, "LOTE 2 — Zona Rural")
    add_table(
        doc,
        ["Nº", "Unidade", "Endereço", "Distância estimada", "Telefone"],
        [list(r) for r in UNIDADES_RURAIS],
    )
    add_just(
        doc,
        "As distâncias da zona rural são referenciais (trajeto até a sede da SMS) e destinam-se "
        "apenas ao planejamento logístico.",
    )
    save(doc, OUTA / "anexo-locais-de-entrega-ubs-ubai.docx")


def cronograma():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(
        doc,
        "Cronograma Paralelo — Processos 1 e 2",
        "Dispensa emergencial + Pregão eletrônico (oxigênio medicinal)",
    )
    add_just(
        doc,
        f"Base: Processo nº {DADOS['processo_origem']} (deserto em {DADOS['data_deserto']}). "
        "Objetivo: não interromper o abastecimento (Processo 1) e resolver em definitivo (Processo 2).",
    )
    add_table(
        doc,
        ["Semana", "Processo 1 — Dispensa Emergencial (Opção C)", "Processo 2 — Pregão Regular (Opção A)"],
        [
            [
                "Semana 1",
                "Cotações + TR emergencial + justificativa + jurídico",
                "Revisão do TR + nova pesquisa de preços + ETP completo",
            ],
            [
                "Semana 2",
                "Autorização + publicação PNCP + assinatura + início do fornecimento",
                "Finalização do edital + divulgação ampliada + publicação",
            ],
            [
                "Semana 3",
                "Fornecimento normal + fiscalização de estoque",
                "Prazo de propostas / sessão pública do Pregão",
            ],
            [
                "Semana 4",
                "Acompanhar consumo e medições",
                "Habilitação / adjudicação / homologação",
            ],
            [
                "Após homologar",
                "Rescindir a Dispensa C (transição assistida)",
                "Assinar contrato definitivo (12 meses)",
            ],
        ],
    )
    add_heading_local(doc, "Checklist integrado de urgência")
    for item in [
        "Preencher estoque atual em dias na Justificativa da Emergência;",
        "Coletar no mínimo 3 cotações (ou justificar);",
        "Indicar gestor e fiscal nos dois processos;",
        "Publicar dispensa no PNCP;",
        "Enviar ofícios a empresas de gases medicinais;",
        "Manter ambos os autos tramitando simultaneamente.",
    ]:
        add_left(doc, f"☐ {item}", indent=0.25)
    assinaturas(
        doc,
        [
            [DADOS["diretora"], "Diretoria de Atenção Primária"],
            ["Chefia de Compras / Licitações", DADOS["sms"]],
        ],
    )
    save(doc, ROOT / "docs" / "00-cronograma-paralelo.docx")


def despacho_encaminhamento():
    doc = Document()
    set_doc_defaults(doc)
    cabecalho(
        doc,
        "Despacho de Encaminhamento à Diretoria Administrativa / Compras / Licitações",
        f"Retorno do Processo nº {DADOS['processo_origem']}",
    )
    add_just(doc, "A/C Ana Cristina Porto — Chefe do Setor de Compras SMS / Setor de Licitações")
    add_just(
        doc,
        f"Em atenção à tramitação de {DADOS['data_deserto']}, que comunicou a desertificação do "
        f"Processo nº {DADOS['processo_origem']} e apresentou as Opções A, B e C, a Diretoria de "
        "Atenção Primária manifesta a seguinte decisão estratégica, diante da urgência do "
        "abastecimento de oxigênio medicinal:",
    )
    add_just(
        doc,
        "Adotar SIMULTANEAMENTE a Opção C (contratação emergencial — art. 75, VIII) e a Opção A "
        "(repetição do certame via Pregão Eletrônico com TR e pesquisa revisados), conforme "
        "pacote documental anexo.",
    )
    add_just(
        doc,
        "Solicita-se o imediato andamento dos dois expedientes em paralelo, com prioridade, "
        "orientação jurídica e publicação no PNCP, a fim de evitar desabastecimento das UBSs e UBAIs.",
    )
    add_heading_local(doc, "Documentos anexos — Processo 1 (Emergencial)")
    add_just(
        doc,
        "Capa; DFD emergencial; ETP simplificado; TR emergencial; Justificativa da emergência; "
        "Mapa de preços; Razão da escolha; Checklist de habilitação; Minuta contratual; "
        "Despacho de autorização/extrato; Roteiro de parecer; Locais de entrega; Cronograma.",
    )
    add_heading_local(doc, "Documentos anexos — Processo 2 (Pregão)")
    add_just(
        doc,
        "Capa; DFD revisado; ETP completo; TR revisado; Pesquisa atualizada; Minuta de edital; "
        "Minuta de contrato; Modelo de proposta; Ofício de divulgação ampliada; Locais de entrega.",
    )
    assinaturas(
        doc,
        [
            [DADOS["solicitante"], "Diretoria de Atenção Primária"],
            [DADOS["diretora"], "Diretora de Atenção Primária"],
        ],
    )
    save(doc, ROOT / "docs" / "00-despacho-encaminhamento-opcao-a-e-c.docx")


def main():
    for d in (OUT1, OUT2, OUTA):
        d.mkdir(parents=True, exist_ok=True)
    p1_capa()
    p1_dfd()
    p1_etp()
    p1_tr()
    p1_justificativa()
    p1_pesquisa_precos()
    p1_razao_escolha()
    p1_habilitacao()
    p1_minuta_contrato()
    p1_autorizacao()
    p1_parecer_juridico_roteiro()
    p2_capa()
    p2_dfd()
    p2_etp()
    p2_tr()
    p2_pesquisa()
    p2_edital()
    p2_minuta_contrato()
    p2_modelo_proposta()
    p2_divulgacao()
    anexos_locais()
    cronograma()
    despacho_encaminhamento()
    print("\nPacote documental gerado com sucesso.")


if __name__ == "__main__":
    main()
