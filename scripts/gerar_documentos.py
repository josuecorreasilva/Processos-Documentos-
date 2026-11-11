#!/usr/bin/env python3
"""Gera o pacote documental oficial SMS Pelotas — Processos 1 e 2 em paralelo."""

from __future__ import annotations

import shutil
from pathlib import Path

from oficial_template import (
    ROOT,
    boxed,
    bullet,
    checkbox_line,
    heading,
    meta_line,
    new_official_doc,
    para,
    save,
    signatures,
    simple_table,
    title,
)

DOCS = ROOT / "docs"
P0 = DOCS / "00-INSTRUCAO-GERAL"
P1 = DOCS / "01-PROCESSO-DISPENSA-EMERGENCIAL"
P2 = DOCS / "02-PROCESSO-PREGAO-ELETRONICO"
ANX = DOCS / "03-ANEXOS-COMUNS"

D = {
    "processo_origem": "1242/2026",
    "sl": "00710/2026",
    "memo": "003746/2026",
    "memo_ccgof": "005720/2026",
    "abertura": "27/01/2026",
    "data_deserto": "29/10/2026",
    "projeto": "10.301.0105.2167.00 — Atenção Primária em Saúde",
    "fonte": "1600.000",
    "conta": "Agência 2820 / C/C 2200/575835306-6 — Caixa Econômica Federal",
    "endereco": "Rua Lobo da Costa, nº 1764, Pelotas/RS",
    "contato": "Carol — (53) 98462-4786",
    "diretora": "Cândida Silveira Rodrigues",
    "secretaria": "Angela Moreira Vitória",
    "solicitante": "Josué Silva",
    "prazo_emerg": "180 (cento e oitenta) dias",
    "prazo_emerg_curto": "180 dias",
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


def reset_tree():
    for old in [
        DOCS / "processo-1-dispensa-emergencial",
        DOCS / "processo-2-pregao-regular",
        DOCS / "anexos",
        DOCS / "00-cronograma-paralelo.docx",
        DOCS / "00-despacho-encaminhamento-opcao-a-e-c.docx",
    ]:
        if old.is_dir():
            shutil.rmtree(old)
        elif old.exists():
            old.unlink()
    for d in (P0, P1, P2, ANX):
        d.mkdir(parents=True, exist_ok=True)


# =============================================================================
# 00 — INSTRUÇÃO GERAL
# =============================================================================

def doc_indice():
    doc = new_official_doc()
    title(doc, "Índice dos Autos", "Oxigênio Medicinal — Processos paralelos (Opções A e C)")
    meta_line(doc, "Processo de origem", D["processo_origem"])
    meta_line(doc, "Objeto", "Fornecimento de oxigênio medicinal com comodato de cilindros — UBSs e UBAIs")
    para(
        doc,
        "Este índice organiza a instrução simultânea de dois expedientes administrativos, "
        "em resposta à desertificação do certame e ao risco de desabastecimento.",
        first_indent=False,
    )
    heading(doc, "VOLUME 0 — Instrução geral")
    simple_table(
        doc,
        ["Peça", "Documento"],
        [
            ["00.01", "Despacho de encaminhamento — adoção das Opções A e C"],
            ["00.02", "Cronograma paralelo de execução"],
            ["00.03", "Índice dos autos (este documento)"],
        ],
    )
    heading(doc, "VOLUME 1 — Processo de Dispensa Emergencial (Art. 75, VIII)")
    simple_table(
        doc,
        ["Peça", "Documento"],
        [
            ["01.01", "Capa / autuação do processo emergencial"],
            ["01.02", "Documento de Formalização da Demanda — DFD emergencial"],
            ["01.03", "Estudo Técnico Preliminar — ETP simplificado"],
            ["01.04", "Mapa de Gerenciamento de Riscos"],
            ["01.05", "Termo de Referência Emergencial"],
            ["01.06", "Justificativa da emergência e do enquadramento legal"],
            ["01.07", "Mapa de pesquisa de preços"],
            ["01.08", "Razão da escolha do fornecedor e justificativa do preço"],
            ["01.09", "Checklist de habilitação"],
            ["01.10", "Declarações do fornecedor"],
            ["01.11", "Minuta do contrato emergencial"],
            ["01.12", "Solicitação de parecer jurídico"],
            ["01.13", "Despacho de autorização da autoridade competente"],
            ["01.14", "Aviso / extrato de dispensa para PNCP"],
            ["01.15", "Extrato de contrato para publicação"],
        ],
    )
    heading(doc, "VOLUME 2 — Processo de Pregão Eletrônico (solução definitiva)")
    simple_table(
        doc,
        ["Peça", "Documento"],
        [
            ["02.01", "Capa / autuação do pregão"],
            ["02.02", "DFD revisado"],
            ["02.03", "ETP completo"],
            ["02.04", "Mapa de Gerenciamento de Riscos"],
            ["02.05", "Termo de Referência revisado"],
            ["02.06", "Pesquisa de preços atualizada"],
            ["02.07", "Minuta de Edital"],
            ["02.08", "Anexo II — Minuta do Contrato"],
            ["02.09", "Anexo III — Modelo de Proposta"],
            ["02.10", "Ofício-circular de divulgação ampliada"],
            ["02.11", "Despacho de autorização para publicação do edital"],
        ],
    )
    heading(doc, "VOLUME 3 — Anexos comuns")
    simple_table(
        doc,
        ["Peça", "Documento"],
        [
            ["03.01", "Locais de entrega — Zona Urbana e Zona Rural"],
            ["03.02", "Relatório de estoque atual (a anexar pela unidade)"],
            ["03.03", "Ata / despacho da licitação deserta (autos do 1242/2026)"],
        ],
    )
    signatures(doc, [[D["diretora"], "Diretora de Atenção Primária"]])
    save(doc, P0 / "00.03-indice-dos-autos.docx")


def doc_despacho_geral():
    doc = new_official_doc()
    title(doc, "Despacho de Encaminhamento", f"Processo nº {D['processo_origem']} — Oxigênio Medicinal")
    boxed(
        doc,
        "DESTINO",
        [
            "A/C Ana Cristina Porto — Chefe do Setor de Compras da SMS",
            "A/C Setor de Licitações da SMS",
            "Com ciência da Diretoria Administrativa e Financeira",
        ],
    )
    para(
        doc,
        f"Em atenção à tramitação de {D['data_deserto']}, que comunicou haver restado deserto o "
        f"Processo nº {D['processo_origem']} (SL nº {D['sl']}; Memorando nº {D['memo']}; "
        f"Memorando CCGOF nº {D['memo_ccgof']}) e apresentou as Opções A, B e C, a Diretoria de "
        "Atenção Primária, diante da urgência do abastecimento de oxigênio medicinal nas UBSs e "
        "UBAIs do Município, decide:",
    )
    boxed(
        doc,
        "DECISÃO ESTRATÉGICA",
        [
            "Adotar SIMULTANEAMENTE:",
            "• Opção C — Contratação emergencial por dispensa de licitação (Art. 75, VIII, Lei nº 14.133/2021), "
            f"pelo prazo estritamente necessário de {D['prazo_emerg_curto']}, ou até a homologação do pregão, "
            "o que ocorrer primeiro; e",
            "• Opção A — Repetição do certame, na modalidade Pregão Eletrônico, com revisão do Termo de "
            "Referência, atualização da pesquisa de preços, racionalização das exigências de habilitação e "
            "ampliação da divulgação junto a empresas do ramo de gases medicinais.",
        ],
    )
    para(
        doc,
        "A opção isolada de repetir o certame, sem medida emergencial, é incompatível com o risco "
        "assistencial. A opção isolada de contratação direta, sem novo pregão, não resolve a necessidade "
        "de forma definitiva. Por isso, os dois expedientes devem tramitar em paralelo, com prioridade.",
    )
    heading(doc, "Documentação anexa")
    para(
        doc,
        "Segue o pacote completo dos Volumes 0, 1, 2 e 3, pronto para preenchimento dos campos "
        "variáveis (estoque em dias, cotações, gestor/fiscal, dados da contratada) e tramitação.",
        first_indent=False,
    )
    para(
        doc,
        "Solicita-se o imediato andamento, com encaminhamento à Assessoria Jurídica para parecer "
        "quanto à dispensa emergencial e providências de publicação no PNCP e Portal da Transparência.",
    )
    signatures(
        doc,
        [
            [D["solicitante"], "Servidor responsável — Diretoria de Atenção Primária"],
            [D["diretora"], "Diretora de Atenção Primária"],
            [D["secretaria"], "Secretária Municipal de Saúde — Ciência/De acordo"],
        ],
    )
    save(doc, P0 / "00.01-despacho-encaminhamento-opcoes-a-e-c.docx")


def doc_cronograma():
    doc = new_official_doc()
    title(doc, "Cronograma Paralelo de Execução", "Dispensa emergencial + Pregão eletrônico")
    para(
        doc,
        f"Base factual: Processo nº {D['processo_origem']}, licitação deserta comunicada em "
        f"{D['data_deserto']}. Objetivo: garantir o oxigênio agora (Processo 1) e contratar em "
        "definitivo (Processo 2).",
        first_indent=False,
    )
    simple_table(
        doc,
        ["Semana", "Processo 1 — Dispensa (Opção C)", "Processo 2 — Pregão (Opção A)"],
        [
            [
                "1",
                "Cotações; completar DFD/ETP/TR; justificativa; mapa de riscos; envio ao jurídico",
                "Revisão do TR; nova pesquisa de preços; ETP completo; edital",
            ],
            [
                "2",
                "Autorização; PNCP; empenho; assinatura; início do fornecimento",
                "Publicação do edital; ofícios a empresas do ramo; prazo mínimo 8 dias úteis",
            ],
            [
                "3",
                "Fornecimento + fiscalização de estoque e medições",
                "Sessão pública; análise de propostas e habilitação",
            ],
            [
                "4",
                "Acompanhar consumo; preparar transição",
                "Adjudicação / homologação",
            ],
            [
                "Após homologar",
                "Rescisão automática da dispensa com transição assistida (até 10 dias úteis)",
                "Assinatura do contrato definitivo (12 meses, prorrogável)",
            ],
        ],
    )
    heading(doc, "Checklist de urgência")
    for item in [
        "Informar estoque atual em dias na justificativa da emergência",
        "Anexar ata/despacho da licitação deserta e relatório de estoque",
        "Coletar no mínimo 3 cotações (ou justificar número inferior)",
        "Indicar gestor e fiscal do contrato nos dois processos",
        "Publicar a dispensa no PNCP e no Portal da Transparência",
        "Enviar ofícios/e-mails a distribuidores de gases medicinais da região",
        "Manter ambos os autos tramitando simultaneamente, sem arquivamento do 1242/2026 antes da solução",
    ]:
        checkbox_line(doc, item)
    signatures(
        doc,
        [
            [D["diretora"], "Diretoria de Atenção Primária"],
            ["Chefia de Compras / Licitações", "SMS"],
        ],
    )
    save(doc, P0 / "00.02-cronograma-paralelo.docx")


# =============================================================================
# 01 — DISPENSA EMERGENCIAL
# =============================================================================

def p1_capa():
    doc = new_official_doc()
    title(doc, "Capa do Processo Administrativo", "Processo 1 — Dispensa Emergencial")
    boxed(
        doc,
        "IDENTIFICAÇÃO",
        [
            "Processo nº: _______________ / 2026",
            f"Processo de origem / referência: {D['processo_origem']}",
            "Tipo: Contratação direta — Dispensa de licitação",
            "Classificação: Material hospitalar / serviço continuado sem dedicação exclusiva de mão de obra",
            "Unidade demandante: Secretaria Municipal de Saúde — Diretoria de Atenção Primária",
            f"Solicitante: {D['solicitante']}",
            f"Contato: {D['contato']}",
            f"Endereço: {D['endereco']}",
        ],
    )
    boxed(
        doc,
        "1. OBJETO",
        [
            "Contratação emergencial de empresa especializada no fornecimento contínuo, sob demanda, "
            "de Oxigênio Medicinal gasoso, com cessão de cilindros em regime de comodato, manutenção "
            "e logística de entrega, para atender as Unidades Básicas de Saúde (UBSs) e Unidades "
            "Básicas de Atendimento Imediato (UBAIs) das zonas urbana e rural do Município de Pelotas/RS, "
            f"pelo prazo estritamente necessário de {D['prazo_emerg']}, ou até a conclusão da licitação "
            "regular instaurada em paralelo, o que ocorrer primeiro.",
        ],
    )
    boxed(
        doc,
        "2. FUNDAMENTAÇÃO LEGAL",
        [
            "Art. 75, inciso VIII, da Lei nº 14.133/2021; arts. 72 e 74 da mesma lei, no que couber; "
            "normas municipais de contratação; ANVISA/ABNT aplicáveis a gases medicinais.",
        ],
    )
    boxed(
        doc,
        "3. FATO GERADOR",
        [
            f"Licitação do Processo nº {D['processo_origem']} restou deserta (comunicação de {D['data_deserto']}). "
            "Risco de desabastecimento de oxigênio medicinal, com potencial prejuízo à vida e à continuidade "
            "do serviço público de saúde.",
        ],
    )
    boxed(
        doc,
        "4. QUANTITATIVO E VALOR",
        [
            "Consumo médio mensal: 206,0 m³ (urbano 182,0 m³ + rural 24,0 m³).",
            "Estimativa para 180 dias: 1.236,0 m³ (referencial, sem consumo mínimo).",
            "Valor estimado: R$ _______________ (conforme mapa de preços).",
            f"Dotação: {D['projeto']} | Fonte: {D['fonte']}",
            f"Conta: {D['conta']}",
        ],
    )
    signatures(
        doc,
        [
            [D["diretora"], "Diretora de Atenção Primária"],
            [D["secretaria"], "Secretária Municipal de Saúde"],
        ],
    )
    save(doc, P1 / "01.01-capa-autuacao.docx")


def p1_dfd():
    doc = new_official_doc()
    title(doc, "Documento de Formalização da Demanda — DFD", "Versão emergencial — Art. 75, VIII, Lei nº 14.133/2021")
    boxed(
        doc,
        "UNIDADE DEMANDANTE",
        ["SECRETARIA MUNICIPAL DE SAÚDE – SMS. Diretoria de Atenção Primária."],
    )
    boxed(
        doc,
        "1. OBJETO",
        [
            "Contratação emergencial de empresa especializada para fornecimento contínuo, sob demanda, "
            "de Gases Medicinais (Oxigênio Medicinal), com cessão de cilindros em regime de comodato, "
            "para atender às demandas das UBSs e UBAIs, zonas urbana e rural de Pelotas/RS, pelo prazo "
            f"estritamente necessário ({D['prazo_emerg_curto']}) enquanto se conclui a licitação regular.",
            "",
            "Tipo do objeto:",
            "(   ) Material de consumo",
            "(   ) Material Permanente / Equipamento médico-assistencial",
            "(   ) Serviço continuado com dedicação exclusiva de mão de obra",
            "( x ) Serviço continuado sem dedicação exclusiva de mão de obra",
            "(   ) Serviço não continuado",
        ],
    )
    boxed(
        doc,
        "2. FORMA DE CONTRATAÇÃO SUGERIDA",
        [
            "Dispensa de licitação, com fulcro no art. 75, inciso VIII, da Lei nº 14.133/2021, "
            f"em razão da licitação deserta do Processo nº {D['processo_origem']} e do risco de "
            "desabastecimento de oxigênio medicinal, essencial à assistência à saúde.",
        ],
    )
    boxed(
        doc,
        "3. JUSTIFICATIVA DA NECESSIDADE",
        [
            "A presente demanda visa assegurar o abastecimento contínuo e ininterrupto de Oxigênio "
            "Medicinal nas UBSs e UBAIs, insumo essencial ao suporte à vida, à estabilização clínica "
            "e à realização de procedimentos assistenciais (inalações e nebulizações) na Atenção Primária.",
            "",
            "A contratação contempla fornecimento por metro cúbico (m³) efetivamente consumido, com "
            "cilindros em comodato, sem ônus de locação, cabendo à contratada a logística de entrega/"
            "retirada, substituição de recipientes, manutenção preventiva e corretiva, testes "
            "hidrostáticos, limpeza, substituição de válvulas e atendimento às normas ABNT/ANVISA.",
            "",
            f"O certame do Processo nº {D['processo_origem']} restou deserto. A tramitação média de "
            "novo pregão (30 a 45 dias) é incompatível com a urgência do abastecimento. A situação "
            "não decorre de falta de planejamento, mas de ausência de interessados no certame regular. "
            "Paralelamente, instaurar-se-á novo Pregão Eletrônico para solução definitiva.",
        ],
    )
    boxed(
        doc,
        "4. QUANTIDADE E DESCRIÇÃO",
        [
            "Estimativas referenciais, sem obrigação de consumo mínimo. Faturamento pelo volume efetivamente fornecido.",
        ],
    )
    simple_table(
        doc,
        ["Item", "Especificação", "Unidade", "Qtde/mês", "Qtde/180 dias"],
        [
            ["1", "Oxigênio medicinal gasoso — Zona Urbana (UBSs + UBAIs)", "m³", "182,0", "1.092,0"],
            ["2", "Oxigênio medicinal gasoso — Zona Rural (12 UBSs)", "m³", "24,0", "144,0"],
            ["", "TOTAL ESTIMADO", "m³", "206,0", "1.236,0"],
        ],
    )
    boxed(
        doc,
        "5. ESTIMATIVA DE VALOR / 6. DOTAÇÃO",
        [
            "Valor estimado: R$ _______________ (mapa de preços anexo).",
            f"Projeto/Atividade: {D['projeto']}",
            f"Fonte: {D['fonte']}",
            f"Conta: {D['conta']}",
            "Natureza da despesa: _______________",
        ],
    )
    boxed(
        doc,
        "7. LOCAL DE ENTREGA / 8. INÍCIO",
        [
            "Conforme Anexo de Locais de Entrega (41 pontos urbanos e 12 rurais).",
            "Início imediato após assinatura/empenho, com primeira entrega em até 24 horas; "
            "em situações críticas, em até 12 horas.",
        ],
    )
    para(
        doc,
        "Certificamos que somos responsáveis pela elaboração do presente instrumento e que o mesmo "
        "está em consonância com a legislação vigente.",
        first_indent=False,
    )
    signatures(
        doc,
        [
            [D["diretora"], "Diretora de Atenção Primária"],
            [D["secretaria"], "Secretária Municipal de Saúde — Autorizo o prosseguimento"],
        ],
    )
    save(doc, P1 / "01.02-dfd-emergencial.docx")


def p1_etp():
    doc = new_official_doc()
    title(doc, "Estudo Técnico Preliminar — ETP Simplificado", "Contratação emergencial de oxigênio medicinal")
    secoes = [
        (
            "1. Descrição da necessidade",
            "Assegurar abastecimento ininterrupto de oxigênio medicinal nas UBSs/UBAIs após a "
            f"licitação deserta do Processo nº {D['processo_origem']}, enquanto se conclui Pregão "
            "Eletrônico em paralelo.",
        ),
        (
            "2. Área requisitante",
            "Secretaria Municipal de Saúde / Diretoria de Atenção Primária.",
        ),
        (
            "3. Descrição dos requisitos da contratação",
            "Oxigênio medicinal gasoso, pureza mínima 99,5% (v/v); cilindros em comodato (7 m³ / 10 m³ "
            "ou equivalente); entrega em até 24h (12h em urgência); laudo de pureza; troca imediata de "
            "cilindro defeituoso; assistência 24h; cobertura urbana e rural; conformidade ANVISA/ABNT.",
        ),
        (
            "4. Estimativa das quantidades",
            "206,0 m³/mês; 1.236,0 m³ no período de 180 dias (referencial).",
        ),
        (
            "5. Levantamento de mercado (simplificado)",
            "Cotações diretas a empresas de gases medicinais da região, preferencialmente no mínimo "
            "três fontes, com registro no mapa de preços. Painéis oficiais e contratações similares "
            "podem complementar a análise.",
        ),
        (
            "6. Descrição da solução como um todo",
            "Contratação direta emergencial por m³ consumido, com comodato e logística inclusos, "
            "rescindível automaticamente na homologação do pregão regular.",
        ),
        (
            "7. Estimativa do valor da contratação",
            "R$ _______________ — a consolidar no mapa de preços.",
        ),
        (
            "8. Justificativa para o parcelamento ou não",
            "Na emergência, admite-se contratação única abrangendo zona urbana e rural para garantir "
            "cobertura imediata. No pregão paralelo, mantém-se a divisão em dois lotes para competitividade.",
        ),
        (
            "9. Contratações correlatas",
            "Pregão Eletrônico instaurado em paralelo (Processo 2), com transição operacional.",
        ),
        (
            "10. Alinhamento ao planejamento",
            f"Dotação {D['projeto']}, fonte {D['fonte']}. Necessidade imprevisível quanto ao resultado "
            "do certame (desertificação), não quanto à demanda assistencial, já planejada.",
        ),
        (
            "11. Resultados pretendidos",
            "Continuidade assistencial, ausência de desabastecimento e transição ordenada para o contrato definitivo.",
        ),
        (
            "12. Providências prévias à contratação",
            "Relatório de estoque; cotações; indicação de fiscal/gestor; parecer jurídico; publicação PNCP.",
        ),
        (
            "13. Impactos ambientais",
            "Baixo impacto; responsabilidade da contratada sobre cilindros, transporte e destinação conforme normas.",
        ),
        (
            "14. Viabilidade da contratação",
            "Viável técnica, administrativa e orçamentariamente; única solução imediata apta a evitar dano à saúde.",
        ),
        (
            "15. Posicionamento conclusivo",
            "Recomenda-se a contratação emergencial pelo prazo estritamente necessário, com processo de "
            "pregão em paralelo.",
        ),
    ]
    for t, b in secoes:
        boxed(doc, t, [b])
    signatures(doc, [[D["diretora"], "Diretora de Atenção Primária — Equipe de Planejamento"]])
    save(doc, P1 / "01.03-etp-simplificado.docx")


def p1_riscos():
    doc = new_official_doc()
    title(doc, "Mapa de Gerenciamento de Riscos", "Processo 1 — Dispensa emergencial")
    simple_table(
        doc,
        ["Risco", "Prob.", "Impacto", "Tratamento"],
        [
            [
                "Desabastecimento antes da assinatura",
                "Alta",
                "Alto",
                "Priorizar cotações e autorização; estoque de segurança; entrega 24h/12h",
            ],
            [
                "Preço acima do mercado",
                "Média",
                "Médio",
                "Mapa com ≥3 fontes; justificativa formal do preço",
            ],
            [
                "Fornecedor sem capacidade logística rural",
                "Média",
                "Alto",
                "Exigir cobertura territorial e assistência 24h na proposta",
            ],
            [
                "Atraso do pregão paralelo",
                "Média",
                "Médio",
                "Cronograma integrado; monitorar SL; vedar prorrogação indevida da emergência",
            ],
            [
                "Irregularidade sanitária de cilindros",
                "Baixa",
                "Alto",
                "Exigir ANVISA/licença; fiscalização; troca imediata",
            ],
            [
                "Transição abrupta ao contrato definitivo",
                "Média",
                "Médio",
                "Cláusula de transição de até 10 dias úteis na rescisão automática",
            ],
        ],
    )
    para(
        doc,
        "A equipe de planejamento acompanhará os riscos durante toda a vigência emergencial, "
        "reportando à Diretoria de Atenção Primária e ao fiscal do contrato.",
    )
    signatures(doc, [[D["diretora"], "Diretora de Atenção Primária"]])
    save(doc, P1 / "01.04-mapa-gerenciamento-de-riscos.docx")


def p1_tr():
    doc = new_official_doc()
    title(doc, "Termo de Referência Emergencial", "Fornecimento de Oxigênio Medicinal — Dispensa Art. 75, VIII")
    boxed(
        doc,
        "1. OBJETO",
        [
            "Contratação emergencial de empresa especializada no fornecimento contínuo, sob demanda, "
            "de Oxigênio Medicinal gasoso, com comodato de cilindros, manutenção e logística de entrega "
            f"nas UBSs e UBAIs de Pelotas/RS, pelo prazo de {D['prazo_emerg']}, ou até a conclusão da "
            "licitação regular, o que ocorrer primeiro.",
        ],
    )
    boxed(
        doc,
        "2. FUNDAMENTAÇÃO / 3. DESCRIÇÃO DA SOLUÇÃO",
        [
            f"Art. 75, VIII, Lei nº 14.133/2021; Processo nº {D['processo_origem']}; DFD e ETP deste expediente.",
            "Solução: fornecimento por m³ efetivamente consumido + comodato + manutenção + logística + assistência 24h.",
        ],
    )
    heading(doc, "4. Especificações técnicas")
    for item in [
        "Oxigênio medicinal gasoso, pureza mínima de 99,5% (v/v);",
        "Conformidade com normas ANVISA e ABNT aplicáveis a gases medicinais;",
        "Cilindros de 7 m³, 10 m³ ou volume usual do mercado, em comodato, identificados e com teste hidrostático válido;",
        "Válvulas, mangueiras e dispositivos de segurança em perfeito estado;",
        "Laudo de análise/qualidade quando solicitado pela fiscalização;",
        "Proibição de cobrança de locação de cilindros.",
    ]:
        bullet(doc, item)
    heading(doc, "5. Quantitativos estimados")
    simple_table(
        doc,
        ["Área", "Consumo médio mensal", "Estimativa 180 dias"],
        [
            ["Zona Urbana (UBSs + UBAIs)", "182,0 m³", "1.092,0 m³"],
            ["Zona Rural (12 UBSs)", "24,0 m³", "144,0 m³"],
            ["Total", "206,0 m³", "1.236,0 m³"],
        ],
    )
    para(doc, "Quantitativos referenciais. Não há obrigação de consumo mínimo. Pagamento apenas do volume entregue.")
    heading(doc, "6. Modelo de execução")
    para(
        doc,
        "Fornecimento sob demanda, mediante solicitação das unidades ou da Diretoria. "
        "Prazo ordinário de entrega: até 24 (vinte e quatro) horas. Situações críticas "
        "comunicadas pela unidade: até 12 (doze) horas. Substituição imediata de cilindro "
        "com vazamento, válvula defeituosa ou irregularidade sanitária.",
    )
    heading(doc, "7. Locais de entrega")
    para(doc, "Conforme Anexo de Locais de Entrega (41 unidades urbanas e 12 rurais).")
    heading(doc, "8. Obrigações da contratada")
    for item in [
        "Manter estoque operacional compatível com a demanda municipal;",
        "Realizar comodato, manutenção preventiva/corretiva e testes hidrostáticos;",
        "Disponibilizar canal de atendimento 24 horas;",
        "Apresentar documentação sanitária e registro ANVISA do produto;",
        "Emitir notas fiscais conforme medições mensais;",
        "Cumprir legislação trabalhista, tributária, ambiental e sanitária;",
        "Garantir transição ordenada ao término do contrato.",
    ]:
        bullet(doc, item)
    heading(doc, "9. Obrigações da contratante")
    para(
        doc,
        "Efetuar o pagamento das medições; indicar fiscal e gestor; disponibilizar acesso às unidades; "
        "comunicar demanda com antecedência compatível, ressalvadas urgências clínicas.",
    )
    heading(doc, "10. Gestão e fiscalização")
    meta_line(doc, "Gestor do contrato", "_______________________________ / Matrícula __________")
    meta_line(doc, "Fiscal do contrato", "_______________________________ / Matrícula __________")
    heading(doc, "11. Critérios de medição e pagamento")
    para(
        doc,
        "Medição mensal do volume efetivamente fornecido (m³), com relatório de entregas, NF e atesto "
        "da fiscalização. Pagamento no prazo da legislação municipal. Vedado pagamento por volume não entregue.",
    )
    heading(doc, "12. Vigência e rescisão")
    para(
        doc,
        f"Vigência de {D['prazo_emerg_curto']}, sem prorrogação ordinária. Rescisão automática quando "
        "homologado/assinado o contrato do Pregão regular em paralelo, assegurada transição de até "
        "10 (dez) dias úteis sem desabastecimento.",
    )
    heading(doc, "13. Sanções")
    para(
        doc,
        "Aplicam-se as sanções dos arts. 155 a 163 da Lei nº 14.133/2021 e normas municipais, "
        "especialmente em atraso injustificado, desabastecimento imputável à contratada ou irregularidade sanitária.",
    )
    heading(doc, "14. Estimativa de valor")
    para(doc, "R$ _______________ (mapa de preços). Preço unitário de referência: R$ ________ / m³.")
    signatures(doc, [[D["diretora"], "Diretora de Atenção Primária"]])
    save(doc, P1 / "01.05-termo-de-referencia-emergencial.docx")


def p1_justificativa():
    doc = new_official_doc()
    title(
        doc,
        "Justificativa da Emergência e do Enquadramento Legal",
        f"Art. 75, VIII, Lei nº 14.133/2021 — Processo nº {D['processo_origem']}",
    )
    heading(doc, "I — Dos fatos")
    para(
        doc,
        f"A licitação referente ao Processo nº {D['processo_origem']} (SL nº {D['sl']}) restou deserta, "
        f"conforme comunicação oficial de {D['data_deserto']} e tramitações precedentes "
        f"(Memorando nº {D['memo']}; Memorando CCGOF nº {D['memo_ccgof']}), documentos que instruem este expediente.",
    )
    para(
        doc,
        "O estoque atual de oxigênio medicinal nas unidades da Atenção Primária é de aproximadamente "
        "_____ (_____) dias de autonomia, insuficiente para aguardar a conclusão de novo certame, "
        "cuja tramitação média observada é de 30 a 45 dias.",
    )
    heading(doc, "II — Do direito")
    para(
        doc,
        "O art. 75, VIII, da Lei nº 14.133/2021 admite a dispensa de licitação para contratação "
        "que possa acarretar prejuízo ou comprometer a segurança de pessoas, no caso de emergência "
        "ou calamidade pública, quando caracterizada urgência de atendimento de situação que possa "
        "ocasionar prejuízo ou comprometer a segurança de pessoas, obras, serviços, equipamentos e "
        "outros bens, públicos ou particulares.",
    )
    para(
        doc,
        "No caso concreto, o oxigênio medicinal é insumo crítico, sem substituto equivalente imediato "
        "no âmbito das UBSs e UBAIs. A ausência de interessados no certame regular não decorre de falta "
        "de planejamento da Administração, que instruiu o Processo nº 1242/2026 com DFD, ETP e demais "
        "peças, mas de fator externo (desertificação).",
    )
    heading(doc, "III — Do prazo estritamente necessário")
    para(
        doc,
        f"Propõe-se o prazo de {D['prazo_emerg']}, contados da assinatura, ou até a homologação/"
        "assinatura do contrato do Pregão Eletrônico instaurado em paralelo, o que ocorrer primeiro, "
        "vedada prorrogação ordinária, em conformidade com o limite legal de até 1 (um) ano.",
    )
    heading(doc, "IV — Da estratégia paralela")
    para(
        doc,
        "Simultaneamente, adota-se a Opção A (repetição do certame), com revisão de TR, atualização "
        "de pesquisa de preços e divulgação ampliada, para restabelecer a contratação definitiva e "
        "encerrar a emergência tão logo possível.",
    )
    heading(doc, "V — Do preço")
    para(
        doc,
        "O valor contratado será justificado por pesquisa de mercado (mapa anexo), compatível com "
        "os preços regionais, considerando comodato e logística urbana/rural.",
    )
    heading(doc, "VI — Conclusão")
    para(
        doc,
        "Diante do risco concreto à continuidade da assistência à saúde, requer-se a autorização "
        "da dispensa emergencial e a tramitação prioritária do pregão paralelo.",
    )
    heading(doc, "Anexos obrigatórios")
    for item in [
        "Ata/termo/despacho da licitação deserta;",
        "Relatório/print do estoque atual (em dias);",
        "DFD, ETP, TR, mapa de riscos;",
        "Mapa de preços e cotações;",
        "Minuta contratual.",
    ]:
        bullet(doc, item)
    signatures(
        doc,
        [
            [D["diretora"], "Diretora de Atenção Primária"],
            [D["secretaria"], "Secretária Municipal de Saúde"],
        ],
    )
    save(doc, P1 / "01.06-justificativa-emergencia-enquadramento.docx")


def p1_precos():
    doc = new_official_doc()
    title(doc, "Mapa de Pesquisa de Preços", "Contratação emergencial — Oxigênio medicinal (R$/m³)")
    para(
        doc,
        "Pesquisa elaborada nos termos do art. 23 da Lei nº 14.133/2021 e regulamentação aplicável. "
        "Fontes preferenciais: cotações de empresas de gases medicinais da região, Painel de Preços "
        "e contratações similares. Mínimo de 3 (três) fontes; número inferior exige justificativa.",
        first_indent=False,
    )
    simple_table(
        doc,
        ["Fonte / Empresa", "CNPJ", "R$/m³ Urbano", "R$/m³ Rural", "Validade", "Obs."],
        [
            ["Cotação 1 — _____________", "", "", "", "", ""],
            ["Cotação 2 — _____________", "", "", "", "", ""],
            ["Cotação 3 — _____________", "", "", "", "", ""],
            ["Painel / contrato similar", "", "", "", "", ""],
        ],
    )
    boxed(
        doc,
        "CONSOLIDAÇÃO",
        [
            "Metodologia: ( ) média  ( ) mediana  ( ) menor preço compatível  ( ) outra: ______",
            "Preço de referência — Zona Urbana: R$ ________ / m³",
            "Preço de referência — Zona Rural: R$ ________ / m³",
            "Valor estimado 180 dias (1.236 m³): R$ _______________",
            "Justificativa do preço escolhido: _______________________________________________",
        ],
    )
    signatures(doc, [["Responsável pela pesquisa", "Setor de Licitações / Compras — SMS"]])
    save(doc, P1 / "01.07-mapa-pesquisa-de-precos.docx")


def p1_razao():
    doc = new_official_doc()
    title(doc, "Razão da Escolha do Fornecedor e Justificativa do Preço")
    boxed(
        doc,
        "FORNECEDOR SELECIONADO",
        [
            "Razão social: _______________________________________________",
            "CNPJ: _____________________________ IE/IM: ________________",
            "Endereço: _________________________________________________",
            "Telefone / e-mail: _________________________________________",
            "Representante legal: _______________________________________",
        ],
    )
    heading(doc, "1. Razão da escolha")
    checkbox_line(doc, "Apresentou o menor preço compatível com a pesquisa de mercado")
    checkbox_line(doc, "Única que assegura entrega em até 24h com estoque imediato")
    checkbox_line(doc, "Única com habilitação sanitária e capacidade logística urbana/rural no prazo emergencial")
    checkbox_line(doc, "Outra: ________________________________________________")
    heading(doc, "2. Comparativo")
    para(
        doc,
        "O preço unitário ofertado de R$ ________ / m³ situa-se ( ) abaixo / ( ) na média / "
        "( ) até _____% acima da média/mediana do mapa, permanecendo justificado em face da "
        "urgência, do comodato e da cobertura territorial.",
    )
    heading(doc, "3. Conclusão")
    para(
        doc,
        "Com fundamento no art. 75, VIII, da Lei nº 14.133/2021, e comprovada a compatibilidade "
        "de preços, recomenda-se a contratação direta da empresa indicada pelo prazo emergencial.",
    )
    signatures(doc, [["Responsável técnico-administrativo", "SMS / Compras / Licitações"]])
    save(doc, P1 / "01.08-razao-escolha-fornecedor.docx")


def p1_hab():
    doc = new_official_doc()
    title(doc, "Checklist de Habilitação", "Dispensa emergencial — Oxigênio medicinal")
    meta_line(doc, "Razão social", "________________________________")
    meta_line(doc, "CNPJ", "____________________")
    simple_table(
        doc,
        ["Documento", "Apresentado", "Validade", "Obs."],
        [
            ["Contrato social / ato constitutório", "[ ]", "", ""],
            ["Cartão CNPJ", "[ ]", "", ""],
            ["Documento do representante / procuração", "[ ]", "", ""],
            ["Certidão Federal (Receita/PGFN)", "[ ]", "", ""],
            ["Certidão Estadual", "[ ]", "", ""],
            ["Certidão Municipal", "[ ]", "", ""],
            ["Certificado FGTS", "[ ]", "", ""],
            ["CNDT (trabalhista)", "[ ]", "", ""],
            ["Licença sanitária / alvará de vigilância", "[ ]", "", ""],
            ["Registro/autorização ANVISA do oxigênio medicinal", "[ ]", "", ""],
            ["Atestado de capacidade técnica", "[ ]", "", ""],
            ["Proposta/cotação assinada", "[ ]", "", ""],
            ["Declarações exigidas", "[ ]", "", ""],
        ],
    )
    para(doc, "Resultado da análise: ( ) Habilitada  ( ) Inabilitada — motivo: ________________")
    signatures(doc, [["Servidor responsável pela conferência", "Matrícula / Setor"]])
    save(doc, P1 / "01.09-checklist-habilitacao.docx")


def p1_declaracoes():
    doc = new_official_doc()
    title(doc, "Modelos de Declarações do Fornecedor", "Anexar assinadas aos autos da dispensa")
    heading(doc, "Declaração 1 — Cumprimento dos requisitos de habilitação")
    para(
        doc,
        "A empresa _____________, CNPJ _____________, declara, sob as penas da lei, que cumpre "
        "plenamente os requisitos de habilitação exigidos para a presente contratação direta.",
    )
    heading(doc, "Declaração 2 — Inexistência de fatos impeditivos")
    para(
        doc,
        "Declara que não possui em seu quadro societário servidor público da ativa, ou empregado "
        "de empresa pública/sociedade de economia mista, e que não há fatos impeditivos à contratação "
        "com o Poder Público.",
    )
    heading(doc, "Declaração 3 — Cumprimento do art. 7º, XXXIII, da CF/88")
    para(
        doc,
        "Declara que não emprega menor de dezoito anos em trabalho noturno, perigoso ou insalubre, "
        "e não emprega menor de dezesseis anos, salvo na condição de aprendiz, a partir de quatorze anos.",
    )
    heading(doc, "Declaração 4 — Capacidade de fornecimento emergencial")
    para(
        doc,
        "Declara possuir condições de iniciar o fornecimento de oxigênio medicinal com comodato de "
        "cilindros em até 24 horas da ordem de serviço/empenho, inclusive nas unidades da zona rural "
        "relacionadas no anexo de locais de entrega, e de manter assistência 24 horas.",
    )
    signatures(doc, [["Representante legal da empresa", "Nome / CPF / cargo"]])
    save(doc, P1 / "01.10-declaracoes-fornecedor.docx")


def p1_contrato():
    doc = new_official_doc()
    title(
        doc,
        "Minuta de Contrato Emergencial nº _____/2026",
        "Fornecimento de oxigênio medicinal — Art. 75, VIII, Lei nº 14.133/2021",
    )
    para(
        doc,
        "Pelo presente instrumento, de um lado o MUNICÍPIO DE PELOTAS, pessoa jurídica de direito "
        "público interno, CNPJ 87.455.531/0001-57, por intermédio da Secretaria Municipal de Saúde, "
        f"com sede na {D['endereco']}, neste ato representada pela Secretária Municipal de Saúde, "
        f"{D['secretaria']}, doravante CONTRATANTE; e, de outro lado, a empresa _____________, "
        "CNPJ _____________, com sede em _____________, neste ato representada por _____________, "
        "doravante CONTRATADA, celebram o presente contrato, mediante as cláusulas seguintes:",
        first_indent=False,
    )
    clausulas = [
        (
            "CLÁUSULA PRIMEIRA — DO OBJETO",
            "Fornecimento contínuo, sob demanda, de oxigênio medicinal gasoso (pureza mínima 99,5%), "
            "com cessão de cilindros em comodato, manutenção e logística de entrega nas UBSs e UBAIs "
            "do Município de Pelotas/RS, conforme Termo de Referência Emergencial anexo, que integra "
            "este contrato independentemente de transcrição.",
        ),
        (
            "CLÁUSULA SEGUNDA — DA FUNDAMENTAÇÃO",
            f"Contratação por dispensa de licitação, art. 75, VIII, Lei nº 14.133/2021, em razão da "
            f"licitação deserta do Processo nº {D['processo_origem']} e do risco de desabastecimento.",
        ),
        (
            "CLÁUSULA TERCEIRA — DO PREÇO E DO PAGAMENTO",
            "O preço unitário é de R$ ________ por m³ efetivamente fornecido. O valor global estimado "
            "é de R$ _______________. Pagamento mensal por medição, após atesto, vedado pagamento por "
            "volume não entregue. Nos preços estão inclusos comodato, manutenção, logística, tributos e encargos.",
        ),
        (
            "CLÁUSULA QUARTA — DA VIGÊNCIA",
            f"O prazo de vigência é de {D['prazo_emerg']}, contados da assinatura, ou até a conclusão "
            "da licitação regular instaurada em paralelo, o que ocorrer primeiro. Não haverá prorrogação ordinária.",
        ),
        (
            "CLÁUSULA QUINTA — DA RESCISÃO AUTOMÁTICA E TRANSIÇÃO",
            "Este contrato extingue-se automaticamente com a homologação do Pregão Eletrônico e "
            "assinatura do contrato definitivo correspondente, assegurada transição operacional sem "
            "interrupção do abastecimento, no prazo de até 10 (dez) dias úteis, salvo ajuste diverso "
            "com a fiscalização.",
        ),
        (
            "CLÁUSULA SEXTA — DAS OBRIGAÇÕES DA CONTRATADA",
            "Cumprir o Termo de Referência; manter habilitação; garantir entrega em até 24 horas "
            "(12 horas em urgência); assistência 24h; conformidade sanitária; comodato e manutenção "
            "de cilindros; emitir documentação fiscal adequada.",
        ),
        (
            "CLÁUSULA SÉTIMA — DAS OBRIGAÇÕES DA CONTRATANTE",
            "Efetuar os pagamentos devidos; indicar gestor e fiscal; assegurar acesso às unidades; "
            "comunicar demandas e fiscalizar a execução.",
        ),
        (
            "CLÁUSULA OITAVA — DA FISCALIZAÇÃO",
            "Gestor: ________________. Fiscal: ________________. A fiscalização não exclui nem reduz "
            "a responsabilidade da CONTRATADA.",
        ),
        (
            "CLÁUSULA NONA — DAS SANÇÕES",
            "Aplicam-se as sanções dos arts. 155 a 163 da Lei nº 14.133/2021 e normas municipais, "
            "assegurados contraditório e ampla defesa.",
        ),
        (
            "CLÁUSULA DÉCIMA — DA DOTAÇÃO",
            f"As despesas correrão à conta da dotação {D['projeto']}, fonte {D['fonte']}, "
            "nota de empenho nº _______________.",
        ),
        (
            "CLÁUSULA DÉCIMA PRIMEIRA — DA PROTEÇÃO DE DADOS",
            "As partes observarão a Lei nº 13.709/2018 (LGPD) no tratamento de dados pessoais eventualmente envolvidos.",
        ),
        (
            "CLÁUSULA DÉCIMA SEGUNDA — DO FORO",
            "Fica eleito o foro da Comarca de Pelotas/RS, com renúncia de qualquer outro.",
        ),
    ]
    for t, b in clausulas:
        heading(doc, t)
        para(doc, b)
    para(doc, "E, por estarem justas e contratadas, as partes assinam o presente instrumento.", first_indent=False)
    signatures(
        doc,
        [
            ["MUNICÍPIO DE PELOTAS / SMS", "CONTRATANTE"],
            ["Empresa _______________________", "CONTRATADA"],
            ["Testemunha 1 — Nome/CPF", "Testemunha 2 — Nome/CPF"],
        ],
    )
    save(doc, P1 / "01.11-minuta-contrato-emergencial.docx")


def p1_parecer_solicitacao():
    doc = new_official_doc()
    title(doc, "Solicitação de Parecer Jurídico", "Dispensa emergencial — Oxigênio medicinal")
    para(
        doc,
        "À Assessoria Jurídica / Procuradoria-Geral do Município,",
        first_indent=False,
    )
    para(
        doc,
        "Solicita-se parecer jurídico quanto à legalidade da contratação direta emergencial de "
        f"oxigênio medicinal, com fulcro no art. 75, VIII, da Lei nº 14.133/2021, em face da "
        f"licitação deserta do Processo nº {D['processo_origem']}, instruída com as peças deste Volume 1.",
    )
    heading(doc, "Pontos a examinar")
    for item in [
        "Competência da autoridade para autorizar a dispensa;",
        "Enquadramento no art. 75, VIII, e observância do art. 72 da Lei nº 14.133/2021;",
        "Comprovação da licitação deserta e do risco de prejuízo/desabastecimento;",
        "Prazo estritamente necessário e cláusula de rescisão automática;",
        "Pesquisa de preços e justificativa do valor;",
        "Habilitação jurídica, fiscal e sanitária da contratada;",
        "Cláusulas essenciais da minuta;",
        "Publicidade no PNCP;",
        "Compatibilidade com o pregão paralelo (mitigação / boa prática).",
    ]:
        bullet(doc, item)
    signatures(doc, [[D["diretora"], "Diretora de Atenção Primária"]])
    save(doc, P1 / "01.12-solicitacao-parecer-juridico.docx")


def p1_autorizacao():
    doc = new_official_doc()
    title(doc, "Despacho de Autorização da Dispensa Emergencial")
    para(
        doc,
        f"Trata-se de pedido de contratação emergencial de oxigênio medicinal, em face da licitação "
        f"deserta do Processo nº {D['processo_origem']}, com risco de desabastecimento das UBSs e UBAIs.",
    )
    para(
        doc,
        "À vista do DFD, ETP, Mapa de Riscos, Termo de Referência, justificativa da emergência, "
        "pesquisa de preços, razão de escolha, habilitação, minuta contratual e parecer jurídico "
        "(quando emitido), e considerando o art. 75, VIII, da Lei nº 14.133/2021,",
    )
    p = doc.add_paragraph()
    p.alignment = 1
    from docx.shared import Pt
    from oficial_template import _set_run_font

    r = p.add_run("AUTORIZO")
    _set_run_font(r, size=14, bold=True)
    para(
        doc,
        f"a dispensa de licitação e a contratação emergencial pelo prazo de {D['prazo_emerg']}, "
        "ou até a conclusão do Pregão Eletrônico paralelo, o que ocorrer primeiro, no valor estimado "
        "de R$ _______________, com a empresa _____________ — CNPJ _____________, determinando:",
    )
    for item in [
        "Publicação no PNCP e no Portal da Transparência;",
        "Emissão de empenho e assinatura do contrato;",
        "Início imediato do fornecimento, sob fiscalização da Diretoria de Atenção Primária;",
        "Prosseguimento prioritário do Pregão Eletrônico em paralelo.",
    ]:
        bullet(doc, item)
    signatures(doc, [[D["secretaria"], "Secretária Municipal de Saúde"]])
    save(doc, P1 / "01.13-despacho-autorizacao.docx")


def p1_aviso_pncp():
    doc = new_official_doc()
    title(doc, "Aviso / Extrato de Dispensa de Licitação", "Para publicação no PNCP e Portal da Transparência")
    boxed(
        doc,
        "EXTRATO",
        [
            "ESPÉCIE: Dispensa de Licitação — Art. 75, VIII, Lei nº 14.133/2021.",
            "ÓRGÃO: Secretaria Municipal de Saúde — Município de Pelotas/RS.",
            f"PROCESSO: _____/2026 (origem {D['processo_origem']}).",
            "OBJETO: Fornecimento emergencial de oxigênio medicinal gasoso, com comodato de cilindros, "
            "manutenção e logística, para UBSs e UBAIs (zonas urbana e rural).",
            "CONTRATADA: _____________ — CNPJ _____________.",
            f"VIGÊNCIA: {D['prazo_emerg_curto']} ou até homologação do pregão regular, o que ocorrer primeiro.",
            "VALOR GLOBAL ESTIMADO: R$ _____________.",
            "PREÇO UNITÁRIO: R$ ________ / m³.",
            "FUNDAMENTO: Art. 75, VIII, Lei nº 14.133/2021.",
            "DATA DA AUTORIZAÇÃO: ___/___/2026.",
            "AUTORIDADE: Secretária Municipal de Saúde.",
        ],
    )
    para(
        doc,
        "O inteiro teor do processo poderá ser consultado na Secretaria Municipal de Saúde / Setor de "
        "Licitações, sem prejuízo da publicidade legal no PNCP.",
        first_indent=False,
    )
    signatures(doc, [["Setor de Licitações / Compras", "SMS Pelotas"]])
    save(doc, P1 / "01.14-aviso-extrato-dispensa-pncp.docx")


def p1_extrato_contrato():
    doc = new_official_doc()
    title(doc, "Extrato de Contrato Emergencial", "Publicação — Art. 94 da Lei nº 14.133/2021")
    boxed(
        doc,
        "EXTRATO DE CONTRATO Nº _____/2026",
        [
            "CONTRATANTE: Município de Pelotas / Secretaria Municipal de Saúde.",
            "CONTRATADA: _____________ — CNPJ _____________.",
            "OBJETO: Fornecimento emergencial de oxigênio medicinal com comodato de cilindros.",
            f"VIGÊNCIA: {D['prazo_emerg_curto']} (ou até homologação do pregão paralelo).",
            "VALOR: R$ _____________ (unitário R$ ________ / m³).",
            "DOTAÇÃO: " + D["projeto"] + f" | Fonte {D['fonte']}.",
            "DATA DA ASSINATURA: ___/___/2026.",
            "PROCESSO: _____/2026.",
        ],
    )
    signatures(doc, [["Autoridade competente", "SMS"]])
    save(doc, P1 / "01.15-extrato-contrato.docx")


# =============================================================================
# 02 — PREGÃO
# =============================================================================

def p2_capa():
    doc = new_official_doc()
    title(doc, "Capa do Processo Administrativo", "Processo 2 — Pregão Eletrônico (solução definitiva)")
    boxed(
        doc,
        "IDENTIFICAÇÃO",
        [
            "Processo nº: _______________ / 2026",
            f"Referência / origem: Processo nº {D['processo_origem']} (licitação deserta)",
            "Modalidade: Pregão Eletrônico — Lei nº 14.133/2021",
            "Critério: Menor preço por lote",
            "Unidade demandante: SMS — Diretoria de Atenção Primária",
        ],
    )
    boxed(
        doc,
        "1. OBJETO",
        [
            "Contratação de empresa(s) especializada(s) para fornecimento contínuo, sob demanda, de "
            "Oxigênio Medicinal gasoso, com comodato de cilindros, manutenção e logística, para UBSs "
            "e UBAIs das zonas urbana e rural de Pelotas/RS, pelo prazo de 12 (doze) meses, prorrogável "
            "nos termos da lei, em 2 (dois) lotes.",
        ],
    )
    boxed(
        doc,
        "2. ESTRATÉGIA / 3. QUANTITATIVOS / 4. DOTAÇÃO",
        [
            "Repetição do certame (Opção A) em paralelo à dispensa emergencial, com TR revisado, "
            "pesquisa atualizada, habilitação proporcional e divulgação ampliada.",
            "Lote 1 (urbano): 2.184,0 m³/ano | Lote 2 (rural): 288,0 m³/ano | Total: 2.472,0 m³/ano.",
            f"Dotação: {D['projeto']} | Fonte: {D['fonte']} | Valor estimado: R$ _______________",
        ],
    )
    signatures(
        doc,
        [
            [D["diretora"], "Diretora de Atenção Primária"],
            [D["secretaria"], "Secretária Municipal de Saúde"],
        ],
    )
    save(doc, P2 / "02.01-capa-autuacao.docx")


def p2_dfd():
    doc = new_official_doc()
    title(doc, "Documento de Formalização da Demanda — DFD", "Versão revisada para novo Pregão Eletrônico")
    boxed(doc, "UNIDADE DEMANDANTE", ["SECRETARIA MUNICIPAL DE SAÚDE – SMS. Diretoria de Atenção Primária."])
    boxed(
        doc,
        "1. OBJETO",
        [
            "Contratação de empresa especializada para fornecimento contínuo, sob demanda, de Gases "
            "Medicinais (Oxigênio Medicinal), com cessão de cilindros em regime de comodato, para "
            "UBSs e UBAIs, zonas urbana e rural de Pelotas/RS.",
            "",
            "Tipo do objeto: ( x ) Serviço continuado sem dedicação exclusiva de mão de obra",
        ],
    )
    boxed(
        doc,
        "2. FORMA DE CONTRATAÇÃO SUGERIDA",
        [
            "PREGÃO ELETRÔNICO, critério de MENOR PREÇO por lote, art. 6º, XLI, e art. 29, parágrafo "
            "único, da Lei nº 14.133/2021.",
        ],
    )
    boxed(
        doc,
        "3. JUSTIFICATIVA",
        [
            "Assegurar abastecimento contínuo de oxigênio medicinal na Atenção Primária. O certame "
            f"anterior (Processo nº {D['processo_origem']}) restou deserto. Esta DFD revisada corrige "
            "pontos de atratividade: preço de referência atualizado; prazos realistas; habilitação "
            "proporcional; divulgação ampliada. Mantém-se a divisão em dois lotes (urbano e rural) "
            "para ampliar competitividade e refletir diferenças logísticas.",
        ],
    )
    simple_table(
        doc,
        ["Item", "Especificação", "Unidade", "Qtde anual"],
        [
            [
                "1.1 Lote 1",
                "Oxigênio medicinal gasoso — Zona Urbana (38 UBSs + 3 UBAIs). Média: 152 m³/mês UBSs + 30 m³/mês UBAIs.",
                "m³",
                "2.184,0",
            ],
            [
                "2.1 Lote 2",
                "Oxigênio medicinal gasoso — Zona Rural (12 UBSs). Média: 24 m³/mês.",
                "m³",
                "288,0",
            ],
        ],
    )
    boxed(
        doc,
        "5. VALOR / 6. DOTAÇÃO / 7. LOCAIS / 8. INÍCIO",
        [
            "Valor: conforme pesquisa ATUALIZADA (art. 23, Lei nº 14.133/2021).",
            f"Projeto/Atividade: {D['projeto']} | Fonte: {D['fonte']}",
            "Locais: Anexo de Locais de Entrega.",
            "Início: até 10 (dez) dias úteis da Nota de Empenho / assinatura do contrato.",
        ],
    )
    signatures(
        doc,
        [
            [D["diretora"], "Diretora de Atenção Primária"],
            [D["secretaria"], "Secretária Municipal de Saúde — Autorizo"],
        ],
    )
    save(doc, P2 / "02.02-dfd-revisado.docx")


def p2_etp():
    doc = new_official_doc()
    title(doc, "Estudo Técnico Preliminar — ETP Completo", "Pregão Eletrônico — Oxigênio Medicinal")
    secoes = [
        ("1. Necessidade", f"Abastecimento contínuo após desertificação do Processo nº {D['processo_origem']}."),
        ("2. Área requisitante", "SMS / Diretoria de Atenção Primária."),
        ("3. Requisitos", "Oxigênio ≥ 99,5%; comodato; logística urbana/rural; ANVISA/ABNT; assistência 24h; medição por m³."),
        ("4. Quantidades", "Anual: 2.472,0 m³ (Lote 1: 2.184,0; Lote 2: 288,0)."),
        ("5. Mercado", "Nova pesquisa com painéis, municípios similares e cotações diretas; registrar defasagem do preço anterior."),
        ("6. Solução", "Pregão eletrônico em 2 lotes; contrato 12 meses prorrogável; pagamento por consumo; comodato incluso."),
        ("7. Valor", "R$ _______________ — mapa atualizado."),
        ("8. Parcelamento", "Dois lotes ampliam competitividade e refletem logística rural."),
        ("9. Contratações correlatas", "Dispensa emergencial em paralelo, com rescisão automática na homologação deste pregão."),
        ("10. Planejamento", f"Dotação {D['projeto']}, fonte {D['fonte']}."),
        ("11. Resultados", "Segurança assistencial, economicidade e ampliação de interessados."),
        ("12. Providências", "Revisão de TR; ofícios a empresas; PNCP; prazo mínimo de 8 dias úteis."),
        ("13. Meio ambiente", "Impactos mínimos; responsabilidade da contratada sobre cilindros/transporte."),
        ("14. Viabilidade", "Solução viável e essencial."),
        ("15. Conclusão", "Recomenda-se a instauração do Pregão Eletrônico com os ajustes deste ETP/TR."),
    ]
    for t, b in secoes:
        boxed(doc, t, [b])
    signatures(doc, [[D["diretora"], "Diretora de Atenção Primária"]])
    save(doc, P2 / "02.03-etp-completo.docx")


def p2_riscos():
    doc = new_official_doc()
    title(doc, "Mapa de Gerenciamento de Riscos", "Processo 2 — Pregão Eletrônico")
    simple_table(
        doc,
        ["Risco", "Prob.", "Impacto", "Tratamento"],
        [
            ["Nova desertificação", "Média", "Alto", "Preço atualizado; divulgação ampliada; habilitação proporcional; prazos realistas"],
            ["Preço de referência defasado", "Média", "Alto", "Pesquisa nova com múltiplas fontes"],
            ["Exigências excessivas afastando licitantes", "Média", "Alto", "Revisão do TR: remover exigências desarrazoadas"],
            ["Atraso na sessão/homologação", "Média", "Médio", "Cronograma paralelo com dispensa emergencial"],
            ["Falha de transição entre contratos", "Baixa", "Alto", "Cláusula de transição de até 10 dias úteis"],
        ],
    )
    signatures(doc, [[D["diretora"], "Diretora de Atenção Primária"]])
    save(doc, P2 / "02.04-mapa-gerenciamento-de-riscos.docx")


def p2_tr():
    doc = new_official_doc()
    title(doc, "Termo de Referência Revisado", "Anexo I do Pregão Eletrônico — Oxigênio Medicinal")
    boxed(
        doc,
        "1. OBJETO",
        [
            "Contratação de empresa(s) para fornecimento contínuo, sob demanda, de Oxigênio Medicinal "
            "gasoso, com comodato de cilindros, manutenção e logística, em 2 lotes (zona urbana e zona rural).",
        ],
    )
    heading(doc, "2. Pontos obrigatoriamente revisados em relação ao certame deserto")
    for item in [
        "Pesquisa de preços NOVA, com valores atuais do mercado local/regional;",
        "Prazo de início mais realista: até 10 dias úteis; reposição ordinária em 24h;",
        "Habilitação proporcional — remover exigências excessivas;",
        "Comodato, manutenção e troca de cilindros previstos expressamente;",
        "Vigência de 12 meses, prorrogável na forma da Lei nº 14.133/2021;",
        "Divulgação ampliada a empresas do ramo, além do PNCP.",
    ]:
        bullet(doc, item)
    heading(doc, "3. Especificação / 4. Lotes")
    para(doc, "Oxigênio medicinal gasoso, pureza mínima 99,5%; cilindros em comodato; ANVISA/ABNT; assistência 24h.")
    simple_table(
        doc,
        ["Lote", "Abrangência", "Qtde anual (m³)"],
        [
            ["1", "Zona Urbana — 38 UBSs + 3 UBAIs", "2.184,0"],
            ["2", "Zona Rural — 12 UBSs", "288,0"],
        ],
    )
    heading(doc, "5. Critério de julgamento")
    para(doc, "Menor preço por lote, pregão eletrônico.")
    heading(doc, "6. Habilitação (proporcional)")
    for item in [
        "Habilitação jurídica;",
        "Regularidade fiscal e trabalhista;",
        "Licença sanitária e registro ANVISA do oxigênio medicinal;",
        "Atestado de capacidade técnica compatível, sem quantitativos históricos desarrazoados;",
        "Não exigir documentos não previstos em lei ou desproporcionais ao objeto.",
    ]:
        bullet(doc, item)
    heading(doc, "7. Vigência / 8. Pagamento / 9. Fiscalização / 10. Sanções")
    para(doc, "Vigência: 12 meses, prorrogável. Pagamento mensal por medição. Gestor/Fiscal a indicar. Sanções da Lei nº 14.133/2021.")
    signatures(doc, [[D["diretora"], "Diretora de Atenção Primária"]])
    save(doc, P2 / "02.05-termo-de-referencia-revisado.docx")


def p2_precos():
    doc = new_official_doc()
    title(doc, "Pesquisa de Preços Atualizada", "Novo Pregão — Oxigênio Medicinal")
    para(
        doc,
        "Atualização motivada pela licitação deserta. Fontes: Painel de Preços, contratações "
        "similares e cotações diretas. Justificar eventual defasagem do preço anterior.",
        first_indent=False,
    )
    simple_table(
        doc,
        ["Fonte", "Data", "R$/m³ Lote 1", "R$/m³ Lote 2", "Documento", "Obs."],
        [
            ["Painel de Preços", "", "", "", "", ""],
            ["Município similar 1", "", "", "", "", ""],
            ["Município similar 2", "", "", "", "", ""],
            ["Cotação empresa A", "", "", "", "", ""],
            ["Cotação empresa B", "", "", "", "", ""],
            ["Cotação empresa C", "", "", "", "", ""],
        ],
    )
    boxed(
        doc,
        "CONSOLIDAÇÃO",
        [
            "Preço de referência Lote 1: R$ ________ / m³",
            "Preço de referência Lote 2: R$ ________ / m³",
            "Valor anual estimado total: R$ _______________",
            "Justificativa de diferença em relação ao edital anterior: ________________",
        ],
    )
    signatures(doc, [["Responsável pela pesquisa", "Setor de Licitações — SMS"]])
    save(doc, P2 / "02.06-pesquisa-de-precos-atualizada.docx")


def p2_edital():
    doc = new_official_doc()
    title(doc, "Minuta de Edital — Pregão Eletrônico nº _____/2026", "Fornecimento de oxigênio medicinal")
    meta_line(doc, "Processo", f"_____/2026 (ref. {D['processo_origem']})")
    heading(doc, "1. Preâmbulo")
    para(
        doc,
        "O Município de Pelotas, por intermédio da Secretaria Municipal de Saúde, torna público que "
        "realizará Pregão Eletrônico, na forma da Lei nº 14.133/2021 e normas municipais, destinado "
        "à contratação do objeto abaixo.",
    )
    heading(doc, "2. Objeto")
    para(
        doc,
        "Fornecimento contínuo de oxigênio medicinal com comodato de cilindros, em 2 lotes "
        "(zona urbana e zona rural), conforme Anexo I — Termo de Referência.",
    )
    heading(doc, "3. Datas")
    meta_line(doc, "Início de acolhimento de propostas", "___/___/2026 às __:__")
    meta_line(doc, "Limite para recebimento", "___/___/2026 às __:__")
    meta_line(doc, "Sessão pública", "___/___/2026 às __:__")
    meta_line(doc, "Plataforma", "________________ (PNCP / sistema municipal)")
    heading(doc, "4. Critério / 5. Publicidade")
    para(doc, "Menor preço por lote. Observar o mínimo legal de 8 (oito) dias úteis para propostas.")
    heading(doc, "6. Anexos")
    for item in [
        "Anexo I — Termo de Referência;",
        "Anexo II — Minuta do Contrato;",
        "Anexo III — Modelo de Proposta;",
        "Anexo IV — Locais de Entrega;",
        "Anexo V — Declarações.",
    ]:
        bullet(doc, item)
    heading(doc, "7. Divulgação ampliada")
    para(
        doc,
        "Além do PNCP e do Portal do Município, serão encaminhados ofícios/e-mails a empresas do "
        "ramo de gases medicinais da região, com comprovação nos autos.",
    )
    heading(doc, "8. Disposições finais")
    para(
        doc,
        "Integram o edital a Lei nº 14.133/2021 e o Termo de Referência. Esclarecimentos: Setor de "
        "Licitações — e-mail/telefone _______________.",
    )
    signatures(doc, [["Pregoeiro(a) / Autoridade competente", "SMS"]])
    save(doc, P2 / "02.07-minuta-edital.docx")


def p2_contrato():
    doc = new_official_doc()
    title(doc, "Anexo II — Minuta do Contrato", "Pregão Eletrônico — Oxigênio Medicinal")
    para(
        doc,
        "CONTRATO Nº _____/2026 que entre si celebram o Município de Pelotas / SMS e a empresa "
        "____________________, vencedora do Pregão Eletrônico nº _____/2026, Lote(s) _____.",
        first_indent=False,
    )
    for t, b in [
        ("CLÁUSULA 1 — OBJETO", "Fornecimento contínuo de oxigênio medicinal com comodato, conforme TR e proposta."),
        ("CLÁUSULA 2 — VIGÊNCIA", "12 (doze) meses, prorrogável na forma da Lei nº 14.133/2021."),
        ("CLÁUSULA 3 — PREÇO E PAGAMENTO", "Preço unitário de R$ ______ / m³ (Lote ____). Pagamento mensal por medição atestada."),
        ("CLÁUSULA 4 — OBRIGAÇÕES", "A CONTRATADA cumprirá o TR. A CONTRATANTE pagará e fiscalizará."),
        ("CLÁUSULA 5 — FISCALIZAÇÃO", "Gestor: ________. Fiscal: ________."),
        ("CLÁUSULA 6 — SANÇÕES", "Arts. 155 a 163 da Lei nº 14.133/2021 e edital."),
        ("CLÁUSULA 7 — RESCISÃO", "Nos casos legais, com contraditório e ampla defesa."),
        ("CLÁUSULA 8 — FORO", "Comarca de Pelotas/RS."),
    ]:
        heading(doc, t)
        para(doc, b)
    signatures(
        doc,
        [
            ["CONTRATANTE — Município de Pelotas / SMS", ""],
            ["CONTRATADA", ""],
            ["Testemunhas", ""],
        ],
    )
    save(doc, P2 / "02.08-anexo-ii-minuta-contrato.docx")


def p2_proposta():
    doc = new_official_doc()
    title(doc, "Anexo III — Modelo de Proposta de Preços")
    meta_line(doc, "Pregão Eletrônico nº", "_____/2026")
    meta_line(doc, "Razão social / CNPJ", "_____________________ / ______________")
    meta_line(doc, "Endereço / contato", "______________________________________________")
    simple_table(
        doc,
        ["Lote", "Descrição", "Unidade", "Qtde anual", "Preço unit. (R$)", "Preço total (R$)"],
        [
            ["1", "Oxigênio medicinal — Zona Urbana", "m³", "2.184,0", "", ""],
            ["2", "Oxigênio medicinal — Zona Rural", "m³", "288,0", "", ""],
        ],
    )
    para(
        doc,
        "Declaramos que nos preços estão inclusos comodato de cilindros, manutenção, logística, "
        "tributos e demais encargos. Validade da proposta: 60 dias. Prazo de início: até 10 dias "
        "úteis da ordem de serviço/empenho.",
    )
    signatures(doc, [["Representante legal", "Nome / CPF / cargo"]])
    save(doc, P2 / "02.09-anexo-iii-modelo-proposta.docx")


def p2_divulgacao():
    doc = new_official_doc()
    title(doc, "Ofício-Circular — Divulgação Ampliada do Pregão")
    meta_line(doc, "Ofício nº", "_____/2026 — SMS/Licitações")
    para(doc, "À empresa: _______________________________", first_indent=False)
    para(doc, "Prezados(as),", first_indent=False)
    para(
        doc,
        "Informamos que o Município de Pelotas, por meio da Secretaria Municipal de Saúde, "
        "publicará/publicou Pregão Eletrônico para fornecimento de oxigênio medicinal com comodato "
        "de cilindros para UBSs e UBAIs (zonas urbana e rural).",
    )
    para(
        doc,
        f"O certame anterior (Processo nº {D['processo_origem']}) restou deserto. Visando ampliar "
        "a competitividade, convidamos V.Sas. a participar do novo certame, cujos documentos estão "
        "disponíveis no PNCP e no Portal da Prefeitura.",
    )
    meta_line(doc, "Link PNCP / edital", "________________________________")
    meta_line(doc, "Data da sessão", "___/___/2026")
    meta_line(doc, "Contato", "________________________________")
    para(doc, "Atenciosamente,", first_indent=False)
    signatures(doc, [["Setor de Licitações / Pregoeiro(a)", "SMS"]])
    heading(doc, "Controle de envio (anexar comprovantes)")
    simple_table(
        doc,
        ["Empresa", "E-mail/WhatsApp", "Data", "Comprovante"],
        [["", "", "", "[ ]"], ["", "", "", "[ ]"], ["", "", "", "[ ]"], ["", "", "", "[ ]"], ["", "", "", "[ ]"]],
    )
    save(doc, P2 / "02.10-oficio-divulgacao-ampliada.docx")


def p2_autorizacao_edital():
    doc = new_official_doc()
    title(doc, "Despacho de Autorização para Publicação do Edital", "Pregão Eletrônico — Oxigênio Medicinal")
    para(
        doc,
        "Autorizo a publicação do Edital do Pregão Eletrônico nº _____/2026, destinado ao "
        "fornecimento de oxigênio medicinal com comodato de cilindros, com fulcro na Lei nº "
        "14.133/2021, determinando a divulgação no PNCP, no Portal da Transparência e o envio "
        "de ofícios a empresas do ramo, bem como a observância do prazo mínimo legal de "
        "apresentação de propostas.",
    )
    signatures(doc, [[D["secretaria"], "Secretária Municipal de Saúde"]])
    save(doc, P2 / "02.11-despacho-autorizacao-publicacao-edital.docx")


# =============================================================================
# 03 — ANEXOS
# =============================================================================

def anexo_locais():
    doc = new_official_doc()
    title(doc, "Anexo — Locais de Entrega", "UBSs e UBAIs — Zona Urbana e Zona Rural")
    heading(doc, "LOTE 1 — Zona Urbana")
    simple_table(doc, ["Nº", "Unidade", "Endereço", "Telefone"], [list(r) for r in UNIDADES_URBANAS])
    heading(doc, "LOTE 2 — Zona Rural")
    simple_table(
        doc,
        ["Nº", "Unidade", "Endereço", "Distância", "Telefone"],
        [list(r) for r in UNIDADES_RURAIS],
    )
    para(
        doc,
        "As distâncias da zona rural são referenciais (trajeto até a sede da SMS) e destinam-se "
        "apenas ao planejamento logístico.",
        first_indent=False,
    )
    save(doc, ANX / "03.01-locais-de-entrega-ubs-ubai.docx")


def anexo_estoque_modelo():
    doc = new_official_doc()
    title(doc, "Modelo — Relatório de Estoque Atual de Oxigênio Medicinal", "Preencher e anexar à justificativa da emergência")
    meta_line(doc, "Data da apuração", "___/___/2026")
    meta_line(doc, "Responsável", "_______________________________")
    simple_table(
        doc,
        ["Unidade / região", "Cilindros cheios", "Volume approx. (m³)", "Autonomia estimada (dias)", "Obs."],
        [
            ["Zona Urbana — consolidado", "", "", "", ""],
            ["UBAIs — consolidado", "", "", "", ""],
            ["Zona Rural — consolidado", "", "", "", ""],
            ["TOTAL", "", "", "", ""],
        ],
    )
    para(
        doc,
        "Conclusão: o estoque atual assegura aproximadamente _____ dias de consumo, insuficiente "
        "para aguardar novo certame sem risco assistencial.",
    )
    signatures(doc, [["Responsável pelo estoque / farmácia / logística", "SMS"]])
    save(doc, ANX / "03.02-modelo-relatorio-estoque.docx")


def write_readme():
    text = """# Pacote documental oficial — Oxigênio Medicinal (SMS Pelotas)

Documentos no **padrão oficial da Secretaria Municipal de Saúde** (brasão do Município no cabeçalho, caixas no estilo do DFD e logo da Prefeitura no rodapé), para tramitação **simultânea** das Opções **C** e **A**.

## Estrutura

```
docs/
  assets/                      # Brasão e logo oficiais
  00-INSTRUCAO-GERAL/          # Despacho, cronograma, índice
  01-PROCESSO-DISPENSA-EMERGENCIAL/   # 15 peças
  02-PROCESSO-PREGAO-ELETRONICO/      # 11 peças
  03-ANEXOS-COMUNS/            # Locais + estoque
```

## Ordem de uso

1. `00-INSTRUCAO-GERAL/00.01-despacho-encaminhamento-opcoes-a-e-c.docx`
2. Completar estoque (`03-ANEXOS-COMUNS/03.02-...`) e cotações
3. Tramitar o Volume 1 (emergencial) com prioridade
4. Em paralelo, publicar o Volume 2 (pregão)

## Regenerar

```bash
pip install python-docx
python3 scripts/gerar_documentos.py
```
"""
    (DOCS / "README.md").write_text(text, encoding="utf-8")
    (ROOT / "README.md").write_text(
        """# Processos-Documentos — Oxigênio Medicinal (SMS Pelotas)

Pacote documental **oficial** (brasão + formatação SMS) para andamento simultâneo de:

1. **Dispensa emergencial** (Art. 75, VIII)
2. **Pregão eletrônico** (solução definitiva)

Comece por [`docs/README.md`](docs/README.md) e pela pasta [`docs/00-INSTRUCAO-GERAL/`](docs/00-INSTRUCAO-GERAL/).
""",
        encoding="utf-8",
    )


def main():
    reset_tree()
    write_readme()
    # 00
    doc_despacho_geral()
    doc_cronograma()
    doc_indice()
    # 01
    p1_capa()
    p1_dfd()
    p1_etp()
    p1_riscos()
    p1_tr()
    p1_justificativa()
    p1_precos()
    p1_razao()
    p1_hab()
    p1_declaracoes()
    p1_contrato()
    p1_parecer_solicitacao()
    p1_autorizacao()
    p1_aviso_pncp()
    p1_extrato_contrato()
    # 02
    p2_capa()
    p2_dfd()
    p2_etp()
    p2_riscos()
    p2_tr()
    p2_precos()
    p2_edital()
    p2_contrato()
    p2_proposta()
    p2_divulgacao()
    p2_autorizacao_edital()
    # 03
    anexo_locais()
    anexo_estoque_modelo()
    print("\nPacote oficial regenerado.")


if __name__ == "__main__":
    main()
