"""Gerador do DFD Revisado — PICSPel Materiais de Consumo."""

from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH

from .dados import (
    DIFERENCA_NOMINAL,
    ITENS_CONSUMO,
    META,
    SALDO_RESIDUAL_EMENDA,
    SUBTOTAL_MATERIAIS,
    TOTAL_CONSOLIDADO_EMENDA,
    TOTAL_CONSUMO,
    TOTAL_UNIDADES,
    fmt_brl,
    fmt_int,
)
from .docx_helpers import (
    add_body_p,
    add_callout,
    add_section_heading,
    add_signatures_block,
    add_styled_table,
    add_subheading,
    create_institutional_doc,
)


def gerar_dfd(output_path: Path) -> Path:
    doc = create_institutional_doc(
        "DOCUMENTO DE FORMALIZAÇÃO DA DEMANDA nº [___]/2026 – PICSPel (MATERIAIS DE CONSUMO)",
        "Processo Administrativo / Memorando nº [______]/2026  |  Emenda Parlamentar nº 43530013  |  Elemento 3.3.90.30",
    )

    add_body_p(
        doc,
        META["unidade_demandante"] + ".",
        bold_prefix="UNIDADE DEMANDANTE: ",
        space_after=6,
    )

    # 1 OBJETO
    add_section_heading(doc, "1 OBJETO")
    add_body_p(
        doc,
        "Aquisição pontual de materiais de consumo destinados à implementação e ao fortalecimento "
        "das Práticas Integrativas e Complementares em Saúde (Programa PICSPel) na rede de "
        "Atenção Primária à Saúde do Município de Pelotas/RS, compreendendo 15 (quinze) itens de "
        "insumos para escalda-pés, auriculoterapia, aromaterapia, práticas corporais/meditação e "
        "apoio educativo/logístico (bacias para escalda-pés, protetores descartáveis para bacias, "
        "placas com esferas de cristal para auriculoterapia, pinças Adson, óleos essenciais puros "
        "de lavanda, hortelã-pimenta, melaleuca e laranja doce, óleo vegetal carreador de semente "
        "de uva, frascos âmbar 5 ml com gotejador, inaladores nasais portáteis, tapetes de yoga "
        "em EVA, apalpadores com mola padrão Nogier, plásticos para plastificação A4 e sal grosso "
        "marinho iodado), para abastecimento das 08 (oito) equipes eMulti, das Unidades Básicas "
        "de Saúde (UBS) de referência e para suporte aos cursos de formação previstos no âmbito "
        f"da Emenda Parlamentar nº {META['emenda']}, no valor total estimado de "
        f"{fmt_brl(TOTAL_CONSUMO)} (quarenta e cinco mil, quinhentos e vinte reais).",
    )

    add_subheading(doc, "1.1 TIPO DE OBJETO")
    for marcado, rotulo in (
        (True, "Material de Consumo (Elemento 3.3.90.30)"),
        (False, "Material Permanente / Equipamento (Elemento 4.4.90.52)"),
        (False, "Serviço continuado com dedicação exclusiva de mão de obra"),
        (False, "Serviço continuado sem dedicação exclusiva de mão de obra"),
        (False, "Serviço não continuado (Elemento 3.3.90.39)"),
    ):
        marca = " X " if marcado else "   "
        add_body_p(doc, f"({marca}) {rotulo}", justify=False, space_after=1.5, font_size=10.0)

    add_body_p(
        doc,
        "O presente DFD contempla exclusivamente os 15 (quinze) itens classificados como material "
        f"de consumo (elemento 3.3.90.30, total de {fmt_brl(TOTAL_CONSUMO)}). As demais parcelas "
        f"da Emenda Parlamentar nº {META['emenda']} tramitam em processos próprios e autônomos, "
        "sem qualquer sobreposição de itens: (i) Materiais Permanentes (elemento 4.4.90.52 — "
        f"5 itens, valor estimado de {fmt_brl(META['valor_processo_permanentes'])}); e "
        "(ii) Serviços de Capacitação presencial em Auriculoterapia 80h (elemento 3.3.90.39 — "
        f"turma de 30 a 50 vagas, valor estimado de {fmt_brl(META['valor_processo_servicos'])}).",
        bold_prefix="Observação (autonomia entre processos da Emenda nº 43530013): ",
        space_after=6,
    )

    # 2 FORMA DE CONTRATAÇÃO SUGERIDA
    add_section_heading(doc, "2 FORMA DE CONTRATAÇÃO SUGERIDA")
    for marcado, rotulo in (
        (True, "Pregão Eletrônico (aquisição pontual — sem Sistema de Registro de Preços)"),
        (False, "Pregão Eletrônico com Sistema/Ata de Registro de Preços"),
        (False, "Dispensa com Disputa Eletrônica"),
        (False, "Adesão à Ata do Consórcio Público do Extremo Sul"),
        (False, "Adesão à Ata de Registro de Preços de Outro(s) Órgão(s)"),
    ):
        marca = " X " if marcado else "   "
        add_body_p(doc, f"({marca}) {rotulo}", justify=False, space_after=1.5, font_size=10.0)

    add_body_p(
        doc,
        "Menor preço por item, com adjudicação por item e participação exclusiva para "
        "Microempresas (ME) e Empresas de Pequeno Porte (EPP) em todos os 15 itens, nos termos do "
        "art. 48, inciso I, da Lei Complementar nº 123/2006 c/c art. 4º da Lei Federal nº "
        "14.133/2021 (valor máximo por item de R$ 9.500,00 — amplamente inferior ao teto legal "
        "de R$ 80.000,00 por item).",
        bold_prefix="Critério de julgamento e tratamento diferenciado ME/EPP: ",
        space_after=4,
    )
    add_body_p(
        doc,
        "Embora o valor global estimado desta demanda seja inferior ao limite atualizado de "
        "dispensa de licitação por valor (art. 75, II, da Lei nº 14.133/2021), adota-se o Pregão "
        "Eletrônico com julgamento por item para afastar qualquer risco de fracionamento de "
        "despesas de mesma natureza/ramo de atividade na SMS no exercício de 2026 (art. 75, § 1º), "
        "assegurar ampla competitividade entre fornecedores especializados e garantir máxima "
        "publicidade no Portal Nacional de Contratações Públicas (PNCP). A existência de atas "
        "vigentes (Consórcio Público do Extremo Sul ou outros órgãos) deverá ser certificada "
        "previamente à publicação do edital.",
        bold_prefix="Nota técnica sobre a modalidade: ",
    )

    # 3 JUSTIFICATIVA DA CONTRATAÇÃO
    add_section_heading(doc, "3 JUSTIFICATIVA DA CONTRATAÇÃO")
    add_body_p(
        doc,
        "A formalização desta demanda destina-se a viabilizar a oferta segura, regular e "
        "qualificada de Práticas Integrativas e Complementares em Saúde (PICS) na Atenção Primária "
        "do Município de Pelotas/RS, em conformidade com a Política Nacional de Práticas "
        "Integrativas e Complementares (PNPIC) e com o planejamento local do Programa PICSPel, "
        f"utilizando recursos da Emenda Parlamentar nº {META['emenda']}.",
    )
    add_body_p(
        doc,
        "Os insumos de consumo relacionados são indispensáveis à realização de auriculoterapia, "
        "aromaterapia, escalda-pés, práticas corporais/meditação e ações educativas nas 08 (oito) "
        "equipes eMulti já estruturadas (Centro/Porto, Fragata 1 e 2, Três Vendas 1 e 2, Areal, "
        "Colônia e Praias) e nas UBS de referência, além de fornecer suporte material aos cursos "
        "de formação previstos (30 a 50 vagas, organizados em até 2 grupos de até 25 participantes "
        "nas atividades práticas). A aquisição pontual por Pregão Eletrônico com julgamento por "
        "item — sem adoção do Sistema de Registro de Preços — justifica-se pelo dimensionamento "
        "certo dos quantitativos vinculados à emenda parlamentar e assegura competitividade, "
        "isonomia e economicidade.",
    )
    add_body_p(
        doc,
        "[A INFORMAR: número do item no Plano de Contratações Anual (PCA) 2026 da SMS/DAP ou "
        "justificativa formal de inclusão excepcional vinculada à Emenda Parlamentar nº 43530013].",
        bold_prefix="Alinhamento ao PCA 2026: ",
    )

    # 4 DESCRIÇÃO DO OBJETO
    add_section_heading(doc, "4 DESCRIÇÃO DO OBJETO")
    add_body_p(
        doc,
        "Os itens abaixo possuem especificações técnicas objetivas, sem direcionamento de marca "
        "(art. 41, I, da Lei nº 14.133/2021), alinhadas integralmente ao Estudo Técnico Preliminar "
        "(ETP) e ao Termo de Referência (TR), incluindo exigências de validade mínima na entrega "
        "e regularização sanitária ANVISA somente quando tecnicamente aplicáveis ao item:",
    )

    rows = [
        (
            it.codigo,
            it.especificacao,
            it.unidade,
            fmt_int(it.quantidade),
            fmt_brl(it.valor_unitario),
            fmt_brl(it.valor_total),
        )
        for it in ITENS_CONSUMO
    ]
    add_styled_table(
        doc,
        headers=("Item", "Especificação Técnica", "Unid.", "Qtd.", "Valor unit. (R$)", "Valor total (R$)"),
        rows=rows,
        col_widths_cm=(1.0, 9.7, 1.2, 1.3, 1.9, 1.9),
        alignments=(
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ),
        font_size=8.0,
        total_row=(
            "VALOR TOTAL ESTIMADO (15 itens)",
            "",
            "—",
            fmt_int(TOTAL_UNIDADES),
            "—",
            fmt_brl(TOTAL_CONSUMO),
        ),
        merge_total_cols=2,
    )

    add_callout(
        doc,
        "Nota de Revisão Aritmética",
        f"Na versão preliminar do DFD e do ETP, constava como total estimado o valor nominal de "
        f"{fmt_brl(META['valor_nominal_original_pdf'])}. Conferido cada um dos 15 itens "
        f"(quantidades × valores unitários referenciais), a soma aritmética exata é de "
        f"{fmt_brl(TOTAL_CONSUMO)} (diferença a menor de {fmt_brl(DIFERENCA_NOMINAL)}). Todos os "
        "instrumentos deste processo (DFD, ETP e TR) adotam o somatório real conferido de "
        f"{fmt_brl(TOTAL_CONSUMO)}, mantendo inalteradas as quantidades e valores unitários "
        "referenciais, os quais deverão ser validados na pesquisa formal de preços (art. 23 da "
        "Lei nº 14.133/2021).",
    )

    # 5 DOTAÇÃO ORÇAMENTÁRIA
    add_section_heading(doc, "5 DOTAÇÃO ORÇAMENTÁRIA E CONSOLIDAÇÃO DA EMENDA")
    dot_rows = [
        ("Processo Administrativo / Memorando", "[A INFORMAR: nº do Processo / Memorando / 2026]"),
        ("Emenda Parlamentar", f"Emenda nº {META['emenda']} (valor global: {fmt_brl(META['teto_emenda_global'])})"),
        ("Projeto / Atividade", "[A INFORMAR — Atenção Primária / Programa PICSPel]"),
        ("Fonte de Recurso", "[A INFORMAR — vinculada à Emenda Parlamentar nº 43530013]"),
        ("Reduzido / Dotação Orçamentária", "[A INFORMAR — validar junto à Contabilidade / CCGOF]"),
        ("Previsão no PCA 2026", "[A INFORMAR — nº do item no PCA ou justificativa de inclusão]"),
        ("Classificação da Despesa", META["elemento_desc"]),
        ("Valor Estimado deste Processo (Consumo)", f"{fmt_brl(TOTAL_CONSUMO)} (soma real conferida dos 15 itens)"),
        ("Teto da Cota de Materiais na Emenda", f"{fmt_brl(META['teto_emenda_materiais'])} (cota conjunta de materiais)"),
        (
            "Consolidação dos 3 Processos da Emenda",
            f"Consumo ({fmt_brl(TOTAL_CONSUMO)}) + Permanentes ({fmt_brl(META['valor_processo_permanentes'])}) "
            f"= {fmt_brl(SUBTOTAL_MATERIAIS)} em materiais  |  Serviços ({fmt_brl(META['valor_processo_servicos'])})  |  "
            f"Total Consolidado: {fmt_brl(TOTAL_CONSOLIDADO_EMENDA)} (saldo residual de {fmt_brl(SALDO_RESIDUAL_EMENDA)})",
        ),
    ]
    add_styled_table(
        doc,
        headers=("Campo", "Informação"),
        rows=dot_rows,
        col_widths_cm=(5.8, 11.2),
        alignments=(WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT),
        font_size=8.8,
    )
    add_callout(
        doc,
        "Validação Contábil Prévia (CCGOF)",
        "Por tratar-se de aquisição pontual (não SRP), a indicação definitiva da dotação e a "
        "Nota de Reserva Orçamentária são obrigatórias antes da publicação do edital (art. 150 da "
        "Lei nº 14.133/2021). Recomenda-se também à Contabilidade/Patrimônio confirmar os "
        "subelementos definitivos de 3.3.90.30 para os 15 itens e ratificar a dispensa de "
        "tombamento patrimonial nos termos da Portaria STN nº 448/2002, bem como conciliar a "
        "parcela de capital (4.4.90.52) do processo correlato de Materiais Permanentes.",
    )

    # 6 LOCAL DE ENTREGA
    add_section_heading(doc, "6 LOCAL DE ENTREGA")
    add_body_p(
        doc,
        f"Os materiais deverão ser entregues no {META['almoxarifado_endereco']}, "
        f"{META['almoxarifado_horario']}, mediante prévio agendamento pelo e-mail "
        f"{META['almoxarifado_email']} (ou em local urbano previamente indicado na Ordem de "
        "Fornecimento pela Diretoria de Atenção Primária / coordenação PICSPel), com posterior "
        "distribuição interna às 08 (oito) equipes eMulti ("
        + ", ".join(META["equipes_emulti"])
        + ") e às UBS de referência.",
    )

    # 7 PREVISÃO DE INÍCIO E PRAZO DE ENTREGA
    add_section_heading(doc, "7 PREVISÃO DE INÍCIO E PRAZO DE ENTREGA")
    add_body_p(
        doc,
        "Entrega integral por item adjudicado em até 30 (trinta) dias corridos, contados do "
        "recebimento da Nota de Empenho / Ordem de Fornecimento pela contratada (admitindo-se "
        "parcelamento em até 2 remessas apenas se formalmente autorizado pela SMS conforme "
        "necessidade de armazenamento). Os insumos com prazo de validade (Itens 03, 05 a 09 e 15) "
        "deverão ser entregues com validade mínima de 12 (doze) meses a contar da data da entrega "
        "(correspondendo a, no mínimo, 75% do prazo total de validade do fabricante). Após o "
        "recebimento definitivo, será realizada reunião técnica com as equipes eMulti e APS de "
        "referência para organização do armazenamento, controle de estoque e fluxos de utilização.",
    )

    # EQUIPE DE PLANEJAMENTO / FISCALIZAÇÃO
    add_section_heading(doc, "EQUIPE DE PLANEJAMENTO / FISCALIZAÇÃO")
    add_signatures_block(doc, autorizo_suffix=True)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path
