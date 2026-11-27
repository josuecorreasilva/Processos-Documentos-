"""Gerador do ETP Revisado — PICSPel Materiais de Consumo."""

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
    add_bullet,
    add_callout,
    add_section_heading,
    add_signatures_block,
    add_styled_table,
    add_subheading,
    create_institutional_doc,
)


def gerar_etp(output_path: Path) -> Path:
    doc = create_institutional_doc(
        "ESTUDO TÉCNICO PRELIMINAR – PICSPel (MATERIAIS DE CONSUMO)",
        "Lei Federal nº 14.133/2021 (art. 18, § 1º)  |  Emenda Parlamentar nº 43530013  |  Elemento 3.3.90.30",
    )

    # 1 IDENTIFICAÇÃO DO PROCESSO
    add_section_heading(doc, "1 IDENTIFICAÇÃO DO PROCESSO")
    add_body_p(
        doc,
        META["unidade_demandante"] + ".",
        bold_prefix="1.1 Unidade responsável pela contratação: ",
    )
    add_body_p(
        doc,
        "Aquisição pontual de materiais de consumo destinados à implementação e ao fortalecimento "
        "das Práticas Integrativas e Complementares em Saúde (PICS) na rede de Atenção Primária à "
        "Saúde do Município de Pelotas/RS, compreendendo 15 (quinze) itens de insumos essenciais "
        "para uso pelas 08 (oito) equipes eMulti e Unidades Básicas de Saúde (UBS) de referência, "
        "bem como para suporte aos cursos de formação previstos no âmbito da Emenda Parlamentar "
        f"nº {META['emenda']}. Os materiais destinam-se às práticas de auriculoterapia, "
        "aromaterapia, escalda-pés, práticas corporais/meditação e atividades educativas "
        "correlatas, contribuindo para a ampliação da oferta de PICS, a humanização do cuidado e "
        "a resolutividade da Atenção Primária.",
        bold_prefix="1.2 Objeto da contratação: ",
    )
    add_body_p(
        doc,
        f"Processo Administrativo / Memorando nº [______]/2026 – DFD nº [___]/2026 – "
        f"Emenda Parlamentar nº {META['emenda']}.",
        bold_prefix="1.3 Número do processo: ",
    )
    add_body_p(
        doc,
        f"Emenda Parlamentar nº {META['emenda']}, no valor global de "
        f"{fmt_brl(META['teto_emenda_global'])}, estruturada em {fmt_brl(META['teto_emenda_materiais'])} "
        f"para materiais e {fmt_brl(META['teto_emenda_servicos'])} para outros serviços de "
        "terceiros – pessoa jurídica (capacitação). O presente ETP abrange exclusivamente a "
        f"parcela de Materiais de Consumo (elemento 3.3.90.30, soma real conferida de "
        f"{fmt_brl(TOTAL_CONSUMO)} para 15 itens). As demais parcelas tramitam em processos "
        "próprios e autônomos: Materiais Permanentes (elemento 4.4.90.52, 5 itens — "
        f"{fmt_brl(META['valor_processo_permanentes'])}) e Serviços de Capacitação em "
        f"Auriculoterapia (elemento 3.3.90.39, 80h — {fmt_brl(META['valor_processo_servicos'])}), "
        f"totalizando {fmt_brl(TOTAL_CONSOLIDADO_EMENDA)} e preservando saldo residual de "
        f"{fmt_brl(SALDO_RESIDUAL_EMENDA)} frente ao teto global da emenda.",
        bold_prefix="1.4 Fonte de recurso e consolidação da emenda: ",
    )

    # 2 DESCRIÇÃO DA NECESSIDADE
    add_section_heading(doc, "2 DESCRIÇÃO DA NECESSIDADE")
    add_body_p(
        doc,
        "A Política Nacional de Práticas Integrativas e Complementares (PNPIC) e a organização "
        "local do Programa PICSPel demandam insumos específicos, padronizados e sanitariamente "
        "seguros para a oferta regular de PICS na Atenção Primária à Saúde. A Secretaria Municipal "
        "de Saúde necessita adquirir o conjunto de 15 itens de consumo descritos na Lista de "
        "Materiais do PICSPel, de modo a municiar as 08 (oito) equipes eMulti e as UBS de "
        "referência, bem como assegurar suporte material aos cursos de formação previstos "
        "(30 a 50 vagas, com subdivisão em até 2 grupos de até 25 participantes nas etapas práticas).",
    )
    add_body_p(
        doc,
        "A ausência desses insumos impede ou restringe a oferta de auriculoterapia, aromaterapia, "
        "escalda-pés e práticas corporais/meditação, reduzindo a resolutividade e a equidade do "
        "cuidado na APS. A demanda foi dimensionada a partir do planejamento técnico do PICSPel, "
        "considerando as 08 (oito) equipes eMulti já contempladas e a necessidade de estoque "
        "operacional nas UBS e na coordenação do programa:",
    )
    for eq in META["equipes_emulti"]:
        add_bullet(doc, eq, bold_prefix="Equipe eMulti: ")
    add_bullet(
        doc,
        "Eventual inclusão futura de até 04 (quatro) equipes adicionais a definir — sem alteração "
        "automática dos quantitativos deste ETP; eventual suplementação dependerá de novo "
        "planejamento e disponibilidade orçamentária.",
        bold_prefix="Nota sobre expansão futura: ",
    )
    add_body_p(
        doc,
        "Após a entrega e o recebimento definitivo dos materiais, será realizada reunião técnica "
        "com as equipes eMulti e com as equipes da APS de referência para organizar o "
        "armazenamento, o controle de estoque (inclusive monitoramento de validade de óleos e "
        "placas) e os fluxos de utilização, conforme encaminhamento do Programa PICSPel.",
    )

    # 3 DA CONTRATAÇÃO E DO PLANEJAMENTO
    add_section_heading(doc, "3 DA CONTRATAÇÃO E DO PLANEJAMENTO")
    add_body_p(
        doc,
        "A contratação tem por finalidade a aquisição pontual dos 15 itens de materiais de "
        f"consumo necessários à execução do PICSPel com recursos da Emenda Parlamentar nº {META['emenda']}. "
        "Os itens apresentam padrões de desempenho e qualidade objetivamente definíveis por "
        "especificações usuais de mercado, não se enquadram como bens de luxo (art. 20) e "
        "caracterizam-se como bens comuns de consumo, nos termos do art. 6º, XIII, da Lei "
        "Federal nº 14.133/2021.",
    )
    add_body_p(
        doc,
        "Pregão Eletrônico (art. 6º, XLI, e art. 28, I, da Lei nº 14.133/2021), critério de "
        "julgamento de menor preço por item, com adjudicação por item e participação exclusiva "
        "para Microempresas e Empresas de Pequeno Porte (ME/EPP) em todos os 15 itens (art. 48, "
        "inciso I, da LC nº 123/2006 c/c art. 4º da Lei nº 14.133/2021), dado que todos os itens "
        "possuem valor estimado individual inferior a R$ 80.000,00 (maior item: R$ 9.500,00). "
        "Por tratar-se de compra com entrega integral por item (em até 30 dias corridos) sem "
        "obrigações futuras complexas, o termo de contrato poderá ser substituído pela Nota de "
        "Empenho / Ordem de Fornecimento, na forma do art. 95, II, da Lei nº 14.133/2021.",
        bold_prefix="Forma de contratação sugerida: ",
    )
    add_body_p(
        doc,
        "Justifica-se a aquisição pontual — e não o Sistema de Registro de Preços (SRP) — porque: "
        "(I) os quantitativos estão vinculados a recurso de emenda parlamentar com objeto e "
        "cronograma definidos; (II) a demanda foi previamente dimensionada para as 08 equipes "
        "eMulti, UBS de referência e cursos de formação; e (III) trata-se de implantação e "
        "estruturação com quantitativo certo e disponibilidade financeira definida, exigindo "
        "reserva orçamentária prévia nos termos do art. 150 da Lei nº 14.133/2021.",
        bold_prefix="Justificativa para não adoção do SRP: ",
    )
    add_body_p(
        doc,
        "[A INFORMAR: confirmar se a contratação consta do PCA 2026 da SMS/DAP e o respectivo "
        "número do item — caso não conste, anexar justificativa formal de inclusão excepcional "
        "vinculada à Emenda Parlamentar nº 43530013].",
        bold_prefix="Previsão no Plano de Contratações Anual (PCA 2026): ",
    )

    # 4 REQUISITOS DA CONTRATAÇÃO
    add_section_heading(doc, "4 REQUISITOS DA CONTRATAÇÃO")
    add_body_p(
        doc,
        "Todos os produtos deverão atender integralmente às especificações técnicas de cada item, "
        "ser novos, de primeiro uso, livres de vícios, defeitos de fabricação, oxidação ou "
        "avarias, entregues em embalagem original íntegra e lacrada, com rotulagem em língua "
        "portuguesa contendo identificação do fabricante/importador e, quando aplicável, lote, "
        "data de fabricação e prazo de validade. Não haverá exigência de marca ou modelo "
        "específico (art. 41, I, da Lei nº 14.133/2021).",
    )
    add_subheading(doc, "4.1 Requisitos técnicos específicos por família de produtos")
    add_bullet(
        doc,
        "bacias anatômicas (Item 01) em polipropileno ou polímero plástico rígido atóxico, "
        "lavável e resistente à água morna e desinfecção; protetores descartáveis em polietileno "
        "de uso único (Item 02) compatíveis com a bacia; e sal grosso marinho iodado em pacotes "
        "lacrados de 1 kg (Item 15), sem corantes ou fragrâncias artificiais, rotulado conforme "
        "padrões sanitários vigentes (RDC ANVISA nº 604/2022) e com validade mínima de 12 meses.",
        bold_prefix="Escalda-pés (Itens 01, 02 e 15): ",
    )
    add_bullet(
        doc,
        "placas com 70 pontos de cristal polido em fita adesiva microporosa hipoalergênica — sem "
        "vinculação a marca comercial — (Item 03) com regularização sanitária ANVISA e validade "
        "mínima de 12 meses na entrega; pinças Adson 12 cm retas sem dente (Item 04) e apalpadores "
        "com mola padrão Nogier (Item 13) em aço inoxidável cirúrgico autoclavável, com "
        "regularização sanitária vigente na ANVISA (notificação/cadastro conforme RDC nº 751/2022).",
        bold_prefix="Auriculoterapia (Itens 03, 04 e 13): ",
    )
    add_bullet(
        doc,
        "óleos essenciais 100% puros de Lavanda, Hortelã-pimenta/Peppermint, Melaleuca e Laranja "
        "Doce (Itens 05 a 08) e óleo vegetal carreador de semente de uva prensado a frio (Item 09), "
        "em frascos âmbar com lacre de inviolabilidade, rotulagem completa em português (incluindo "
        "espécie botânica, lote, fabricação e validade), regularização sanitária vigente na ANVISA "
        "(notificação, cadastro ou registro conforme enquadramento regulatório) e, para os óleos "
        "essenciais (Itens 05 a 08), laudo/certificado de análise ou cromatografia gasosa (CG) "
        "do lote atestando pureza de 100%; frascos âmbar 5 ml vazios com gotejador e tampa vedante "
        "(Item 10) e inaladores nasais portáteis vazios com pavio de algodão (Item 11), novos e "
        "higienizados.",
        bold_prefix="Aromaterapia (Itens 05 a 11): ",
    )
    add_bullet(
        doc,
        "tapetes de yoga em EVA antiderrapante, lavável, com espessura mínima de 5 mm (Item 12) "
        "e filmes duplos de plastificação térmica padrão A4, 0,07 mm / 175 micras (Item 14), "
        "compatíveis com a plastificadora térmica do processo correlato de Materiais Permanentes.",
        bold_prefix="Práticas corporais e apoio educativo/logístico (Itens 12 e 14): ",
    )

    add_subheading(doc, "4.2 Validade mínima, regularização ANVISA e habilitação proporcional")
    add_bullet(
        doc,
        "os itens perecíveis ou com prazo de validade controlado (Itens 03, 05, 06, 07, 08, 09 e "
        "15) deverão apresentar validade mínima de 12 (doze) meses a contar da data de entrega no "
        "Almoxarifado da SMS, correspondendo a pelo menos 75% do prazo total de validade do "
        "fabricante (recomendável ≥ 18 meses para os Itens 03 e 05 a 07). Itens entregues com "
        "validade inferior serão recusados e substituídos sem ônus em até 5 (cinco) dias úteis.",
        bold_prefix="Prazo de validade mínimo na entrega: ",
    )
    add_bullet(
        doc,
        "exigida de forma estritamente proporcional ao enquadramento regulatório de cada item: "
        "(a) notificação/cadastro/registro ANVISA obrigatório para os Itens 03, 04, 05 a 09 e 13 "
        "(além da conformidade sanitária de rotulagem/iodação do Item 15); (b) AFE/Licença "
        "Sanitária exigíveis apenas das licitantes vencedoras desses itens regulados, quando "
        "legalmente obrigatórias; (c) expressamente dispensadas para os Itens 01, 02, 10, 11, 12 "
        "e 14 (utilidades plásticas, embalagens/acessórios vazios, artigo esportivo e papelaria), "
        "evitando restrição indevida à competitividade.",
        bold_prefix="Regularização sanitária (ANVISA) diferenciada por item: ",
    )
    add_bullet(
        doc,
        "a Administração poderá solicitar ficha técnica/catálogo e, quando tecnicamente "
        "justificado, amostra do licitante provisoriamente classificado em 1º lugar (art. 41, II, "
        "da Lei nº 14.133/2021), no prazo de 5 (cinco) dias úteis. A contratada responderá pela "
        "substituição sem ônus de produtos defeituosos ou em desconformidade em até 5 (cinco) dias "
        "úteis da notificação. A habilitação jurídica, fiscal, trabalhista e econômico-financeira "
        "será proporcional ao valor e à baixa complexidade de cada item (arts. 62 a 70 da Lei nº "
        "14.133/2021).",
        bold_prefix="Amostras, substituição e habilitação: ",
    )

    # 5 ESTIMATIVA DAS QUANTIDADES
    add_section_heading(doc, "5 ESTIMATIVA DAS QUANTIDADES")
    add_body_p(
        doc,
        "As quantidades foram estimadas com base na Lista de Materiais para Aquisição do Programa "
        "PICSPel, considerando: (a) 08 equipes eMulti, com quantitativos unitários por equipe "
        "multiplicados por 8; (b) demanda das UBS de referência para auriculoterapia e aromaterapia "
        "(incluindo 4 kits bimestrais de óleos essenciais); e (c) suporte aos cursos de formação "
        "(30 a 50 vagas) e insumos de apoio da coordenação. Por tratar-se de aquisição pontual com "
        "quantitativos certos, a Administração contratará os totais adjudicados, observados o "
        "empenho e a disponibilidade orçamentária da emenda.",
    )

    rows_q = [
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
        rows=rows_q,
        col_widths_cm=(1.0, 9.7, 1.2, 1.3, 1.9, 1.9),
        alignments=(
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ),
        font_size=7.8,
        total_row=(
            "VALOR TOTAL ESTIMADO (soma real dos 15 itens)",
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
        "Nota de Conferência Aritmética",
        f"A soma nominal registrada na versão original dos PDFs ({fmt_brl(META['valor_nominal_original_pdf'])}) "
        f"apresentava divergência a maior de {fmt_brl(DIFERENCA_NOMINAL)} em relação ao somatório "
        f"exato das 15 linhas da tabela ({fmt_brl(TOTAL_CONSUMO)}). Este ETP adota o valor real "
        f"auditado de {fmt_brl(TOTAL_CONSUMO)}, mantendo integralmente os quantitativos e os "
        "valores unitários referenciais originais.",
    )

    add_subheading(doc, "5.1 Classificação por natureza da despesa, destinação e controle sanitário")
    rows_nat = [
        (
            it.codigo,
            it.nome_curto,
            "Consumo (3.3.90.30)",
            it.destinacao,
            it.anvisa,
            fmt_brl(it.valor_total),
        )
        for it in ITENS_CONSUMO
    ]
    add_styled_table(
        doc,
        headers=("Item", "Material de Consumo", "Natureza", "Destinação", "ANVISA / Conformidade", "Valor (R$)"),
        rows=rows_nat,
        col_widths_cm=(0.9, 4.4, 2.3, 3.2, 4.3, 1.9),
        alignments=(
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ),
        font_size=7.8,
        total_row=(
            "VALOR TOTAL DE MATERIAL DE CONSUMO (3.3.90.30)",
            "",
            "",
            "",
            "",
            fmt_brl(TOTAL_CONSUMO),
        ),
        merge_total_cols=5,
    )
    add_body_p(
        doc,
        "Todos os 15 itens constantes deste processo estão classificados na categoria de material "
        "de consumo (elemento 3.3.90.30), dispensando procedimentos de tombamento patrimonial, "
        "sem prejuízo de validação pela Contabilidade/Patrimônio quanto aos subelementos "
        "definitivos (3.3.90.30.__) e ao enquadramento contábil de itens de uso reiterado de "
        "baixo custo (bacias, pinças, apalpadores e tapetes) à luz da Portaria STN nº 448/2002 e "
        "da norma patrimonial municipal.",
        bold_prefix="Observação contábil/patrimonial: ",
    )

    # 6 LEVANTAMENTO DE MERCADO
    add_section_heading(doc, "6 LEVANTAMENTO DE MERCADO")
    add_body_p(
        doc,
        "Realizou-se levantamento preliminar de mercado com base na Lista de Materiais do PICSPel, "
        f"cujos valores unitários de referência totalizam {fmt_brl(TOTAL_CONSUMO)} para os 15 "
        "itens de consumo. Verificou-se a ampla existência de fornecedores locais, regionais e "
        "nacionais atuantes nos ramos de produtos médico-hospitalares/auriculoterapia, "
        "aromaterapia/óleos essenciais, utilidades plásticas/descartáveis, artigos esportivos, "
        "papelaria e gêneros alimentícios, aptos a ofertar os itens isoladamente ou em conjunto.",
    )
    add_body_p(doc, "Alternativas de mercado avaliadas:", bold_prefix="Análise comparativa — ")
    add_bullet(
        doc,
        "solução mais adequada, pois permite que empresas especializadas em cada segmento "
        "(saúde/auriculoterapia, aromaterapia, descartáveis, esportivo e papelaria) disputem "
        "apenas os itens de seu ramo, ampliando a competitividade, viabilizando a exclusividade "
        "ME/EPP e garantindo transparência no PNCP.",
        bold_prefix="(I) Licitação própria por Pregão Eletrônico com julgamento por item (RECOMENDADA): ",
    )
    add_bullet(
        doc,
        "inviável técnica e economicamente devido à heterogeneidade mercadológica dos 15 itens, "
        "que restringiria a competição a intermediários generalistas e elevaria custos.",
        bold_prefix="(II) Contratação em lote único / preço global (DESCARTADA): ",
    )
    add_bullet(
        doc,
        "deverá ser consultada antes da publicação do edital; caso exista ata vigente e "
        "comprovadamente mais vantajosa para item específico, este poderá ser excluído do certame "
        "mediante justificativa nos autos.",
        bold_prefix="(III) Adesão a Ata de Registro de Preços (Consórcio Público do Extremo Sul / outros órgãos): ",
    )
    add_bullet(
        doc,
        "embora o total estimado esteja abaixo do teto atualizado do art. 75, II, da Lei nº "
        "14.133/2021, a dispensa exigiria controle rigoroso de somatório de despesas do mesmo "
        "ramo de atividade na SMS no exercício (art. 75, § 1º), razão pela qual o Pregão "
        "Eletrônico oferece maior segurança jurídica e competição.",
        bold_prefix="(IV) Dispensa de licitação por valor (art. 75, II): ",
    )
    add_body_p(
        doc,
        "A pesquisa formal de preços será consolidada nos termos do art. 23 da Lei nº 14.133/2021 "
        "(PNCP / Painel de Preços, contratações similares da Administração Pública, mídia/tabelas "
        "especializadas e/ou pesquisa direta com fornecedores), adotando-se critério estatístico "
        "justificado (mediana ou média saneada), com Mapa Comparativo de Preços anexo ao processo "
        "antes da publicação do edital.",
    )

    # 7 ESTIMATIVA DO VALOR DA CONTRATAÇÃO
    add_section_heading(doc, "7 ESTIMATIVA DO VALOR DA CONTRATAÇÃO")
    add_body_p(
        doc,
        f"O valor total estimado desta contratação é de {fmt_brl(TOTAL_CONSUMO)} (quarenta e cinco "
        "mil, quinhentos e vinte reais), correspondente à soma aritmética exata dos 15 itens de "
        "material de consumo (1.508 unidades comerciais), plenamente compatível com a parcela de "
        f"{fmt_brl(META['teto_emenda_materiais'])} prevista para materiais na Emenda Parlamentar "
        f"nº {META['emenda']}. Os valores unitários são referenciais e serão validados/atualizados "
        "pela pesquisa oficial de preços (art. 23) previamente ao edital.",
    )

    emenda_rows = [
        ("Processo 1 — Serviços de Capacitação (Auriculoterapia 80h)", "3.3.90.39", "1", fmt_brl(META["valor_processo_servicos"]), fmt_brl(META["teto_emenda_servicos"])),
        ("Processo 2 — Materiais Permanentes (Equipamentos/Apoio)", "4.4.90.52*", "5", fmt_brl(META["valor_processo_permanentes"]), "Incluído na cota de R$ 60.000,00"),
        ("Processo 3 — Materiais de Consumo (este processo)", "3.3.90.30", "15", fmt_brl(TOTAL_CONSUMO), fmt_brl(META["teto_emenda_materiais"])),
    ]
    add_styled_table(
        doc,
        headers=("Processo (Emenda nº 43530013)", "Elemento", "Itens", "Valor Conferido", "Teto da Cota na Emenda"),
        rows=emenda_rows,
        col_widths_cm=(6.2, 2.0, 1.3, 3.0, 4.5),
        alignments=(
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ),
        font_size=8.2,
        total_row=(
            "TOTAL CONSOLIDADO DOS 3 PROCESSOS (Saldo residual: R$ 2.880,00)",
            "—",
            "21",
            fmt_brl(TOTAL_CONSOLIDADO_EMENDA),
            fmt_brl(META["teto_emenda_global"]),
        ),
    )
    add_body_p(
        doc,
        "*Observação contábil: o subtotal consolidado de materiais (Consumo R$ 45.520,00 + "
        "Permanentes R$ 11.600,00 = R$ 57.120,00) respeita a cota de R$ 60.000,00, cabendo à "
        "Contabilidade/CCGOF validar, no âmbito do processo correlato de Materiais Permanentes, a "
        "compatibilidade de empenho em Despesa de Capital (GND 4 / 4.4.90.52) na programação da "
        "Emenda nº 43530013.",
        font_size=9.0,
    )

    # 8 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO
    add_section_heading(doc, "8 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO")
    add_body_p(
        doc,
        "A solução consiste na aquisição pontual, mediante Pregão Eletrônico com julgamento por "
        "item, dos 15 itens de consumo necessários ao Programa PICSPel, com entrega integral por "
        "item em até 30 (trinta) dias corridos (admitido fracionamento em até 2 remessas apenas "
        "se autorizado pela SMS por conveniência de estocagem) no Almoxarifado da SMS (Av. "
        "Fernando Osório, nº 273, Bairro Três Vendas, Pelotas/RS), seguida de distribuição "
        "planejada às 08 equipes eMulti e às UBS de referência.",
    )
    add_body_p(
        doc,
        "O ciclo de vida do objeto compreende: planejamento e formalização da demanda; seleção "
        "competitiva de fornecedores; entrega com proteção física/térmica adequada; recebimento "
        "provisório e definitivo com conferência de integridade, validade, laudos/cromatografia "
        "dos óleos essenciais e regularização ANVISA; armazenamento protegido contra luz direta, "
        "calor e umidade; reunião de pactuação e distribuição às equipes; utilização assistencial "
        "e educativa (incluindo higienização e autoclavagem dos instrumentais de aço inoxidável); "
        "e descarte ambientalmente adequado de embalagens e resíduos conforme a legislação vigente.",
    )

    # 9 JUSTIFICATIVA PARA PARCELAMENTO
    add_section_heading(doc, "9 JUSTIFICATIVA PARA PARCELAMENTO")
    add_body_p(
        doc,
        "A contratação será realizada com parcelamento por item e julgamento por menor preço por "
        "item, nos termos do art. 40, § 2º, e do art. 47, II, da Lei Federal nº 14.133/2021, por "
        "ser tecnicamente viável e economicamente vantajoso: os 15 itens pertencem a segmentos "
        "mercadológicos distintos (dispositivos/instrumentais de saúde, óleos essenciais e vegetal, "
        "embalagens de vidro/inaladores, descartáveis plásticos, artigos esportivos, papelaria e "
        "sal marinho). O parcelamento por item amplia a competitividade, favorece a participação "
        "direta de fabricantes, distribuidores especializados e ME/EPP, e reduz preços finais. A "
        "reunião dos 15 itens em um único certame decorre da unidade de finalidade (fortalecimento "
        "do PICSPel sob a mesma emenda parlamentar e mesma classificação 3.3.90.30) e da "
        "racionalização administrativa.",
    )

    # 10 DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS
    add_section_heading(doc, "10 DEMONSTRATIVO DOS RESULTADOS PRETENDIDOS")
    for res in (
        "viabilizar a oferta segura e regular de PICS (auriculoterapia, aromaterapia, escalda-pés e práticas corporais/meditação) nas 08 equipes eMulti e nas UBS de referência;",
        "dar suporte material adequado à capacitação de 30 a 50 profissionais da rede nos cursos previstos pelo PICSPel;",
        "padronizar kits operacionais, controle de validade e fluxos de utilização entre todas as equipes contempladas;",
        "fortalecer a atenção integral, a humanização do cuidado e a resolutividade da Atenção Primária à Saúde;",
        f"aplicar de forma eficiente, econômica e transparente os recursos da Emenda Parlamentar nº {META['emenda']};",
        "obter preços vantajosos mediante competição por item no Pregão Eletrônico com incentivo às ME/EPP.",
    ):
        add_bullet(doc, res)

    # 11 PROVIDÊNCIAS PRÉVIAS AO CONTRATO
    add_section_heading(doc, "11 PROVIDÊNCIAS PRÉVIAS AO CONTRATO")
    for prov in (
        "Formalização do DFD revisado e aprovação pela autoridade competente (Secretária Municipal de Saúde);",
        "elaboração e aprovação do Termo de Referência com especificações completas, códigos CATMAT definitivos e cronograma de entrega;",
        "realização da pesquisa formal de preços (art. 23 da Lei nº 14.133/2021) e juntada do Mapa Comparativo de Preços;",
        "verificação prévia de atas vigentes (Consórcio Público do Extremo Sul / outros órgãos) para eventual adesão vantajosa;",
        "confirmação junto à Contabilidade/CCGOF da dotação, fonte de recurso, subelementos 3.3.90.30.__ e emissão da Nota de Reserva Orçamentária (art. 150);",
        "confirmação do item no PCA 2026 ou formalização da justificativa de inclusão excepcional;",
        "designação formal de Gestor e Fiscal da contratação e preparação do Almoxarifado da SMS para recebimento e estocagem adequada de óleos e insumos;",
        "planejamento da reunião pós-entrega com equipes eMulti e APS de referência para distribuição e controle de estoque/validades.",
    ):
        add_bullet(doc, prov)

    # 12 CONTRATAÇÕES CORRELATAS/INTERDEPENDENTES
    add_section_heading(doc, "12 CONTRATAÇÕES CORRELATAS/INTERDEPENDENTES")
    add_body_p(
        doc,
        f"No âmbito da mesma Emenda Parlamentar nº {META['emenda']}, tramitam dois processos "
        "correlatos e autônomos, sem sobreposição de escopo ou duplicidade de itens:",
    )
    add_bullet(
        doc,
        f"parcela de {fmt_brl(META['valor_processo_servicos'])} destinada à contratação de curso "
        "presencial de capacitação em Auriculoterapia (80h; 30 a 50 vagas). Não há "
        "interdependência impeditiva entre as licitações, recomendando-se alinhamento de "
        "cronogramas para que os insumos de consumo estejam disponíveis preferencialmente antes "
        "ou durante a realização das aulas práticas.",
        bold_prefix="Processo 1 — Serviços de Capacitação (3.3.90.39): ",
    )
    add_bullet(
        doc,
        f"parcela de {fmt_brl(META['valor_processo_permanentes'])} (5 itens: caixas de som, "
        "garrafas térmicas, maletas rígidas, jarras elétricas e plastificadora térmica A3/A4). "
        "Os processos são autônomos e tecnicamente complementares: a plastificadora térmica "
        "(0,05–0,07 mm) opera com os plásticos A4 0,07 mm (Item 14 deste ETP); as jarras "
        "elétricas e garrafas térmicas complementam as bacias e sal grosso do escalda-pés; e as "
        "maletas rígidas acondicionam os kits das 08 equipes eMulti.",
        bold_prefix="Processo 2 — Materiais Permanentes (4.4.90.52): ",
    )

    # 13 POSSÍVEIS IMPACTOS AMBIENTAIS
    add_section_heading(doc, "13 POSSÍVEIS IMPACTOS AMBIENTAIS")
    add_body_p(
        doc,
        "Os impactos ambientais decorrem principalmente da geração de resíduos de embalagens "
        "secundárias, protetores plásticos descartáveis de bacias, placas adesivas usadas e "
        "frascos de vidro vazios. Deverão ser rigorosamente observados a Política Nacional de "
        "Resíduos Sólidos (Lei nº 12.305/2010), as diretrizes municipais de coleta seletiva e "
        "logística de recicláveis (papelão, polietileno limpo e vidro) e, sempre que houver "
        "contato com pele lesada, secreções ou fluidos corporais, o Plano de Gerenciamento de "
        "Resíduos de Serviços de Saúde (PGRSS — RDC ANVISA nº 222/2018) das unidades.",
    )
    add_body_p(
        doc,
        "Os óleos essenciais e o óleo vegetal carreador deverão ser armazenados em local fresco, "
        "ventilado, ao abrigo da luz solar direta e de fontes de calor/ignição, em frascos âmbar "
        "bem vedados, prevenindo vazamentos, foto-oxidação e descarte precoce. A opção por "
        "instrumentais reutilizáveis autoclaváveis em aço inoxidável (pinças Adson e apalpadores "
        "Nogier) e bacias laváveis reduz significativamente o volume de resíduos descartáveis "
        "gerados ao longo do programa.",
    )

    # 14 ANÁLISE DE RISCOS
    add_section_heading(doc, "14 ANÁLISE DE RISCOS")
    riscos = [
        (
            "Risco 1: Atraso na entrega dos materiais",
            "Média",
            "Médio",
            "Adiamento do início/ampliação da oferta de PICS nas equipes eMulti/UBS e desalinhamento com os cursos de capacitação.",
            "Prazo de entrega objetivo (até 30 dias corridos da NE/OF), prévio agendamento no Almoxarifado, sanções claras no TR e acompanhamento ativo pelo fiscal.",
            "Notificação formal, aplicação das sanções da Lei nº 14.133/2021 e do Decreto Municipal nº 6.799/2023 e, se necessário, ajuste de cronograma das turmas.",
        ),
        (
            "Risco 2: Entrega em desacordo com as especificações técnicas",
            "Média",
            "Alto",
            "Inadequação clínica/educativa do insumo, risco ao usuário da APS e necessidade de retrabalho logístico.",
            "Especificações técnicas objetivas sem direcionamento de marca no TR, exigência de ficha técnica, faculdade de amostra (art. 41, II) e recebimento provisório/definitivo criterioso.",
            "Recusa formal do lote desconforme, substituição integral às expensas da contratada em até 5 dias úteis e instauração de processo sancionatório quando cabível.",
        ),
        (
            "Risco 3: Itens desertos, fracassados ou com variação excessiva de preços",
            "Média",
            "Médio",
            "Lacuna na composição dos kits das equipes eMulti/UBS ou necessidade de repetição do certame.",
            "Julgamento por item, exclusividade ME/EPP, especificações com tolerâncias dimensionais razoáveis, pesquisa formal de preços atualizada (art. 23) e ampla divulgação no PNCP.",
            "Verificação de atas vigentes para adesão, republicação dos itens remanescentes ou adoção de contratação direta legalmente cabível (art. 75, III), além de remanejamento interno temporário.",
        ),
        (
            "Risco 4: Armazenamento ou controle de estoque inadequado pós-entrega",
            "Média",
            "Médio",
            "Deterioração, oxidação, quebra de frascos de vidro âmbar, vazamento de óleos ou perda por vencimento.",
            "Reunião técnica pós-entrega com equipes eMulti e APS, definição de local adequado (protegido de calor e luz) e planilha/sistema de controle de lotes e validades.",
            "Orientação corretiva imediata, apuração de responsabilidade funcional quando cabível e redistribuição preventiva entre equipes antes do vencimento.",
        ),
        (
            "Risco 5: Entrega de insumos com validade exígua, irregularidade ANVISA ou pureza comprometida (óleos)",
            "Média",
            "Alto",
            "Risco sanitário/alérgico aos usuários do SUS, vencimento prematuro em estoque ou aquisição de óleos essenciais sintéticos/adulterados.",
            "Exigência expressa no TR de validade mínima de 12 meses (≥ 75% da vida útil), comprovação de regularização ANVISA nos itens aplicáveis e apresentação de laudo/cromatografia gasosa (CG) dos óleos essenciais.",
            "Recusa imediata no recebimento, exigência de troca em até 5 dias úteis sem ônus, comunicação à Vigilância Sanitária em caso de suspeita de irregularidade e aplicação de penalidades contratuais.",
        ),
    ]
    for titulo, prob, imp, dano, prev, cont in riscos:
        add_subheading(doc, titulo)
        add_body_p(
            doc,
            f"a) Probabilidade: {prob}.  b) Impacto: {imp}.",
            space_after=2,
            font_size=9.5,
        )
        add_body_p(doc, dano, bold_prefix="c) Dano: ", space_after=2, font_size=9.5)
        add_body_p(doc, prev, bold_prefix="d) Ação preventiva: ", space_after=2, font_size=9.5)
        add_body_p(doc, cont, bold_prefix="e) Ação de contingência: ", space_after=4, font_size=9.5)

    # 15 DECLARAÇÃO DE VIABILIDADE DA CONTRATAÇÃO
    add_section_heading(doc, "15 DECLARAÇÃO DE VIABILIDADE DA CONTRATAÇÃO")
    add_body_p(
        doc,
        "A solução proposta mostra-se tecnicamente adequada, operacionalmente viável e "
        f"economicamente compatível com os recursos da Emenda Parlamentar nº {META['emenda']} "
        f"(valor conferido de {fmt_brl(TOTAL_CONSUMO)} para os 15 itens de consumo, dentro da "
        f"parcela de {fmt_brl(META['teto_emenda_materiais'])} destinada a materiais e do teto "
        f"global de {fmt_brl(META['teto_emenda_global'])}). Há fornecedores qualificados no "
        "mercado, os quantitativos estão fundamentados no planejamento do Programa PICSPel e os "
        "riscos identificados são mitigáveis. Declara-se, portanto, a viabilidade técnica, "
        "operacional, econômica e administrativa da contratação, recomendando-se o prosseguimento "
        "com a aprovação do Termo de Referência, a consolidação da pesquisa formal de preços "
        "(art. 23), a reserva orçamentária prévia e a instauração de Pregão Eletrônico com "
        "julgamento por item.",
    )

    # 16 EQUIPE DE PLANEJAMENTO
    add_section_heading(doc, "16 EQUIPE DE PLANEJAMENTO")
    add_signatures_block(
        doc,
        include_team_intro=(
            "Equipe de planejamento – Programa PICSPel, Diretoria de Atenção Primária e área "
            "administrativa da SMS, com apoio do Almoxarifado para recebimento e distribuição."
        ),
        autorizo_suffix=True,
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path
