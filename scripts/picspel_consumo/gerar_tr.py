"""Gerador do Termo de Referência (TR) — PICSPel Materiais de Consumo (14 seções padrão RP05_TR)."""

from pathlib import Path

from docx.enum.text import WD_ALIGN_PARAGRAPH

from .dados import (
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
    create_institutional_doc,
)


def gerar_tr(output_path: Path) -> Path:
    doc = create_institutional_doc(
        "TERMO DE REFERÊNCIA – PICSPel (MATERIAIS DE CONSUMO)",
        "Processo Administrativo / Memorando nº [______]/2026  |  DFD nº [___]/2026  |  Emenda Parlamentar nº 43530013",
    )

    # 1 DO OBJETO, VALOR E CONDIÇÕES GERAIS DA CONTRATAÇÃO
    add_section_heading(doc, "1 DO OBJETO, VALOR E CONDIÇÕES GERAIS DA CONTRATAÇÃO")
    add_body_p(
        doc,
        "O presente Termo de Referência tem por objeto a aquisição pontual de materiais de "
        "consumo e insumos destinados às Práticas Integrativas e Complementares em Saúde (Programa "
        "PICSPel), compreendendo 15 (quinze) itens voltados às práticas de escalda-pés, "
        "auriculoterapia, aromaterapia, práticas corporais/meditação e apoio educativo/logístico, "
        "destinados ao abastecimento das 08 (oito) equipes eMulti, das Unidades Básicas de Saúde "
        "(UBS) de referência da Secretaria Municipal de Saúde de Pelotas/RS e ao suporte dos "
        f"cursos de formação previstos com recursos da Emenda Parlamentar nº {META['emenda']}. "
        "Os quantitativos, valores referenciais e condições encontram-se estabelecidos neste "
        "instrumento.",
        bold_prefix="1.1 ",
    )

    tr_rows = [
        (
            it.codigo,
            it.especificacao,
            it.catmat,
            it.unidade,
            fmt_int(it.quantidade),
            fmt_brl(it.valor_unitario),
            fmt_brl(it.valor_total),
        )
        for it in ITENS_CONSUMO
    ]
    add_styled_table(
        doc,
        headers=("Item", "Descrição", "CATMAT", "Unid.", "Qtd.", "Valor unit. (R$)", "Valor total (R$)"),
        rows=tr_rows,
        col_widths_cm=(0.9, 8.2, 1.9, 1.1, 1.1, 1.9, 1.9),
        alignments=(
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.LEFT,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.CENTER,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
            WD_ALIGN_PARAGRAPH.RIGHT,
        ),
        font_size=7.7,
        total_row=(
            "VALOR TOTAL ESTIMADO (15 itens)",
            "",
            "—",
            "—",
            fmt_int(TOTAL_UNIDADES),
            "—",
            fmt_brl(TOTAL_CONSUMO),
        ),
        merge_total_cols=2,
    )

    add_body_p(
        doc,
        "Os objetos deste instrumento não se enquadram como bens de luxo, conforme dispõe o "
        "art. 20 da Lei Federal nº 14.133/2021 e a regulamentação vigente.",
        bold_prefix="1.2 ",
    )
    add_body_p(
        doc,
        "Os objetos são caracterizados como bens comuns de consumo (art. 6º, XIII, da Lei nº "
        "14.133/2021), pois apresentam padrões de desempenho e qualidade objetivamente definidos "
        "por meio de especificações usuais de mercado.",
        bold_prefix="1.3 ",
    )
    add_body_p(
        doc,
        "Por se tratar de aquisição pontual com entrega integral por item (em até 30 dias "
        "corridos) e sem obrigações futuras complexas, o instrumento contratual poderá ser "
        "substituído pela Nota de Empenho acompanhada da respectiva Ordem de Fornecimento, nos "
        "termos do art. 95, II, da Lei Federal nº 14.133/2021, vinculando-se integralmente ao "
        "Edital, a este Termo de Referência e à proposta vencedora. O prazo de vigência da "
        "obrigação decorrente do empenho abrangerá o prazo de entrega, recebimento, eventuais "
        "substituições e pagamento.",
        bold_prefix="1.4 ",
    )
    add_body_p(
        doc,
        "A aquisição será pontual (sem adoção do Sistema de Registro de Preços), com entrega "
        "integral de cada item adjudicado mediante emissão de Nota de Empenho / Ordem de "
        "Fornecimento. Excepcionalmente, havendo conveniência de capacidade de armazenamento no "
        "Almoxarifado da SMS, a entrega poderá ser autorizada em até 2 (duas) remessas, conforme "
        "cronograma previamente formalizado na Ordem de Fornecimento.",
        bold_prefix="1.5 ",
    )
    add_body_p(
        doc,
        "O parcelamento do objeto por itens é tecnicamente viável e economicamente vantajoso "
        "(art. 40, § 2º, da Lei nº 14.133/2021), aplicando-se participação exclusiva para "
        "Microempresas (ME) e Empresas de Pequeno Porte (EPP) em todos os 15 itens, nos termos "
        "do art. 48, I, da Lei Complementar nº 123/2006 c/c art. 4º da Lei nº 14.133/2021.",
        bold_prefix="1.6 ",
    )
    add_body_p(
        doc,
        f"O valor total estimado é de {fmt_brl(TOTAL_CONSUMO)} (quarenta e cinco mil, quinhentos "
        "e vinte reais), conforme soma aritmética exata da tabela do subitem 1.1, a ser validado "
        "pelo Mapa Comparativo de Preços anexo (art. 23 da Lei nº 14.133/2021).",
        bold_prefix="1.7 ",
    )
    add_body_p(
        doc,
        "Em razão de eventuais limitações na precisão das descrições do catálogo CATMAT, poderão "
        "ocorrer divergências em relação às especificações deste Termo de Referência. Nesses "
        "casos, prevalecerão integralmente as especificações deste Termo.",
        bold_prefix="1.8 ",
    )

    # 2 DA FUNDAMENTAÇÃO DA CONTRATAÇÃO
    add_section_heading(doc, "2 DA FUNDAMENTAÇÃO DA CONTRATAÇÃO")
    add_body_p(
        doc,
        "A presente contratação visa atender às necessidades da Secretaria Municipal de Saúde de "
        "Pelotas/RS — Diretoria de Atenção Primária / Programa PICSPel — quanto ao fornecimento "
        "de insumos de consumo para auriculoterapia, aromaterapia, escalda-pés, práticas "
        "corporais/meditação e apoio educativo nas 08 (oito) equipes eMulti (Centro/Porto, "
        "Fragata 1 e 2, Três Vendas 1 e 2, Areal, Colônia e Praias), nas UBS de referência e nos "
        "cursos de formação (30 a 50 vagas), em conformidade com a Política Nacional de Práticas "
        f"Integrativas e Complementares (PNPIC) e a Emenda Parlamentar nº {META['emenda']}. A "
        "fundamentação técnica e a memória de quantitativos constam do Estudo Técnico Preliminar "
        "(ETP). As parcelas de Materiais Permanentes (4.4.90.52) e Serviços de Capacitação "
        "(3.3.90.39) tramitam em processos próprios e autônomos.",
        bold_prefix="2.1 ",
    )
    add_body_p(
        doc,
        "Justificativa para a aquisição pontual (sem Sistema de Registro de Preços): diferentemente "
        "de itens de demanda contínua indeterminada, a presente demanda possui quantitativos "
        "certos, previamente dimensionados para a implantação e fortalecimento do PICSPel e "
        "vinculados à execução de recurso específico de emenda parlamentar com objeto e "
        "cronograma definidos, razão pela qual se adota o Pregão Eletrônico tradicional com "
        "reserva orçamentária prévia.",
        bold_prefix="2.2 ",
    )
    add_body_p(
        doc,
        "A contratação está alinhada ao planejamento estratégico da Secretaria Municipal de Saúde "
        "e ao Plano de Contratações Anual da Diretoria de Atenção Primária: [A INFORMAR: nº do "
        "item no PCA 2026 ou justificativa formal de inclusão excepcional vinculada à Emenda "
        "Parlamentar nº 43530013].",
        bold_prefix="2.3 ",
    )

    # 3 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO CONSIDERADO O CICLO DE VIDA DO OBJETO
    add_section_heading(doc, "3 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO CONSIDERADO O CICLO DE VIDA DO OBJETO")
    add_body_p(
        doc,
        "A solução consiste no fornecimento pontual dos 15 itens de consumo relacionados no "
        "subitem 1.1, com entrega no Almoxarifado da Secretaria Municipal de Saúde de Pelotas/RS, "
        "recebimento provisório e definitivo mediante conferência física, documental, sanitária "
        "e de prazo de validade, armazenamento sob condições adequadas (local seco, ventilado e "
        "protegido de luz direta e calor excessivo para óleos essenciais e carreador), "
        "distribuição planejada às 08 equipes eMulti e UBS de referência, utilização assistencial "
        "e educativa (com higienização e autoclavagem dos instrumentais de aço inoxidável) e "
        "descarte ambientalmente adequado de embalagens e resíduos em observância à Política "
        "Nacional de Resíduos Sólidos (Lei nº 12.305/2010), à coleta seletiva municipal e ao "
        "Plano de Gerenciamento de Resíduos de Serviços de Saúde (PGRSS — RDC ANVISA nº 222/2018). "
        "A fundamentação detalhada encontra-se pormenorizada no Estudo Técnico Preliminar.",
    )

    # 4 GESTÃO/ANÁLISE DE RISCOS
    add_section_heading(doc, "4 GESTÃO/ANÁLISE DE RISCOS")
    add_body_p(
        doc,
        "Os principais riscos mapeados no Estudo Técnico Preliminar são: (I) atraso na entrega "
        "dos insumos (probabilidade média; impacto médio); (II) entrega em desacordo com as "
        "especificações técnicas (probabilidade média; impacto alto); (III) itens desertos, "
        "fracassados ou com variação excessiva de preços (probabilidade média; impacto médio); "
        "(IV) armazenamento ou controle de estoque inadequado pós-entrega (probabilidade média; "
        "impacto médio); e (V) entrega de insumos com validade exígua, irregularidade sanitária "
        "na ANVISA ou óleos essenciais com pureza comprometida (probabilidade média; impacto alto).",
    )
    add_body_p(
        doc,
        "Como medidas de prevenção e mitigação, serão observados o julgamento por item, a "
        "definição precisa das especificações sem direcionamento de marca, a exigência "
        "proporcional de regularização ANVISA e de laudos/cromatografia gasosa para os óleos "
        "essenciais, o controle rigoroso de validade no recebimento, a faculdade de amostras, a "
        "reunião pós-entrega com as equipes e a fiscalização ativa. Em caso de descumprimento, "
        "serão aplicadas as sanções previstas na Lei Federal nº 14.133/2021 e no Decreto "
        "Municipal nº 6.799/2023.",
    )

    # 5 REQUISITOS DA CONTRATAÇÃO
    add_section_heading(doc, "5 REQUISITOS DA CONTRATAÇÃO")
    add_body_p(
        doc,
        "Não serão exigidas marcas ou modelos comerciais específicos (art. 41, I, da Lei nº "
        "14.133/2021), sendo admitidos produtos de diferentes fabricantes desde que atendam "
        "integralmente às especificações técnicas deste Termo, preservando-se a isonomia e a "
        "competitividade. Eventuais designações técnicas consolidadas na área de PICS (como "
        "\"tipo Adson\" e \"padrão Nogier\") referem-se exclusivamente à geometria/funcionalidade "
        "do instrumental, sem vinculação a marca comercial.",
        bold_prefix="5.1 ",
    )
    add_body_p(doc, "Requisitos técnicos específicos por família de produtos:", bold_prefix="5.2 ")
    add_bullet(
        doc,
        "bacias anatômicas (Item 01) em polipropileno ou plástico rígido resistente à água morna "
        "e desinfecção; protetores descartáveis de uso único em polietileno (Item 02) compatíveis "
        "com o revestimento da bacia; e sal grosso marinho iodado 1 kg (Item 15), 100% natural, "
        "sem aditivos ou fragrâncias, em embalagem íntegra rotulada conforme RDC ANVISA nº 604/2022.",
        bold_prefix="Escalda-pés (Itens 01, 02 e 15): ",
    )
    add_bullet(
        doc,
        "placas com 70 pontos de esfera de cristal (~1,5–1,6 mm) em fita adesiva microporosa "
        "hipoalergênica quadrada (Item 03, caixa com 30 placas); pinças anatômicas tipo Adson "
        "12 cm retas sem dente (Item 04) e apalpadores de pressão com mola padrão Nogier "
        "(Item 13), ambos em aço inoxidável de grau cirúrgico/clínico, autoclaváveis.",
        bold_prefix="Auriculoterapia (Itens 03, 04 e 13): ",
    )
    add_bullet(
        doc,
        "óleos essenciais 100% puros de Lavanda (Lavandula angustifolia), Hortelã-pimenta "
        "(Mentha piperita), Melaleuca (Melaleuca alternifolia) e Laranja Doce (Citrus sinensis) "
        "em frascos de vidro âmbar de 10 ml com gotejador e lacre de inviolabilidade (Itens 05 a "
        "08); óleo vegetal carreador 100% puro de semente de uva (Vitis vinifera) prensado a frio "
        "em frasco âmbar de 250 ml com lacre (Item 09); frascos âmbar 5 ml vazios com batoque "
        "gotejador e tampa vedante (Item 10); e inaladores nasais portáteis vazios com pavio de "
        "algodão (Item 11).",
        bold_prefix="Aromaterapia (Itens 05 a 11): ",
    )
    add_bullet(
        doc,
        "tapetes de yoga/práticas corporais em EVA atóxico antiderrapante com espessura mínima "
        "de 5 mm (Item 12) e filmes duplos para plastificação térmica padrão A4, 0,07 mm / 175 "
        "micras, pacote com 100 folhas (Item 14).",
        bold_prefix="Práticas corporais e apoio educativo/logístico (Itens 12 e 14): ",
    )

    add_body_p(
        doc,
        "Os produtos deverão ser novos, de primeiro uso, acondicionados em embalagem original "
        "íntegra, limpa e lacrada, com rotulagem em língua portuguesa contendo identificação do "
        "fabricante/importador, composição/espécie botânica (quando aplicável), volume ou "
        "quantidade, número de lote, data de fabricação, prazo de validade e cuidados de "
        "armazenamento/uso. Frascos de vidro (Itens 05 a 10) deverão ser embalados com proteção "
        "adequada contra choque mecânico durante o transporte.",
        bold_prefix="5.3 ",
    )
    add_body_p(
        doc,
        "Prazo de validade mínimo na entrega: (a) para os Itens 03 (placas de auriculoterapia), "
        "05 a 08 (óleos essenciais 10 ml), 09 (óleo vegetal carreador 250 ml) e 15 (sal grosso "
        "marinho iodado 1 kg), os produtos deverão apresentar, no ato da entrega no Almoxarifado "
        "da SMS, validade mínima de 12 (doze) meses a partir da data da entrega, correspondendo "
        "ainda a, no mínimo, 75% (setenta e cinco por cento) do prazo total de validade estipulado "
        "pelo fabricante (recomendando-se validade igual ou superior a 18 meses para os Itens 03 "
        "e 05 a 07); (b) para os demais itens não perecíveis (Itens 01, 02, 04, 10, 11, 12, 13 e "
        "14), os produtos deverão ser novos, de fabricação recente e livres de ressecamento, "
        "oxidação ou avarias — e, caso o fabricante estipule validade, esta deverá ser de no "
        "mínimo 12 (doze) meses na entrega. Produtos entregues com validade inferior serão "
        "recusados e substituídos pela contratada sem ônus para a Administração.",
        bold_prefix="5.4 ",
    )
    add_body_p(
        doc,
        "Regularização sanitária (ANVISA) e documentação técnica por categoria: a licitante "
        "provisoriamente vencedora deverá apresentar, juntamente com a proposta ajustada ou "
        "quando solicitado pelo Pregoeiro, catálogo ou ficha técnica do fabricante comprovando o "
        "atendimento integral às especificações e, exclusivamente para os itens regulados:",
        bold_prefix="5.5 ",
    )
    add_bullet(
        doc,
        "comprovante de notificação, cadastro ou registro vigente do produto na ANVISA "
        "(RDC nº 751/2022 ou norma substitutiva), ou documento oficial de enquadramento/isenção.",
        bold_prefix="Itens 03, 04 e 13 (dispositivos/instrumentais de auriculoterapia): ",
    )
    add_bullet(
        doc,
        "comprovante de regularização sanitária vigente na ANVISA (notificação/comunicação "
        "prévia, cadastro ou registro conforme enquadramento regulatório do produto) e, para os "
        "óleos essenciais (Itens 05 a 08), laudo/certificado de análise ou cromatografia gasosa "
        "(CG) do lote (ou ficha técnica/declaração técnica do fabricante) atestando 100% de pureza "
        "e correspondência à espécie botânica especificada.",
        bold_prefix="Itens 05 a 08 (óleos essenciais) e Item 09 (óleo vegetal carreador): ",
    )
    add_bullet(
        doc,
        "rotulagem em conformidade com os padrões sanitários vigentes para sal marinho iodado "
        "(RDC ANVISA nº 604/2022 e normas correlatas).",
        bold_prefix="Item 15 (sal grosso marinho iodado): ",
    )
    add_bullet(
        doc,
        "dispensados de registro ou notificação sanitária na ANVISA, sendo exigida apenas a "
        "ficha técnica/catálogo do fabricante para conferência das especificações.",
        bold_prefix="Itens 01, 02, 10, 11, 12 e 14 (itens não sujeitos a registro sanitário): ",
    )

    add_body_p(
        doc,
        "Amostras: a Administração poderá exigir, do licitante provisoriamente classificado em "
        "primeiro lugar, amostra de item específico, no prazo de 5 (cinco) dias úteis contados da "
        "convocação, para análise técnica de conformidade, nos termos do art. 41, II, da Lei nº "
        "14.133/2021 (verificação de integridade, aderência das placas, acabamento/mola dos "
        "instrumentais, vedação e características organolépticas/rotulagem dos óleos, resistência "
        "das bacias e tapetes). A reprovação fundamentada resultará na desclassificação da proposta "
        "para o respectivo item.",
        bold_prefix="5.6 ",
    )
    add_body_p(
        doc,
        "Para fins de habilitação, a licitante vencedora deverá apresentar, nos termos dos "
        "arts. 62 a 70 da Lei nº 14.133/2021:",
        bold_prefix="5.7 ",
    )
    for hab in (
        "Comprovante de inscrição e de situação cadastral ativa no CNPJ, com ramo de atividade compatível com o(s) item(ns) disputado(s);",
        "Contrato social, estatuto ou documento equivalente vigente, registrado no órgão competente, com as respectivas alterações;",
        "Certidões negativas ou positivas com efeito de negativas de regularidade fiscal e trabalhista: Fazenda Federal (inclusive contribuições previdenciárias), Fazenda Estadual, Fazenda Municipal, FGTS e CNDT;",
        "Declaração de que não emprega menor em condições vedadas, nos termos do inciso XXXIII do art. 7º da Constituição Federal, e declaração de enquadramento como ME/EPP (art. 4º da Lei nº 14.133/2021);",
        "Autorização de Funcionamento de Empresa (AFE) da ANVISA e Licença/Alvará Sanitário vigente, exigíveis EXCLUSIVAMENTE para os licitantes vencedores dos itens sujeitos à vigilância sanitária (Itens 03, 04, 05 a 09 e 13), quando legalmente obrigatórios à atividade/porte do fornecedor — vedada tal exigência para os Itens 01, 02, 10, 11, 12, 14 e 15;",
        "No ato da entrega, laudo/certificado de análise ou cromatografia gasosa (CG) dos lotes entregues para os Itens 05 a 08 e documentação fiscal completa com lote e validade.",
    ):
        add_bullet(doc, hab)

    add_body_p(
        doc,
        "É vedada a subcontratação total ou parcial do fornecimento do objeto, não se confundindo "
        "com a contratação de serviços de frete/transporte rodoviário por conta e risco da "
        "contratada.",
        bold_prefix="5.8 ",
    )
    add_body_p(
        doc,
        "O fornecimento deverá contemplar todos os custos diretos e indiretos (modalidade CIF), "
        "incluindo transporte, seguro, embalagens de proteção contra avarias/quebra de vidros, "
        "carga e descarga no local de entrega, sob inteira responsabilidade da contratada.",
        bold_prefix="5.9 ",
    )
    add_body_p(
        doc,
        "Os produtos somente serão aceitos após conferência técnica efetuada pela Secretaria "
        "Municipal de Saúde e, caso não satisfaçam às especificações, apresentem avarias, lacres "
        "violados ou validade inferior à exigida, deverão ser retirados e substituídos pela "
        "contratada no prazo máximo de 5 (cinco) dias úteis, contados da notificação, sem ônus "
        "adicional para a Administração.",
        bold_prefix="5.10 ",
    )

    # 6 DA EXECUÇÃO DO OBJETO
    add_section_heading(doc, "6 DA EXECUÇÃO DO OBJETO")
    add_body_p(
        doc,
        f"Os produtos deverão ser entregues no {META['almoxarifado_endereco']}, "
        f"{META['almoxarifado_horario']}, mediante prévio agendamento pelo e-mail "
        f"{META['almoxarifado_email']} (ou em outro endereço urbano em Pelotas/RS formalmente "
        "indicado na Ordem de Fornecimento pela Diretoria de Atenção Primária / coordenação "
        "PICSPel).",
        bold_prefix="6.1 ",
    )
    add_body_p(
        doc,
        "O prazo de entrega dos materiais será de até 30 (trinta) dias corridos, contados do "
        "recebimento da Nota de Empenho / Ordem de Fornecimento pela contratada.",
        bold_prefix="6.2 ",
    )
    add_body_p(
        doc,
        "Caso não seja possível a entrega na data assinalada, a contratada deverá comunicar "
        "formalmente as razões, acompanhadas de comprovação documental, com pelo menos 5 (cinco) "
        "dias corridos de antecedência ao término do prazo, para análise do pedido de prorrogação "
        "pela fiscalização, ressalvadas as hipóteses de caso fortuito e força maior.",
        bold_prefix="6.3 ",
    )
    add_body_p(
        doc,
        "Os produtos entregues deverão respeitar rigorosamente o prazo de validade estipulado no "
        "subitem 5.4, bem como ser transportados sob proteção contra luz direta, calor excessivo, "
        "umidade e impacto mecânico (especialmente os frascos de vidro âmbar dos Itens 05 a 10).",
        bold_prefix="6.4 ",
    )
    add_body_p(
        doc,
        "Após o recebimento definitivo no Almoxarifado, a coordenação do PICSPel promoverá "
        "reunião interna com as 08 (oito) equipes eMulti e equipes da APS de referência para "
        "organização do armazenamento, controle de estoque/validades e distribuição dos insumos.",
        bold_prefix="6.5 ",
    )

    # 7 DA GESTÃO E FISCALIZAÇÃO DA CONTRATAÇÃO
    add_section_heading(doc, "7 DA GESTÃO E FISCALIZAÇÃO DA CONTRATAÇÃO")
    add_body_p(
        doc,
        "A contratação (formalizada mediante Nota de Empenho / Ordem de Fornecimento) deverá ser "
        "executada fielmente pelas partes, de acordo com as cláusulas avençadas no Edital, neste "
        "Termo de Referência e nas normas da Lei Federal nº 14.133/2021, respondendo cada parte "
        "pelas consequências de sua inexecução total ou parcial.",
        bold_prefix="7.1 ",
    )
    add_body_p(
        doc,
        "A gestão e a fiscalização do fornecimento serão exercidas pelos servidores formalmente "
        "designados pela Secretaria Municipal de Saúde (unificados aos demais processos do "
        f"Programa PICSPel): Gestor: {META['gestor_nome']} (Matrícula {META['gestor_matricula']}); "
        f"Fiscal: {META['fiscal_nome']} (Matrícula {META['fiscal_matricula']}), com apoio técnico-"
        "operacional do Almoxarifado da SMS no recebimento físico e conferência de lotes.",
        bold_prefix="7.2 ",
    )
    add_body_p(
        doc,
        "As comunicações entre a Administração e a contratada deverão ser realizadas por escrito "
        "sempre que o ato exigir tal formalidade, admitindo-se o uso de mensagem eletrônica "
        "institucional para esse fim.",
        bold_prefix="7.3 ",
    )
    add_body_p(
        doc,
        "O descumprimento total ou parcial das obrigações assumidas caracterizará inadimplência, "
        "sujeitando a contratada às infrações e sanções administrativas previstas nos arts. 155 a "
        "163 da Lei Federal nº 14.133/2021 e no Decreto Municipal nº 6.799/2023 (advertência, "
        "multa, impedimento de licitar e contratar e declaração de inidoneidade), assegurados o "
        "contraditório e a ampla defesa, sem prejuízo das responsabilidades civil e criminal.",
        bold_prefix="7.4 ",
    )
    add_body_p(
        doc,
        "Fica dispensada a exigência de garantia contratual (art. 96 da Lei nº 14.133/2021), "
        "tendo em vista tratar-se de compra pontual com entrega integral por item e pagamento "
        "condicionado ao recebimento definitivo, inexistindo risco financeiro de adiantamento.",
        bold_prefix="7.5 ",
    )

    # 8 RECEBIMENTO DO OBJETO
    add_section_heading(doc, "8 RECEBIMENTO DO OBJETO")
    add_body_p(
        doc,
        "Os bens serão recebidos provisoriamente, de forma sumária, no ato da entrega e no prazo "
        "de até 7 (sete) dias corridos, pelo responsável pelo acompanhamento e fiscalização (com "
        "apoio do Almoxarifado da SMS), nos termos do art. 140, II, \"a\", da Lei nº 14.133/2021, "
        "para verificação preliminar de volumes, integridade externa das embalagens e Nota Fiscal.",
        bold_prefix="8.1 ",
    )
    add_body_p(
        doc,
        "Os bens poderão ser rejeitados, no todo ou em parte, quando em desacordo com as "
        "especificações constantes deste Termo de Referência e da proposta, quando apresentarem "
        "avarias, vazamentos, lacres rompidos, prazo de validade inferior ao exigido no subitem "
        "5.4 ou ausência de documentação sanitária/laudos obrigatórios, devendo ser retirados e "
        "substituídos no prazo máximo de 5 (cinco) dias úteis, a contar da notificação formal, às "
        "expensas da contratada, sem prejuízo da aplicação das penalidades cabíveis.",
        bold_prefix="8.2 ",
    )
    add_body_p(
        doc,
        "Os bens serão recebidos definitivamente no prazo de até 8 (oito) dias corridos, contados "
        "do recebimento provisório, nos termos do art. 140, II, \"b\", da Lei nº 14.133/2021, "
        "após a verificação qualitativa e quantitativa detalhada (incluindo conferência de "
        "especificações, lotes, validades, regularização ANVISA dos itens regulados e laudos/CG "
        "dos óleos essenciais) e a aceitação formal mediante termo detalhado / ateste na Nota "
        "Fiscal pela fiscalização.",
        bold_prefix="8.3 ",
    )

    # 9 DOS CRITÉRIOS DE MEDIÇÃO E PAGAMENTO
    add_section_heading(doc, "9 DOS CRITÉRIOS DE MEDIÇÃO E PAGAMENTO")
    add_body_p(
        doc,
        "O pagamento será efetuado por item efetivamente entregue, após o recebimento definitivo "
        "e a conferência/ateste da Nota Fiscal/DANFE pela fiscalização competente da Secretaria "
        "Municipal de Saúde, mediante crédito em conta bancária de titularidade da contratada, "
        "observada a ordem cronológica do art. 141 da Lei nº 14.133/2021 e o cronograma fixado "
        "pela Secretaria Municipal da Fazenda, sendo vedado o pagamento antecipado (art. 145).",
    )
    add_body_p(
        doc,
        "A Nota Fiscal/DANFE deverá discriminar cada item entregue, quantidade, valor unitário e "
        "total, marca/fabricante, número de lote e data de validade (para os itens aplicáveis), "
        "número da Nota de Empenho e fazer referência expressa à \"Emenda Parlamentar nº "
        f"{META['emenda']} – Programa PICSPel\". A liquidação dependerá da manutenção das "
        "condições de regularidade fiscal e trabalhista da contratada, procedendo-se às retenções "
        "tributárias previstas na legislação vigente (IRRF e demais encargos aplicáveis).",
    )

    # 10 DA FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR
    add_section_heading(doc, "10 DA FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR")
    add_body_p(
        doc,
        "O fornecedor será selecionado por meio de licitação na modalidade Pregão, em sua forma "
        "eletrônica (aquisição pontual — sem Sistema de Registro de Preços), nos termos da Lei "
        "Federal nº 14.133/2021 (art. 6º, XLI, e art. 28, I), por possuir o objeto especificações "
        "usuais de mercado e padrões de desempenho e qualidade objetivamente definidos.",
        bold_prefix="10.1 ",
    )
    add_body_p(
        doc,
        "Critério de julgamento: menor preço por item, com adjudicação por item.",
        bold_prefix="10.2 ",
    )
    add_body_p(
        doc,
        "Tratamento diferenciado: participação exclusiva de Microempresas (ME) e Empresas de "
        "Pequeno Porte (EPP) em todos os 15 itens, nos termos do art. 48, inciso I, da Lei "
        "Complementar nº 123/2006 c/c art. 4º da Lei Federal nº 14.133/2021, observadas as "
        "ressalvas legais do art. 49 da LC nº 123/2006.",
        bold_prefix="10.3 ",
    )
    add_body_p(
        doc,
        "Estimativa de preços: a pesquisa formal de preços será consolidada nos termos do art. 23 "
        "da Lei nº 14.133/2021, conforme Mapa Comparativo de Preços a ser anexado ao processo "
        "[A INFORMAR: fontes consultadas — PNCP/Painel de Preços, contratações similares e/ou "
        "cotações de fornecedores — e metodologia estatística adotada], validando os valores "
        f"unitários referenciais e o total estimado de {fmt_brl(TOTAL_CONSUMO)}.",
        bold_prefix="10.4 ",
    )

    # 11 DAS OBRIGAÇÕES DO CONTRATANTE
    add_section_heading(doc, "11 DAS OBRIGAÇÕES DO CONTRATANTE")
    for ob in (
        "Exigir o integral cumprimento, pela contratada, de todas as obrigações assumidas de acordo com este Termo de Referência, o Edital e a proposta vencedora;",
        "Receber o objeto no prazo e nas condições estabelecidas, realizando o agendamento de descarga no Almoxarifado da SMS e a conferência física, documental e sanitária;",
        "Prestar à contratada as informações e esclarecimentos necessários à perfeita execução do objeto;",
        "Notificar a contratada, por escrito, sobre imperfeições, avarias, divergências de validade, falta de documentação sanitária/laudos ou irregularidades verificadas, fixando prazo para correção ou substituição;",
        "Exercer o acompanhamento e a fiscalização por servidores formalmente designados (Gestor e Fiscal), anotando em registro próprio todas as ocorrências detectadas;",
        "Efetuar o pagamento no prazo, forma e condições estabelecidos após o recebimento definitivo e o ateste da Nota Fiscal;",
        "Aplicar à contratada as sanções administrativas previstas na Lei nº 14.133/2021 e no Decreto Municipal nº 6.799/2023 em caso de inadimplemento;",
        "Não praticar atos de ingerência na administração da contratada;",
        "Emitir decisão fundamentada sobre solicitações, pedidos de prorrogação e reclamações relacionadas à execução do objeto;",
        "Fiscalizar a manutenção, pela contratada, das condições de habilitação e regularidade fiscal durante toda a execução.",
    ):
        add_bullet(doc, ob)

    # 12 DAS OBRIGAÇÕES DO CONTRATADO
    add_section_heading(doc, "12 DAS OBRIGAÇÕES DO CONTRATADO")
    for ob in (
        "Cumprir integralmente as obrigações assumidas, assumindo todos os riscos e despesas decorrentes da boa e perfeita execução do fornecimento;",
        "Entregar os materiais novos, de primeiro uso, exatamente conforme as especificações técnicas do subitem 1.1, respeitando rigorosamente os prazos de validade mínima (subitem 5.4) e as normas da ANVISA para os itens regulados (subitem 5.5);",
        "Apresentar, nos termos do subitem 5.5, comprovantes de regularização sanitária ANVISA dos itens aplicáveis e laudos/certificados de análise ou cromatografia gasosa (CG) dos lotes dos óleos essenciais (Itens 05 a 08);",
        "Atender prontamente às determinações do fiscal ou da autoridade superior e realizar o agendamento prévio da entrega junto ao Almoxarifado da SMS;",
        "Retirar e substituir, às suas expensas e no prazo máximo de 5 (cinco) dias úteis da notificação, os bens entregues com vícios, defeitos, avarias, vazamentos, lacres rompidos ou validade inferior à exigida;",
        "Responsabilizar-se pelos vícios e danos decorrentes do fornecimento, nos termos do Código de Defesa do Consumidor (Lei nº 8.078/1990);",
        "Não permitir a utilização de trabalho de menor de dezesseis anos, exceto na condição de aprendiz a partir dos quatorze anos, nem de menor de dezoito anos em atividades noturnas, perigosas ou insalubres (art. 7º, XXXIII, CF), bem como cumprir a reserva de cargos prevista em lei quando aplicável (art. 92, XVII, da Lei nº 14.133/2021);",
        "Responsabilizar-se integralmente pelo pagamento de tributos, encargos trabalhistas, previdenciários, fiscais e comerciais incidentes sobre o fornecimento;",
        "Manter, durante toda a execução do objeto, compatibilidade com as obrigações assumidas e todas as condições de habilitação e qualificação exigidas na licitação;",
        "Arcar com todas as despesas de transporte, frete (CIF), seguro, carga, descarga e embalagens de proteção (especialmente contra quebra de frascos de vidro âmbar e variações térmicas dos óleos);",
        f"Emitir Nota Fiscal/DANFE detalhada contendo item, marca/fabricante, lote e validade (quando aplicáveis), número da Nota de Empenho e referência expressa à Emenda Parlamentar nº {META['emenda']} – Programa PICSPel.",
    ):
        add_bullet(doc, ob)

    # 13 DA ADEQUAÇÃO ORÇAMENTÁRIA
    add_section_heading(doc, "13 DA ADEQUAÇÃO ORÇAMENTÁRIA")
    add_body_p(
        doc,
        "Por se tratar de aquisição pontual (sem utilização do Sistema de Registro de Preços), a "
        "indicação definitiva da dotação orçamentária e a prévia emissão da Nota de Reserva "
        "Orçamentária são obrigatórias antes da publicação do edital e da contratação, nos termos "
        "do art. 150 da Lei Federal nº 14.133/2021. As despesas correrão à conta dos recursos da "
        f"Emenda Parlamentar nº {META['emenda']}, conforme quadro abaixo:",
    )

    orc_rows = [
        ("Emenda Parlamentar", f"Emenda nº {META['emenda']} (teto global: {fmt_brl(META['teto_emenda_global'])})"),
        ("Projeto / Atividade", "[A INFORMAR — Atenção Primária / PICSPel]"),
        ("Fonte de Recurso", "[A INFORMAR — vinculada à Emenda Parlamentar nº 43530013]"),
        ("Reduzido / Dotação Orçamentária", "[A INFORMAR — Contabilidade / CCGOF]"),
        ("Nota de Reserva Orçamentária", "[A ANEXAR antes da publicação do edital — art. 150]"),
        ("Classificação da Despesa", META["elemento_desc"]),
        ("Valor Estimado deste TR (Consumo)", f"{fmt_brl(TOTAL_CONSUMO)} (15 itens — 1.508 unidades comerciais)"),
        (
            "Compatibilidade na Emenda nº 43530013",
            f"Consumo ({fmt_brl(TOTAL_CONSUMO)}) + Permanentes ({fmt_brl(META['valor_processo_permanentes'])}) "
            f"= {fmt_brl(SUBTOTAL_MATERIAIS)} (teto materiais: {fmt_brl(META['teto_emenda_materiais'])})  |  "
            f"Total 3 processos: {fmt_brl(TOTAL_CONSOLIDADO_EMENDA)} (saldo residual: {fmt_brl(SALDO_RESIDUAL_EMENDA)})",
        ),
    ]
    add_styled_table(
        doc,
        headers=("Campo Orçamentário", "Especificação"),
        rows=orc_rows,
        col_widths_cm=(5.8, 11.2),
        alignments=(WD_ALIGN_PARAGRAPH.LEFT, WD_ALIGN_PARAGRAPH.LEFT),
        font_size=8.6,
    )
    add_callout(
        doc,
        "Subelementos 3.3.90.30",
        "Caberá à Contabilidade/CCGOF confirmar os subelementos específicos de 3.3.90.30 "
        "aplicáveis aos 15 itens de consumo (material ambulatorial/hospitalar, material de "
        "acondicionamento/embalagem, material esportivo/educativo, expediente e insumos de "
        "escalda-pés/aromaterapia) previamente à emissão das Notas de Empenho.",
    )

    # 14 EQUIPE DE PLANEJAMENTO
    add_section_heading(doc, "14 EQUIPE DE PLANEJAMENTO")
    add_signatures_block(
        doc,
        include_team_intro=(
            "Equipe de planejamento – Departamento/Serviço demandante: Programa PICSPel, "
            "Diretoria de Atenção Primária e área administrativa da SMS, com apoio do "
            "Almoxarifado da Secretaria Municipal de Saúde."
        ),
        autorizo_suffix=False,
    )

    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    return output_path
