"""Dados estruturados e fonte única da verdade — PICSPel Materiais de Consumo (Processo 3)."""

from dataclasses import dataclass
from decimal import Decimal


@dataclass(frozen=True)
class ItemConsumo:
    codigo: str
    nome_curto: str
    especificacao: str
    catmat: str
    unidade: str
    quantidade: int
    valor_unitario: Decimal
    destinacao: str
    familia: str
    anvisa: str
    validade_minima: str

    @property
    def valor_total(self) -> Decimal:
        return self.valor_unitario * self.quantidade


def fmt_brl(valor: Decimal) -> str:
    """Formata Decimal como moeda brasileira (ex.: R$ 45.520,00)."""
    quant = valor.quantize(Decimal("0.01"))
    inteiro, dec = f"{quant:.2f}".split(".")
    inteiro_fmt = f"{int(inteiro):,}".replace(",", ".")
    return f"R$ {inteiro_fmt},{dec}"


def fmt_int(valor: int) -> str:
    return f"{valor:,}".replace(",", ".")


META = {
    "orgao_linha1": "ESTADO DO RIO GRANDE DO SUL",
    "orgao_linha2": "MUNICÍPIO DE PELOTAS",
    "orgao_linha3": "SECRETARIA MUNICIPAL DE SAÚDE",
    "unidade_demandante": (
        "Secretaria Municipal de Saúde – SMS / Diretoria de Atenção Primária / "
        "Programa Municipal de Práticas Integrativas e Complementares em Saúde (PICSPel)"
    ),
    "emenda": "43530013",
    "elemento": "3.3.90.30",
    "elemento_desc": "3.3.90.30 – Material de Consumo (subelementos a validar pela Contabilidade/CCGOF)",
    "gestor_nome": "Dariane Lima Portela",
    "gestor_matricula": "44151",
    "fiscal_nome": "Aline Kohler Geppert",
    "fiscal_matricula": "29371",
    "secretaria_nome": "Angela Moreira Vitória",
    "secretaria_cargo": "Secretária Municipal de Saúde",
    "almoxarifado_endereco": (
        "Almoxarifado da Secretaria Municipal de Saúde de Pelotas/RS, localizado na "
        "Av. Fernando Osório, nº 273, Bairro Três Vendas, Pelotas/RS, CEP 96020-150"
    ),
    "almoxarifado_horario": (
        "de segunda a sexta-feira, das 8h às 11h30min e das 13h30min às 16h, em dias úteis"
    ),
    "almoxarifado_email": "smspel.almoxarifado@gmail.com",
    "equipes_emulti": [
        "Centro/Porto",
        "Fragata 1",
        "Fragata 2",
        "Três Vendas 1",
        "Três Vendas 2",
        "Areal",
        "Colônia",
        "Praias",
    ],
    "valor_nominal_original_pdf": Decimal("47200.00"),
    "teto_emenda_global": Decimal("100000.00"),
    "teto_emenda_materiais": Decimal("60000.00"),
    "teto_emenda_servicos": Decimal("40000.00"),
    "valor_processo_servicos": Decimal("40000.00"),
    "valor_processo_permanentes": Decimal("11600.00"),
}


ITENS_CONSUMO: list[ItemConsumo] = [
    ItemConsumo(
        codigo="01",
        nome_curto="Bacia para escalda-pés",
        especificacao=(
            "Bacia para escalda-pés: recipiente anatômico para imersão simultânea dos pés adultos, "
            "confeccionado em polipropileno (PP) ou polímero plástico rígido de alta resistência, "
            "atóxico, superfície interna lisa, lavável, sem rebarbas, apto a suportar água morna "
            "(até no mínimo 50 °C) e desinfecção de nível intermediário. Dimensões de referência: "
            "aproximadamente 34 cm × 39 cm × 8 cm (admitida variação de ±15% no comprimento e na "
            "largura e altura mínima de 8 cm, adequada à imersão confortável dos pés). Peso "
            "aproximado: 300 g a 450 g (ou superior, assegurada resistência estrutural). Produto "
            "novo, de primeiro uso."
        ),
        catmat="[A CONFIRMAR]",
        unidade="UN",
        quantidade=120,
        valor_unitario=Decimal("35.00"),
        destinacao="eMulti (15 un. × 8 equipes)",
        familia="Escalda-pés",
        anvisa="Dispensado (utensílio de apoio)",
        validade_minima="Indeterminada (novo, 1º uso)",
    ),
    ItemConsumo(
        codigo="02",
        nome_curto="Protetor descartável para bacias (pacote com 100 un.)",
        especificacao=(
            "Protetor descartável para bacias de escalda-pés: revestimento plástico descartável, "
            "de uso único, confeccionado em polietileno (PE) atóxico, impermeável e resistente à "
            "água morna e à ruptura durante o procedimento, com elástico ou abertura compatível "
            "com ajuste seguro na borda da bacia (Item 01). Tamanho único aproximado de 68 cm × "
            "68 cm (tolerância de ±10%). Apresentação: pacote íntegro e lacrado com 100 (cem) "
            "unidades, identificado em língua portuguesa. Validade mínima de 12 (doze) meses na "
            "entrega, quando estipulada pelo fabricante."
        ),
        catmat="[A CONFIRMAR]",
        unidade="PCT",
        quantidade=80,
        valor_unitario=Decimal("40.00"),
        destinacao="eMulti (10 pct. × 8 equipes)",
        familia="Escalda-pés",
        anvisa="Dispensado (barreira higiênica)",
        validade_minima="Indeterminada ou ≥ 12 meses",
    ),
    ItemConsumo(
        codigo="03",
        nome_curto="Placas para auriculoterapia com cristal (caixa com 30 placas)",
        especificacao=(
            "Placas para auriculoterapia com esferas de cristal: cartela plástica contendo 70 "
            "(setenta) pontos prontos para auriculoterapia, cada ponto composto por esfera de "
            "cristal polido de aproximadamente 1,5 mm a 1,6 mm (referência 1,58 mm) fixada em "
            "fita adesiva microporosa hipoalergênica de formato quadrado (cor bege/pele, "
            "respirável, sem vinculação a marca comercial), de uso único. Apresentação: caixa "
            "contendo 30 (trinta) placas (total de 2.100 pontos por caixa), com rotulagem em "
            "português, lote, data de fabricação, validade e regularização sanitária vigente na "
            "ANVISA (notificação, cadastro ou registro conforme RDC nº 751/2022). Validade "
            "mínima de 12 (doze) meses a partir da data da entrega (com pelo menos 75% do prazo "
            "total de validade vigente; recomendável ≥ 18 meses)."
        ),
        catmat="[A CONFIRMAR]",
        unidade="CX",
        quantidade=50,
        valor_unitario=Decimal("190.00"),
        destinacao="UBS / rede APS",
        familia="Auriculoterapia",
        anvisa="Obrigatória (Dispositivo Médico — RDC nº 751/2022)",
        validade_minima="≥ 12 meses (≥ 75% vida útil)",
    ),
    ItemConsumo(
        codigo="04",
        nome_curto="Pinça Adson 12 cm, reta, sem dente",
        especificacao=(
            "Pinça anatômica tipo Adson, reta, sem dente de rato, ponta fina com serrilha interna "
            "transversal para preensão precisa em procedimentos de auriculoterapia; confeccionada "
            "em aço inoxidável cirúrgico (série AISI 420 ou equivalente de grau cirúrgico), "
            "autoclavável (resistente à esterilização em autoclave a vapor), acabamento livre de "
            "rebarbas ou oxidação. Comprimento aproximado: 12 cm (tolerância de ±0,5 cm). "
            "Produto com regularização sanitária vigente na ANVISA (notificação/cadastro conforme "
            "RDC nº 751/2022) e embalagem individual íntegra identificada em língua portuguesa."
        ),
        catmat="[A CONFIRMAR]",
        unidade="UN",
        quantidade=100,
        valor_unitario=Decimal("35.00"),
        destinacao="UBS / rede APS",
        familia="Auriculoterapia",
        anvisa="Obrigatória (Instrumental — RDC nº 751/2022)",
        validade_minima="Indeterminada (novo, 1º uso)",
    ),
    ItemConsumo(
        codigo="05",
        nome_curto="Óleo essencial 10 ml – Lavanda (Lavandula angustifolia)",
        especificacao=(
            "Óleo essencial 100% puro de Lavanda (Lavandula angustifolia), natural, livre de "
            "aditivos, solventes, fragrâncias sintéticas, conservantes ou diluição em óleo "
            "carreador. Apresentação em frasco de vidro âmbar de 10 ml com batoque gotejador e "
            "tampa rosqueável com lacre de inviolabilidade. Rotulagem em língua portuguesa com "
            "nome popular e espécie botânica, volume, lote, data de fabricação, validade e "
            "regularização sanitária vigente na ANVISA (notificação, cadastro ou registro "
            "conforme enquadramento regulatório), acompanhado de laudo/certificado de análise ou "
            "cromatografia gasosa (CG) do lote comprovando pureza. Validade mínima de 12 (doze) "
            "meses a partir da data da entrega (≥ 75% do prazo total; recomendável ≥ 18 meses)."
        ),
        catmat="[A CONFIRMAR]",
        unidade="FR",
        quantidade=60,
        valor_unitario=Decimal("60.00"),
        destinacao="UBS (kits bimestrais)",
        familia="Aromaterapia",
        anvisa="Obrigatória (notificação/cadastro/registro ANVISA + laudo/CG)",
        validade_minima="≥ 12 meses (≥ 75% vida útil)",
    ),
    ItemConsumo(
        codigo="06",
        nome_curto="Óleo essencial 10 ml – Hortelã-pimenta / Peppermint (Mentha piperita)",
        especificacao=(
            "Óleo essencial 100% puro de Hortelã-pimenta / Peppermint (Mentha piperita), natural, "
            "livre de aditivos, solventes, fragrâncias sintéticas, conservantes ou diluição em "
            "óleo carreador. Apresentação em frasco de vidro âmbar de 10 ml com batoque gotejador "
            "e tampa rosqueável com lacre de inviolabilidade. Rotulagem em língua portuguesa com "
            "nome popular e espécie botânica, volume, lote, data de fabricação, validade e "
            "regularização sanitária vigente na ANVISA (notificação, cadastro ou registro "
            "conforme enquadramento regulatório), acompanhado de laudo/certificado de análise ou "
            "cromatografia gasosa (CG) do lote comprovando pureza. Validade mínima de 12 (doze) "
            "meses a partir da data da entrega (≥ 75% do prazo total; recomendável ≥ 18 meses)."
        ),
        catmat="[A CONFIRMAR]",
        unidade="FR",
        quantidade=60,
        valor_unitario=Decimal("60.00"),
        destinacao="UBS (kits bimestrais)",
        familia="Aromaterapia",
        anvisa="Obrigatória (notificação/cadastro/registro ANVISA + laudo/CG)",
        validade_minima="≥ 12 meses (≥ 75% vida útil)",
    ),
    ItemConsumo(
        codigo="07",
        nome_curto="Óleo essencial 10 ml – Melaleuca (Melaleuca alternifolia)",
        especificacao=(
            "Óleo essencial 100% puro de Melaleuca / Tea Tree (Melaleuca alternifolia), natural, "
            "livre de aditivos, solventes, fragrâncias sintéticas, conservantes ou diluição em "
            "óleo carreador. Apresentação em frasco de vidro âmbar de 10 ml com batoque gotejador "
            "e tampa rosqueável com lacre de inviolabilidade. Rotulagem em língua portuguesa com "
            "nome popular e espécie botânica, volume, lote, data de fabricação, validade e "
            "regularização sanitária vigente na ANVISA (notificação, cadastro ou registro "
            "conforme enquadramento regulatório), acompanhado de laudo/certificado de análise ou "
            "cromatografia gasosa (CG) do lote comprovando pureza. Validade mínima de 12 (doze) "
            "meses a partir da data da entrega (≥ 75% do prazo total; recomendável ≥ 18 meses)."
        ),
        catmat="[A CONFIRMAR]",
        unidade="FR",
        quantidade=60,
        valor_unitario=Decimal("60.00"),
        destinacao="UBS (kits bimestrais)",
        familia="Aromaterapia",
        anvisa="Obrigatória (notificação/cadastro/registro ANVISA + laudo/CG)",
        validade_minima="≥ 12 meses (≥ 75% vida útil)",
    ),
    ItemConsumo(
        codigo="08",
        nome_curto="Óleo essencial 10 ml – Laranja Doce (Citrus sinensis)",
        especificacao=(
            "Óleo essencial 100% puro de Laranja Doce (Citrus sinensis), natural, livre de "
            "aditivos, solventes, fragrâncias sintéticas, conservantes ou diluição em óleo "
            "carreador. Apresentação em frasco de vidro âmbar de 10 ml com batoque gotejador e "
            "tampa rosqueável com lacre de inviolabilidade. Rotulagem em língua portuguesa com "
            "nome popular e espécie botânica, volume, lote, data de fabricação, validade e "
            "regularização sanitária vigente na ANVISA (notificação, cadastro ou registro "
            "conforme enquadramento regulatório), acompanhado de laudo/certificado de análise ou "
            "cromatografia gasosa (CG) do lote comprovando pureza. Validade mínima de 12 (doze) "
            "meses a partir da data da entrega (≥ 75% do prazo total de validade)."
        ),
        catmat="[A CONFIRMAR]",
        unidade="FR",
        quantidade=60,
        valor_unitario=Decimal("60.00"),
        destinacao="UBS (kits bimestrais)",
        familia="Aromaterapia",
        anvisa="Obrigatória (notificação/cadastro/registro ANVISA + laudo/CG)",
        validade_minima="≥ 12 meses (≥ 75% vida útil)",
    ),
    ItemConsumo(
        codigo="09",
        nome_curto="Óleo vegetal carreador de semente de uva 250 ml (Vitis vinifera)",
        especificacao=(
            "Óleo vegetal carreador 100% puro de semente de uva (Vitis vinifera), natural, "
            "obtido por prensagem a frio, indicado para diluição e veiculação de óleos essenciais "
            "em aplicações tópicas nas PICS. Livre de solventes, fragrâncias sintéticas, "
            "conservantes ou aditivos químicos; aspecto líquido límpido, incolor a ligeiramente "
            "amarelado, sem sedimentos ou impurezas visíveis. Apresentação em frasco âmbar de "
            "250 ml em vidro ou PET de alta resistência com tampa vedante e lacre de "
            "inviolabilidade. Rotulagem em língua portuguesa e regularização sanitária vigente "
            "(notificação, cadastro ou registro conforme disposições da ANVISA). Validade mínima "
            "de 12 (doze) meses a partir da data da entrega (≥ 75% do prazo total de validade)."
        ),
        catmat="[A CONFIRMAR]",
        unidade="FR",
        quantidade=16,
        valor_unitario=Decimal("70.00"),
        destinacao="UBS / eMulti",
        familia="Aromaterapia",
        anvisa="Obrigatória (notificação/cadastro/registro ANVISA)",
        validade_minima="≥ 12 meses (≥ 75% vida útil)",
    ),
    ItemConsumo(
        codigo="10",
        nome_curto="Frasco âmbar 5 ml com batoque gotejador e tampa",
        especificacao=(
            "Frasco vazio em vidro na coloração marrom âmbar (proteção contra luz/UV), capacidade "
            "útil de 5 ml, acompanhado de batoque gotejador compatível e tampa plástica rosqueável "
            "com vedação estanque (preferencialmente com lacre de segurança). Dimensões "
            "aproximadas: altura de 50 mm (±5 mm), diâmetro externo de 22 mm (±3 mm) e gargalo "
            "padrão de 18 mm (rosca DIN 18 ou equivalente compatível com o batoque fornecido). "
            "Produto novo, limpo, livre de trincas, lascas ou rebarbas."
        ),
        catmat="[A CONFIRMAR]",
        unidade="UN",
        quantidade=400,
        valor_unitario=Decimal("5.00"),
        destinacao="UBS / eMulti (50 un. × 8 equipes)",
        familia="Aromaterapia",
        anvisa="Dispensado (embalagem vazia)",
        validade_minima="Indeterminada (novo, 1º uso)",
    ),
    ItemConsumo(
        codigo="11",
        nome_curto="Inalador nasal portátil para aromaterapia (com pavio de algodão)",
        especificacao=(
            "Inalador nasal portátil para aromaterapia (tipo bastão/tubo inalador vazio, de uso "
            "individual), composto por invólucro externo plástico rígido atóxico resistente a "
            "óleos essenciais com tampa rosqueável vedante, tubo interno perfurado, tampa de "
            "fundo vedante e pavio/refil absorvente interno em algodão puro. Dimensões "
            "aproximadas: corpo externo com 6,6 cm de comprimento e 1,7 cm de diâmetro "
            "(tolerância de ±15%); pavio de algodão com 5,0 cm de comprimento e 0,8 cm de "
            "diâmetro, compatível com o tubo interno. Produto novo, limpo e pronto para montagem."
        ),
        catmat="[A CONFIRMAR]",
        unidade="UN",
        quantidade=240,
        valor_unitario=Decimal("7.00"),
        destinacao="eMulti (30 un. × 8 equipes)",
        familia="Aromaterapia",
        anvisa="Dispensado (acessório vazio)",
        validade_minima="Indeterminada (novo, 1º uso)",
    ),
    ItemConsumo(
        codigo="12",
        nome_curto="Tapete de yoga / práticas corporais em EVA (≥ 5 mm)",
        especificacao=(
            "Tapete para yoga e práticas corporais/meditação, confeccionado em EVA (etileno-"
            "acetato de vinila) atóxico, flexível, enrolável, lavável e com superfície texturizada "
            "antiderrapante. Dimensões: comprimento entre 170 cm e 183 cm, largura entre 50 cm e "
            "68 cm e espessura mínima de 5 mm (admitindo-se faixa de 5 mm a 8 mm). Produto novo, "
            "de primeiro uso, sem deformações ou odor residual acentuado."
        ),
        catmat="[A CONFIRMAR]",
        unidade="UN",
        quantidade=20,
        valor_unitario=Decimal("80.00"),
        destinacao="Cursos / práticas coletivas",
        familia="Práticas corporais",
        anvisa="Dispensado (artigo de práticas corporais)",
        validade_minima="Indeterminada (novo, 1º uso)",
    ),
    ItemConsumo(
        codigo="13",
        nome_curto="Apalpador de pressão com mola padrão Nogier (aço inox)",
        especificacao=(
            "Apalpador de pressão com mola interna, padrão/tipo Nogier (designação técnica do "
            "instrumento, sem vinculação a marca comercial), destinado à localização de pontos "
            "auriculares em auriculoterapia; confeccionado em aço inoxidável de grau "
            "clínico/cirúrgico, resistente à desinfecção e esterilização em autoclave, com ponta "
            "arredondada atraumática e corpo ergonômico/antiderrapante. Dimensões aproximadas: "
            "comprimento de 13 cm a 16 cm (referência 15 cm) e diâmetro de 6 mm a 9 mm (referência "
            "8 mm). Produto com regularização sanitária vigente na ANVISA (notificação/cadastro "
            "conforme RDC nº 751/2022 ou comprovação oficial de enquadramento) e embalagem "
            "íntegra identificada em português."
        ),
        catmat="[A CONFIRMAR]",
        unidade="UN",
        quantidade=80,
        valor_unitario=Decimal("40.00"),
        destinacao="UBS / rede APS e cursos",
        familia="Auriculoterapia",
        anvisa="Obrigatória (Instrumental — RDC nº 751/2022)",
        validade_minima="Indeterminada (novo, 1º uso)",
    ),
    ItemConsumo(
        codigo="14",
        nome_curto="Plástico para plastificação térmica A4, 0,07 mm (pacote com 100 un.)",
        especificacao=(
            "Filme plástico duplo para plastificação térmica (tipo polaseal), formato padrão A4 "
            "(aproximadamente 220 mm × 307 mm), espessura de 0,07 mm (equivalente a 175 micras / "
            "7 mil), transparente, acabamento brilhante, com cantos arredondados e solda em uma "
            "das bordas, compatível com plastificadora térmica de mesa (Item 05 do processo "
            "correlato de Materiais Permanentes). Apresentação: pacote ou caixa lacrada contendo "
            "100 (cem) folhas duplas."
        ),
        catmat="[A CONFIRMAR]",
        unidade="PCT",
        quantidade=2,
        valor_unitario=Decimal("160.00"),
        destinacao="Coordenação PICSPel",
        familia="Apoio educativo/logístico",
        anvisa="Dispensado (material de expediente)",
        validade_minima="Indeterminada (novo, 1º uso)",
    ),
    ItemConsumo(
        codigo="15",
        nome_curto="Sal grosso marinho iodado (pacote de 1 kg)",
        especificacao=(
            "Sal grosso marinho iodado, 100% natural, sem corantes ou fragrâncias artificiais, "
            "destinado às práticas de escalda-pés, acondicionado em embalagem plástica íntegra, "
            "resistente e lacrada de 1 kg (um quilograma), rotulada em língua portuguesa conforme "
            "legislação sanitária vigente (identificação do produto, fabricante/envasador, lote, "
            "data de fabricação, validade e conformidade com os padrões sanitários de iodação — "
            "RDC ANVISA nº 604/2022 e normas correlatas). Validade mínima de 12 (doze) meses a "
            "partir da data da entrega."
        ),
        catmat="[A CONFIRMAR]",
        unidade="KG",
        quantidade=160,
        valor_unitario=Decimal("5.00"),
        destinacao="eMulti / práticas de escalda-pés",
        familia="Escalda-pés",
        anvisa="Conformidade sanitária/rotulagem (RDC nº 604/2022)",
        validade_minima="≥ 12 meses na entrega",
    ),
]


TOTAL_CONSUMO: Decimal = sum((i.valor_total for i in ITENS_CONSUMO), Decimal("0.00"))
TOTAL_UNIDADES: int = sum(i.quantidade for i in ITENS_CONSUMO)
DIFERENCA_NOMINAL: Decimal = META["valor_nominal_original_pdf"] - TOTAL_CONSUMO
SUBTOTAL_MATERIAIS: Decimal = TOTAL_CONSUMO + META["valor_processo_permanentes"]
TOTAL_CONSOLIDADO_EMENDA: Decimal = SUBTOTAL_MATERIAIS + META["valor_processo_servicos"]
SALDO_RESIDUAL_EMENDA: Decimal = META["teto_emenda_global"] - TOTAL_CONSOLIDADO_EMENDA

assert TOTAL_CONSUMO == Decimal("45520.00"), f"Soma esperada 45520.00, obtida {TOTAL_CONSUMO}"
assert TOTAL_UNIDADES == 1508, f"Unidades esperadas 1508, obtidas {TOTAL_UNIDADES}"
assert DIFERENCA_NOMINAL == Decimal("1680.00")
assert SUBTOTAL_MATERIAIS == Decimal("57120.00")
assert TOTAL_CONSOLIDADO_EMENDA == Decimal("97120.00")
assert SALDO_RESIDUAL_EMENDA == Decimal("2880.00")
