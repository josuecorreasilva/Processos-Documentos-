# Resumo da Revisão Técnica — Processo 3: PICSPel (Materiais de Consumo)

**Órgão:** Prefeitura Municipal de Pelotas / Secretaria Municipal de Saúde (SMS) — Diretoria de Atenção Primária / Programa PICSPel  
**Fonte de Recurso:** Emenda Parlamentar nº `43530013` (valor global: `R$ 100.000,00`)  
**Natureza da Despesa:** `3.3.90.30` — Material de Consumo (Despesa Corrente / GND 3)  
**Regime Jurídico:** Lei Federal nº 14.133/2021, LC nº 123/2006, CDC (Lei nº 8.078/1990), Lei nº 12.305/2010 (PNRS), RDC ANVISA nº 751/2022, RDC ANVISA nº 604/2022, RDC ANVISA nº 222/2018 (PGRSS), Portaria STN nº 448/2002 e Decreto Municipal nº 6.799/2023.  
**Equipe de Planejamento / Fiscalização:** Dariane Lima Portela (Matrícula `44151` — Gestor) | Aline Kohler Geppert (Matrícula `29371` — Fiscal) | Angela Moreira Vitória (Secretária Municipal de Saúde — Autorizo).

---

## 1. Documentos Gerados e Validados

| Documento | Arquivo | Padrão / Referência |
|---|---|---|
| DFD Revisado | `PICSPel_Materiais_Consumo_DFD_Revisado.docx` | Padrão institucional SMS Pelotas |
| ETP Revisado | `PICSPel_Materiais_Consumo_ETP_Revisado.docx` | Lei nº 14.133/2021 (art. 18, § 1º) |
| Termo de Referência (TR) | `PICSPel_Materiais_Consumo_TR.docx` | Fiel às **14 seções** do modelo `RP05_TR.docx` |
| Resumo da Revisão | `Resumo_Revisao_Materiais_Consumo.md` | Relatório técnico, conferência aritmética e checklist pré-protocolo |

Os três arquivos `.docx` são gerados e validados programaticamente por `scripts/gerar_picspel_consumo.py` (arquitetura modular em `scripts/picspel_consumo/`), com brasão oficial de Pelotas no cabeçalho, paginação nativa `Página X de Y`, checagem aritmética de todas as tabelas, verificação de igualdade das 15 especificações entre DFD, ETP e TR e ausência de termos de marca comercial ou resquícios do modelo `RP05_TR`.

---

## 2. Conferência Aritmética Exata dos 15 Itens de Consumo

| Item | Material de Consumo | Unid. | Qtd. | Valor Unit. (R$) | Valor Total (R$) | Destinação | Regularização ANVISA / Validade na Entrega |
|---:|---|:---:|---:|---:|---:|---|---|
| 01 | Bacia para escalda-pés (PP/plástico rígido) | UN | 120 | R$ 35,00 | R$ 4.200,00 | eMulti (15 × 8) | Dispensado / indeterminada |
| 02 | Protetor descartável para bacias (pct. 100 un.) | PCT | 80 | R$ 40,00 | R$ 3.200,00 | eMulti (10 × 8) | Dispensado / indet. ou ≥ 12 meses |
| 03 | Placas para auriculoterapia com cristal (cx. 30 placas / 70 pts.) | CX | 50 | R$ 190,00 | R$ 9.500,00 | UBS / rede APS | Obrigatória (RDC 751/2022) / ≥ 12 meses (≥ 75%) |
| 04 | Pinça Adson 12 cm, reta, sem dente (aço inox cirúrgico) | UN | 100 | R$ 35,00 | R$ 3.500,00 | UBS / rede APS | Obrigatória (RDC 751/2022) / indeterminada |
| 05 | Óleo essencial 10 ml – Lavanda (*Lavandula angustifolia*) | FR | 60 | R$ 60,00 | R$ 3.600,00 | UBS (kits bimestrais) | Obrigatória (ANVISA + laudo/CG) / ≥ 12 meses (≥ 75%) |
| 06 | Óleo essencial 10 ml – Hortelã-pimenta / Peppermint (*Mentha piperita*) | FR | 60 | R$ 60,00 | R$ 3.600,00 | UBS (kits bimestrais) | Obrigatória (ANVISA + laudo/CG) / ≥ 12 meses (≥ 75%) |
| 07 | Óleo essencial 10 ml – Melaleuca (*Melaleuca alternifolia*) | FR | 60 | R$ 60,00 | R$ 3.600,00 | UBS (kits bimestrais) | Obrigatória (ANVISA + laudo/CG) / ≥ 12 meses (≥ 75%) |
| 08 | Óleo essencial 10 ml – Laranja Doce (*Citrus sinensis*) | FR | 60 | R$ 60,00 | R$ 3.600,00 | UBS (kits bimestrais) | Obrigatória (ANVISA + laudo/CG) / ≥ 12 meses (≥ 75%) |
| 09 | Óleo vegetal carreador de semente de uva 250 ml (*Vitis vinifera*) | FR | 16 | R$ 70,00 | R$ 1.120,00 | UBS / eMulti | Obrigatória (ANVISA) / ≥ 12 meses (≥ 75%) |
| 10 | Frasco âmbar 5 ml com batoque gotejador e tampa | UN | 400 | R$ 5,00 | R$ 2.000,00 | UBS / eMulti (50 × 8) | Dispensado (embalagem vazia) / indeterminada |
| 11 | Inalador nasal portátil para aromaterapia (com pavio de algodão) | UN | 240 | R$ 7,00 | R$ 1.680,00 | eMulti (30 × 8) | Dispensado (acessório vazio) / indeterminada |
| 12 | Tapete de yoga / práticas corporais em EVA (≥ 5 mm) | UN | 20 | R$ 80,00 | R$ 1.600,00 | Cursos / práticas coletivas | Dispensado / indeterminada |
| 13 | Apalpador de pressão com mola padrão Nogier (aço inox) | UN | 80 | R$ 40,00 | R$ 3.200,00 | UBS / rede APS e cursos | Obrigatória (RDC 751/2022) / indeterminada |
| 14 | Plástico para plastificação térmica A4, 0,07 mm (pct. 100 un.) | PCT | 2 | R$ 160,00 | R$ 320,00 | Coordenação PICSPel | Dispensado / indeterminada |
| 15 | Sal grosso marinho iodado (pacote de 1 kg) | KG | 160 | R$ 5,00 | R$ 800,00 | eMulti / escalda-pés | Conformidade RDC 604/2022 / ≥ 12 meses |
| **—** | **TOTAL GERAL CONFERIDO (15 itens)** | **—** | **1.508** | **—** | **R$ 45.520,00** | **8 eMulti + UBS + Cursos + Coord.** | — |

> **Correção aritmética crítica:** Nos PDFs originais do DFD e do ETP, o total estimado constava nominalmente como **R$ 47.200,00**. A conferência linha a linha dos 15 itens (`4.200 + 3.200 + 9.500 + 3.500 + 3.600 + 3.600 + 3.600 + 3.600 + 1.120 + 2.000 + 1.680 + 1.600 + 3.200 + 320 + 800`) totaliza exatamente **R$ 45.520,00 (quarenta e cinco mil, quinhentos e vinte reais)** — diferença a menor de **R$ 1.680,00** em relação ao valor nominal anterior. Todos os 3 documentos revisados adotam o somatório real conferido de **R$ 45.520,00**, com nota explicativa transparente no DFD e no ETP.

---

## 3. Coerência Consolidada da Emenda Parlamentar nº 43530013 (3 Processos)

| Processo | Objeto | Elemento | Itens | Valor Original PDFs | Soma Real Conferida | Teto da Cota na Emenda |
|---|---|---|---:|---:|---:|---:|
| Processo 1 — Serviços | Capacitação presencial em Auriculoterapia (80h; 30 a 50 vagas) | `3.3.90.39` | 1 | R$ 40.000,00 | **R$ 40.000,00** | R$ 40.000,00 |
| Processo 2 — Materiais Permanentes | Equipamentos/apoio (caixas de som, garrafas, maletas, jarras, plastificadora) | `4.4.90.52`* | 5 | R$ 11.200,00 | **R$ 11.600,00** | *(cota conjunta R$ 60.000,00)* |
| **Processo 3 — Materiais de Consumo (este processo)** | **15 insumos para escalda-pés, auriculoterapia, aromaterapia e práticas corporais** | **`3.3.90.30`** | **15** | R$ 47.200,00 | **R$ 45.520,00** | **R$ 60.000,00** |
| **Total Consolidado dos 3 Processos** | — | — | **21** | R$ 98.400,00 | **R$ 97.120,00** | **R$ 100.000,00** *(saldo residual: **R$ 2.880,00**)* |

- **Zero sobreposição e complementaridade técnica:** o subtotal de materiais (`R$ 45.520,00 + R$ 11.600,00 = R$ 57.120,00`) respeita o teto de `R$ 60.000,00`, com saldo residual de `R$ 2.880,00`. Os plásticos A4 0,07 mm (Item 14) integram-se à plastificadora (Item 05 do Processo 2); as bacias e sal grosso complementam as jarras elétricas e garrafas térmicas; e as maletas rígidas acondicionam os kits das 8 equipes eMulti.
- **\*Observação contábil:** mantido registro da necessidade de validação prévia pela Contabilidade/CCGOF, no âmbito do Processo 2, quanto ao empenho em Despesa de Capital (`4.4.90.52` / GND 4) na programação da Emenda nº 43530013, bem como da validação dos subelementos definitivos de `3.3.90.30.__` neste Processo 3.
- **Responsáveis unificados:** Dariane Lima Portela (Mat. `44151` — Gestor), Aline Kohler Geppert (Mat. `29371` — Fiscal) e Angela Moreira Vitória (Secretária Municipal de Saúde).

---

## 4. Principais Ajustes Realizados no DFD e no ETP

### 4.1 Ajustes no DFD (`PICSPel_Materiais_Consumo_DFD_Revisado.docx`)
1. **Título e distinção inequívoca de processo:** atualizado de `PICSPel (MATERIAIS)` para `PICSPel (MATERIAIS DE CONSUMO)`, diferenciando-o expressamente do processo de Materiais Permanentes (`4.4.90.52`) e do processo de Serviços (`3.3.90.39`).
2. **Correção aritmética global (Seções 1, 4 e 5):** atualizado o total de `R$ 47.200,00` para **R$ 45.520,00** (`1.508` unidades comerciais), com quadro explicativo da conferência linha a linha (`−R$ 1.680,00`).
3. **Modalidade, ME/EPP e segurança jurídica (Seção 2):** mantido **Pregão Eletrônico (aquisição pontual — sem SRP)** com critério de **menor preço por item**, incluída a **exclusividade para ME/EPP em todos os 15 itens** (art. 48, I, da LC nº 123/2006 c/c art. 4º da Lei nº 14.133/2021 — maior item é R$ 9.500,00) e nota técnica fundamentando a escolha frente ao teto de dispensa por valor (art. 75, II, e prevenção ao fracionamento do § 1º).
4. **Especificações técnicas sem direcionamento de marca (Seção 4):** revisados todos os 15 itens (detalhados na Seção 4.2 abaixo), garantindo redação idêntica entre DFD, ETP e TR.
5. **Dotação, consolidação da emenda e PCA (Seções 3 e 5):** estruturados os campos de `Processo/Memorando`, `Projeto/Atividade`, `Fonte`, `Reduzido/Dotação`, `PCA 2026`, consolidação dos 3 processos (`R$ 97.120,00`) e alerta de obrigatoriedade de reserva orçamentária prévia (art. 150 da Lei nº 14.133/2021).
6. **Local de entrega e prazo (Seções 6 e 7):** incluído o endereço oficial completo do Almoxarifado da SMS Pelotas (`Av. Fernando Osório, nº 273, Bairro Três Vendas, Pelotas/RS, CEP 96020-150`, horários e e-mail de agendamento `smspel.almoxarifado@gmail.com`), prazo de entrega de 30 dias corridos e regra de validade mínima de 12 meses na entrega.

### 4.2 Ajustes no ETP (`PICSPel_Materiais_Consumo_ETP_Revisado.docx`)
1. **Revisão técnica dos 15 itens (eliminação de marca, correção de unidades e inclusão de tolerâncias):**
   - **Item 01 (Bacia):** incluída unidade de medida (`cm`), variação dimensional razoável (`±15%` no comprimento/largura, altura mínima `8 cm`), resistência à água morna (≥ 50 °C) e desinfecção;
   - **Item 02 (Protetor descartável):** incluída unidade (`68 cm × 68 cm`, tolerância `±10%`), resistência à ruptura e ajuste compatível com a bacia do Item 01;
   - **Item 03 (Placas para auriculoterapia):** **substituída a marca registrada `"micropore"`** pelo termo técnico genérico **`"fita adesiva microporosa hipoalergênica"`** (art. 41, I, da Lei nº 14.133/2021), incluída faixa para esfera de cristal (`1,5 mm a 1,6 mm`, ref. `1,58 mm`), exigência de regularização ANVISA (RDC nº 751/2022) e validade mínima de 12 meses (≥ 75% da vida útil);
   - **Item 04 (Pinça Adson 12 cm):** especificados aço inox cirúrgico (AISI 420 ou equivalente), autoclavabilidade, tolerância de `±0,5 cm` e regularização ANVISA (RDC nº 751/2022);
   - **Itens 05 a 08 (Óleos essenciais 10 ml):** padronizado nome em português (`Hortelã-pimenta / Peppermint`), incluído lacre de inviolabilidade, **adequado o requisito ANVISA** de `"registro"` estrito para **`"regularização sanitária vigente na ANVISA (notificação, cadastro ou registro conforme enquadramento regulatório)"`** (evitando desclassificação indevida de produtos cosméticos Grau 1 regularmente notificados), acrescida exigência de **laudo/certificado de análise ou cromatografia gasosa (CG)** do lote para garantir pureza de 100%, e validade mínima de 12 meses (≥ 75% da vida útil, compatível inclusive com o óleo cítrico de Laranja Doce);
   - **Item 09 (Óleo vegetal carreador 250 ml):** incluídos lacre de inviolabilidade, regularização sanitária ANVISA e validade mínima de 12 meses (≥ 75% da vida útil);
   - **Item 10 (Frasco âmbar 5 ml):** incluídas tolerâncias dimensionais (`50 mm ±5 mm`, `22 mm ±3 mm`, gargalo padrão `18 mm / rosca DIN 18`) e proteção UV/vedação estanque;
   - **Item 11 (Inalador nasal):** corrigida grafia decimal (`6.6`/`1.7` → `6,6 cm`/`1,7 cm`), substituído `"Aroma stick"` por descrição genérica (`tipo bastão/tubo inalador portátil`), incluída tolerância (`±15%`) e especificação completa das peças;
   - **Item 12 (Tapete de yoga):** ajustada espessura fixa `"5mm"` para **`"mínima de 5 mm (faixa de 5 mm a 8 mm)"`** e superfície antiderrapante lavável;
   - **Item 13 (Apalpador Nogier):** corrigida abreviação (`"C/ Mola"`), esclarecido tratar-se de designação técnica sem vinculação a marca, incluídas faixas dimensionais (`13–16 cm` / `6–9 mm`), autoclavabilidade e regularização ANVISA (RDC nº 751/2022);
   - **Item 14 (Plástico para plastificação A4):** especificada equivalência `0,07 mm / 175 micras (polaseal)` e compatibilidade expressa com a plastificadora térmica do Processo 2 (Permanentes);
   - **Item 15 (Sal grosso marinho iodado 1 kg):** detalhada rotulagem em português, ausência de aditivos/fragrâncias, conformidade com RDC ANVISA nº 604/2022 e validade mínima de 12 meses.
2. **Coerência da Emenda e contratações correlatas (Seções 1.4, 7 e 12):** corrigida a omissão do processo de Materiais Permanentes (`4.4.90.52` — R$ 11.600,00), que não era citado no ETP original, apresentando tabela consolidada dos 3 processos (`R$ 97.120,00`) e o vínculo técnico-operacional entre eles.
3. **Requisitos sanitários proporcionais (Seções 4 e 5.1):** reorganizadas todas as 4 famílias de produtos (incluindo escalda-pés e práticas corporais, antes ausentes nos tópicos da Seção 4) e definida matriz clara de exigência ANVISA/AFE (obrigatória apenas nos itens regulados — Itens 03, 04, 05–09 e 13 — e expressamente dispensada para os Itens 01, 02, 10, 11, 12 e 14).
4. **Riscos ampliados (Seção 14):** incluído o **Risco 5** (validade exígua, irregularidade ANVISA ou pureza comprometida/adulteração de óleos essenciais) e atualizadas as ações de mitigação e contingência com base na Lei nº 14.133/2021 e no Decreto Municipal nº 6.799/2023.

---

## 5. Estrutura do Termo de Referência Criado (`PICSPel_Materiais_Consumo_TR.docx`)

Elaborado fielmente nas **14 seções** do modelo oficial `RP05_TR.docx`, adaptado integralmente para aquisição pontual (sem SRP) de materiais de consumo/insumos de PICS:

1. **1 DO OBJETO, VALOR E CONDIÇÕES GERAIS DA CONTRATAÇÃO** — objeto completo, tabela de 7 colunas com `CATMAT` (`[A CONFIRMAR]`), 15 itens / 1.508 unidades / `R$ 45.520,00`, bens comuns não de luxo (arts. 6º, XIII, e 20), substituição de contrato por Nota de Empenho (art. 95, II), aquisição pontual sem SRP, parcelamento por item + exclusividade ME/EPP e prevalência do TR sobre o CATMAT.
2. **2 DA FUNDAMENTAÇÃO DA CONTRATAÇÃO** — PNPIC, PICSPel, 8 equipes eMulti, UBS, cursos (30–50 vagas), Emenda nº 43530013, justificativa da não adoção do SRP e campo PCA 2026.
3. **3 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO CONSIDERADO O CICLO DE VIDA DO OBJETO** — entrega, conferência sanitária/validade, estocagem protegida de luz/calor, distribuição, reprocessamento de inox e descarte (PNRS + PGRSS).
4. **4 GESTÃO/ANÁLISE DE RISCOS** — síntese dos 5 riscos do ETP e remissão às sanções da Lei nº 14.133/2021 e do Decreto Municipal nº 6.799/2023.
5. **5 REQUISITOS DA CONTRATAÇÃO** — vedação à marca (5.1), requisitos por família (5.2), embalagem/rotulagem/proteção de vidros (5.3), validade mínima de 12 meses e ≥ 75% da vida útil (5.4), matriz diferenciada de regularização ANVISA e laudos/CG dos óleos essenciais (5.5), amostras em 5 dias úteis (5.6), habilitação proporcional com AFE/Licença apenas nos itens regulados (5.7), vedação à subcontratação (5.8), frete CIF (5.9) e substituição em 5 dias úteis (5.10).
6. **6 DA EXECUÇÃO DO OBJETO** — endereço completo e agendamento no Almoxarifado SMS, prazo de entrega de 30 dias corridos, pedido de prorrogação com 5 dias corridos de antecedência, cuidados de transporte/validade e reunião pós-entrega.
7. **7 DA GESTÃO E FISCALIZAÇÃO DA CONTRATAÇÃO** — execução vinculada à Nota de Empenho, designação de Dariane Lima Portela (Gestor) e Aline Kohler Geppert (Fiscal), comunicações formais, sanções (arts. 155–163 e Decreto Municipal nº 6.799/2023) e dispensa de garantia contratual (art. 96).
8. **8 RECEBIMENTO DO OBJETO** — recebimento provisório em até 7 dias corridos (art. 140, II, "a"), rejeição/substituição em até 5 dias úteis e recebimento definitivo em até 8 dias corridos após conferência qualitativa, sanitária, de validade e laudos (art. 140, II, "b").
9. **9 DOS CRITÉRIOS DE MEDIÇÃO E PAGAMENTO** — pagamento por item após recebimento definitivo e ateste, ordem cronológica (art. 141), vedação a pagamento antecipado (art. 145) e dados obrigatórios da NF/DANFE (lote, validade e `Emenda Parlamentar nº 43530013 – Programa PICSPel`).
10. **10 DA FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR** — Pregão Eletrônico (sem SRP), menor preço por item, exclusividade ME/EPP (art. 48, I, LC nº 123/2006) e pesquisa formal de preços (art. 23).
11. **11 DAS OBRIGAÇÕES DO CONTRATANTE** — 10 obrigações adaptadas ao modelo oficial.
12. **12 DAS OBRIGAÇÕES DO CONTRATADO** — 11 obrigações incluindo conformidade ANVISA, laudos/CG, validade mínima, CDC, trabalho infantil, reserva de cargos (art. 92, XVII), frete CIF e emissão de NF detalhada.
13. **13 DA ADEQUAÇÃO ORÇAMENTÁRIA** — obrigatoriedade prévia de dotação e Nota de Reserva Orçamentária (art. 150, por não ser SRP), elemento `3.3.90.30`, valor `R$ 45.520,00` e conciliação da Emenda nº 43530013.
14. **14 EQUIPE DE PLANEJAMENTO** — certificação de conformidade e bloco de assinaturas (Gestor, Fiscal e Secretária Municipal de Saúde).

---

## 6. Checklist de Pendências Antes do Protocolo

Antes da autuação do processo administrativo e publicação do edital, a SMS deverá providenciar:

- [ ] **Numeração administrativa:** preencher `DFD nº [___]/2026` e `Processo Administrativo / Memorando nº [______]/2026` (DFD, ETP e TR).
- [ ] **Códigos CATMAT (TR subitem 1.1):** consultar e preencher os 15 códigos `[A CONFIRMAR]` no catálogo do Compras.gov.br.
- [ ] **Plano de Contratações Anual (PCA 2026):** informar o número do item no PCA ou anexar justificativa formal de inclusão excepcional vinculada à Emenda Parlamentar nº `43530013`.
- [ ] **Pesquisa formal de preços (art. 23 da Lei nº 14.133/2021):** consolidar cotações/fontes (PNCP/Painel de Preços, contratações similares e/ou fornecedores), anexar o Mapa Comparativo de Preços com critério estatístico justificado e validar os valores referenciais (`R$ 45.520,00`).
- [ ] **Validação Contábil / CCGOF e Reserva Orçamentária:** preencher `Projeto/Atividade`, `Fonte de Recurso`, `Reduzido/Dotação`, confirmar os subelementos definitivos de `3.3.90.30.__` e a dispensa de tombamento patrimonial (Portaria STN nº 448/2002), conciliar o saldo residual global da Emenda nº `43530013` (`R$ 2.880,00`) e anexar a **Nota de Reserva Orçamentária** (obrigatória antes do edital — art. 150).
- [ ] **Consulta prévia a atas vigentes:** certificar a existência ou não de atas vigentes (Consórcio Público do Extremo Sul / outros órgãos) mais vantajosas para itens específicos.
- [ ] **Data e assinaturas:** preencher a data (`Pelotas/RS, _____ de ______________ de 2026`) e colher assinaturas de Dariane Lima Portela (Gestor), Aline Kohler Geppert (Fiscal) e Angela Moreira Vitória (Secretária Municipal de Saúde).

---

## 7. Reprodutibilidade

Para regerar e validar programaticamente os três documentos `.docx`:

```bash
python3 scripts/gerar_picspel_consumo.py
```
