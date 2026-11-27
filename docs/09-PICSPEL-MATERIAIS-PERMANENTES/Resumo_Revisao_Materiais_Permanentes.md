# Resumo da Revisão Técnica — Processo 2: PICSPel (Materiais Permanentes)

**Órgão:** Prefeitura Municipal de Pelotas / Secretaria Municipal de Saúde (SMS) — Diretoria de Atenção Primária / Programa PICSPel  
**Fonte de Recurso:** Emenda Parlamentar nº 43530013 (valor global: R$ 100.000,00)  
**Natureza da Despesa:** 4.4.90.52 — Equipamentos e Material Permanente (Despesa de Capital / GND 4)  
**Base Legal:** Lei Federal nº 14.133/2021, LC nº 123/2006, CDC (Lei nº 8.078/1990), Lei nº 12.305/2010 (PNRS), Decreto Federal nº 10.240/2020, Resolução CONAMA nº 401/2008, Portaria STN nº 448/2002 e Decreto Municipal nº 6.799/2023.  
**Equipe de Planejamento / Fiscalização:** Dariane Lima Portela (Matrícula 44151 — Gestor) | Aline Kohler Geppert (Matrícula 29371 — Fiscal) | Angela Moreira Vitória (Secretária Municipal de Saúde — Autorizo).

---

## 1. Documentos Gerados e Validados

| Documento | Arquivo | Padrão / Referência |
|---|---|---|
| DFD Revisado | `PICSPel_Materiais_Permanentes_DFD_Revisado.docx` | Padrão institucional SMS Pelotas |
| ETP Revisado | `PICSPel_Materiais_Permanentes_ETP_Revisado.docx` | Lei nº 14.133/2021 (art. 18, § 1º) |
| Termo de Referência (TR) | `PICSPel_Materiais_Permanentes_TR.docx` | Fiel à estrutura de 14 seções do modelo `RP05_TR` |
| Resumo da Revisão | `Resumo_Revisao_Materiais_Permanentes.md` | Relatório técnico e checklist pré-protocolo |

Os três arquivos `.docx` são gerados e validados programaticamente por `scripts/gerar_picspel_permanentes.py` (com brasão oficial de Pelotas no cabeçalho, paginação nativa `Página X de Y`, verificação aritmética de todas as tabelas e checagem de ausência de resquícios de outros processos).

---

## 2. Quadro Consolidado dos Itens Permanentes (Corrigido)

| Item | Bem Permanente | Unid. | Qtd. | Valor Unit. (R$) | Valor Total (R$) | Destinação | Certificação / Conformidade |
|---:|---|:---:|---:|---:|---:|---|---|
| 01 | Caixa de som portátil recarregável (≥ 20 W RMS, estéreo, Bluetooth v5.0+, 220 V/bivolt/USB) | UN | 8 | R$ 600,00 | R$ 4.800,00 | eMulti (1 × 8 equipes) | Homologação ANATEL |
| 02 | Garrafa térmica 100% aço inox grau alimentício (aprox. 1,8 L, sem vidro, bomba de pressão) | UN | 24 | R$ 150,00 | R$ 3.600,00 | eMulti (3 × 8 equipes) | ABNT NBR 13282 / INMETRO (se aplicável) |
| 03 | Maleta organizadora rígida em ABS (tipo frasqueira, zíper duplo horizontal, alça) | UN | 8 | R$ 250,00 | R$ 2.000,00 | eMulti (1 × 8 equipes) | Ficha técnica / garantia 12 meses |
| 04 | Jarra elétrica em aço inox (1,7 L a 2,2 L, ≥ 1.500 W / ref. 1.800 W, 220 V, base 360°) | UN | 8 | R$ 100,00 | R$ 800,00 | eMulti (1 × 8 equipes) | Certificação compulsória INMETRO (Port. nº 148/2022) |
| 05 | Plastificadora/laminadora A3/A4/A5/A6 (220 V/bivolt, polaseal 0,05–0,07 mm) | UN | 1 | R$ 400,00 | R$ 400,00 | Coordenação PICSPel | Conformidade elétrica / NBR 14136 |
| **—** | **TOTAL GERAL (5 itens)** | **—** | **49** | **—** | **R$ 11.600,00** | **8 eMulti (R$ 11.200,00) + Coordenação (R$ 400,00)** | **Garantia mínima de 12 meses para todos os itens** |

> **Correção aritmética crítica:** Nos PDFs originais do DFD e do ETP, o total estimado constava como **R$ 11.200,00**, somando apenas os 4 itens destinados às equipes eMulti (`4.800 + 3.600 + 2.000 + 800 = 11.200`) e omitindo os **R$ 400,00** do Item 05 (Plastificadora). A soma real exata dos 5 itens é **R$ 11.600,00 (onze mil e seiscentos reais)**, agora unificada no DFD, ETP e TR.

---

## 3. Coerência com os Outros Dois Processos da Emenda Parlamentar nº 43530013

Conferiu-se a integridade dos três processos do Programa PICSPel custeados pela Emenda nº 43530013 (teto global: **R$ 100.000,00**):

| Processo | Elemento | Itens | Valor Original nos PDFs | Soma Aritmética Real | Sobreposição de Itens |
|---|---|---:|---:|---:|---|
| Processo 1 — Materiais de Consumo | 3.3.90.30 | 15 | R$ 47.200,00 | R$ 45.520,00* | Nenhuma |
| **Processo 2 — Materiais Permanentes (este processo)** | **4.4.90.52** | **5** | R$ 11.200,00 | **R$ 11.600,00** | **Nenhuma** |
| Processo 3 — Serviços de Capacitação (Auriculoterapia 80h) | 3.3.90.39 | 1 | R$ 40.000,00 | R$ 40.000,00 | Nenhuma |
| **Total Consolidado dos 3 Processos** | — | **21** | **R$ 98.400,00** | **R$ 97.120,00** | **Zero duplicidade** |

*\*Nota informativa para conciliação financeira pela coordenação: no PDF original do Processo 1 (Consumo), a soma real dos 15 itens é R$ 45.520,00 (e não R$ 47.200,00 — diferença de R$ 1.680,00). Tanto considerando a soma real de Consumo (R$ 45.520,00 + R$ 11.600,00 = R$ 57.120,00) quanto o valor nominal original (R$ 47.200,00 + R$ 11.600,00 = R$ 58.800,00), a soma de materiais permanece abaixo da parcela de R$ 60.000,00 e o total global abaixo dos R$ 100.000,00 da emenda.*

### Pontos de atenção entre os processos
- **Complementaridade sem duplicidade:** a plastificadora (Item 05, Processo 2) foi especificada para compatibilidade expressa com filmes polaseal de 0,05 mm a 0,07 mm, integrando-se diretamente aos plásticos A4 de 0,07 mm do Processo 1 (Consumo); as jarras elétricas e garrafas térmicas complementam as bacias e sal grosso de escalda-pés; e as maletas rígidas transportam os kits de consumo das 8 equipes eMulti.
- **Alerta contábil obrigatório (GND 4 – Capital vs. GND 3 – Custeio):** os PDFs do Processo 1 mencionam R$ 60.000,00 para custeio/materiais e R$ 40.000,00 para serviços (ambos GND 3). Como este Processo 2 classifica os 5 itens em **4.4.90.52 (Despesa de Capital / GND 4)**, inseriu-se advertência expressa no DFD, ETP e TR condicionando o prosseguimento à **validação prévia pela Contabilidade/CCGOF** quanto à possibilidade de empenho em GND 4 na programação da Emenda nº 43530013 (ou indicação de fonte/readequação orçamentária), bem como ao enquadramento patrimonial municipal de cada item segundo a Portaria STN nº 448/2002.

---

## 4. Principais Ajustes Realizados no DFD e no ETP

### 4.1 Ajustes no DFD (`PICSPel_Materiais_Permanentes_DFD_Revisado.docx`)
1. **Correção aritmética e completude da tabela (Seção 4):** incluídas as colunas `Valor unit. (R$)` e `Valor total (R$)` (ausentes no DFD original) e corrigido o valor total estimado de R$ 11.200,00 para **R$ 11.600,00** (49 unidades), com abertura entre equipes eMulti (R$ 11.200,00) e coordenação PICSPel (R$ 400,00).
2. **Alinhamento do objeto e destinação (Seção 1):** explicitada a destinação para as 08 equipes eMulti e para a coordenação do programa (que recebe a plastificadora), bem como a autonomia em relação aos processos de Consumo e Serviços.
3. **Modalidade e tratamento ME/EPP (Seção 2):** mantido o Pregão Eletrônico (menor preço por item), acrescida previsão de exclusividade ME/EPP (art. 48, I, da LC nº 123/2006) e nota técnica fundamentando a opção frente ao limite de dispensa por valor do art. 75, II, da Lei nº 14.133/2021 (prevenção a fracionamento no exercício — § 1º).
4. **Especificações técnicas sem direcionamento de marca (Seção 4):** revisadas todas as 5 especificações (ver detalhamento na Seção 4.2 abaixo), garantindo redação idêntica entre DFD, ETP e TR.
5. **Dotação orçamentária e PCA (Seções 3 e 5):** incluídos campos sinalizados para PCA, número do processo/memorando, dotação/reduzido e alerta formal de validação de GND 4 junto à CCGOF.
6. **Local de entrega, garantia e patrimônio (Seções 6 e 7):** incluído endereço oficial completo do Almoxarifado da SMS Pelotas (`Av. Fernando Osório, nº 273, Bairro Três Vendas, CEP 96020-150`, e-mail de agendamento `smspel.almoxarifado@gmail.com`), garantia mínima de 12 meses e obrigatoriedade de tombamento patrimonial antes da entrega às equipes.

### 4.2 Ajustes no ETP (`PICSPel_Materiais_Permanentes_ETP_Revisado.docx`)
1. **Remoção de duplicidade e inconsistência interna na Seção 4:** removido o parágrafo literalmente duplicado sobre a plastificadora e corrigida a contradição entre a Seção 4 (que falava "plastificadora A4") e a tabela da Seção 5 ("A3/A4/A5/A6").
2. **Eliminação de resquícios de copia-e-cola do ETP de Consumo:**
   - **Seção 13 (Impactos Ambientais):** substituído o texto indevido que citava *"protetores descartáveis de bacias"*, *"contato com fluidos corporais"* e *"óleos e frascos"* por análise ambiental própria de bens permanentes e eletroportáteis (embalagens recicláveis, eficiência energética, logística reversa de REEE — Decreto nº 10.240/2020 — e descarte de baterias — Resolução CONAMA nº 401/2008).
   - **Seção 4 (Requisitos):** removida referência genérica a "lote e validade" de insumos e inseridos requisitos próprios de equipamentos: número de série, manual em português, tensão **220 V / bivolt** (padrão Pelotas/RS), plugue **ABNT NBR 14136**, certificação compulsória **INMETRO**, homologação **ANATEL** (Bluetooth), garantia mínima de **12 meses** e assistência técnica autorizada no Brasil.
   - **Seção 12 (Contratações Correlatas):** incluída menção expressa ao **Processo 1 (Materiais de Consumo — 3.3.90.30)** além do Processo 3 (Serviços de Capacitação — 3.3.90.39), que antes era o único citado.
3. **Aperfeiçoamento técnico das 5 especificações (eliminação de risco de direcionamento e correção de erros):**
   - **Item 01 (Caixa de som):** corrigido `"20 RMS watts"` para **"potência mínima de 20 W RMS"**, corrigido `"resposta de frequência 20 KHz"` (incompleto) para **"faixa mínima de 65 Hz a 20 kHz"**, corrigido `"Stereo"` para **"estéreo"**, admitida tolerância dimensional de **±20%** em torno da referência 170 × 80 × 80 mm, definida autonomia mínima de 6 h, Bluetooth v5.0+, carregamento 220 V / bivolt / USB e **homologação ANATEL**.
   - **Item 02 (Garrafa térmica):** especificado aço inox de grau alimentício com dupla parede a vácuo, sem ampola de vidro, faixa de capacidade de **1,7 L a 2,0 L** (ref. 1,8 L), desempenho térmico mínimo de 6 h e referência à ABNT NBR 13282 / INMETRO.
   - **Item 03 (Maleta organizadora):** limpa a redação de marketplace (`"Frasqueira De Viagem Rígida"`, `"27.5 cm"`), substituindo medidas exatas por faixas dimensionais razoáveis (`25–30 × 32–38 × 16–22 cm`) e admitindo ABS rígido ou polímero de resistência equivalente ou superior.
   - **Item 04 (Jarra elétrica):** corrigido erro de digitação (`"automático.Capacidade"`), ajustada potência fixa `"1800W"` para **"mínima de 1.500 W (referência 1.800 W)"**, mantida tensão **220 V**, adicionados plugue NBR 14136, proteção contra funcionamento a seco e **certificação compulsória INMETRO (Portaria nº 148/2022)**.
   - **Item 05 (Plastificadora):** corrigida grafia (`"Poliseladora"`, `"220v"`), detalhados largura útil mínima (320 mm), formatos A3/A4/A5/A6, compatibilidade expressa com polaseal de **0,05 mm a 0,07 mm** (integração com o Processo 1), sistema anti-atolamento, tensão 220 V/bivolt e plugue NBR 14136.
4. **Correção aritmética e abertura por destinação (Seções 1.4, 5, 6, 7 e 15):** atualizado o total estimado para **R$ 11.600,00** com nota explícita de revisão aritmética e tabela detalhada de natureza/destinação/certificação (subtotais eMulti = R$ 11.200,00; Coordenação = R$ 400,00).
5. **Ampliação da Análise de Riscos (Seção 14):** atualizados os riscos originais e incluídos riscos críticos de bens permanentes: incompatibilidade elétrica (127 V vs. 220 V), defeito/vício durante a garantia e falha de assistência técnica, e risco orçamentário/patrimonial (GND 4 e tombamento).

---

## 5. Estrutura do Termo de Referência (TR) Criado

O arquivo `PICSPel_Materiais_Permanentes_TR.docx` reproduz fielmente a numeração e estrutura das **14 seções** do modelo oficial `RP05_TR.docx_3bae.pdf`, integralmente adaptadas para **aquisição pontual de materiais permanentes (4.4.90.52)** (sem SRP e sem exigências indevidas de ANVISA/AFE):

1. **1 DO OBJETO, VALOR E CONDIÇÕES GERAIS DA CONTRATAÇÃO** (tabela com 7 colunas do modelo incluindo `CATMAT`, 5 itens / 49 unidades / R$ 11.600,00; art. 20 e art. 6º, XIII; substituição de contrato por Nota de Empenho conforme art. 95; entrega integral por item; exclusividade ME/EPP; prevalência do TR sobre o CATMAT).
2. **2 DA FUNDAMENTAÇÃO DA CONTRATAÇÃO** (PNPIC/PICSPel; justificativa da aquisição pontual sem SRP e do Pregão Eletrônico; campo PCA).
3. **3 DESCRIÇÃO DA SOLUÇÃO COMO UM TODO CONSIDERADO O CICLO DE VIDA DO OBJETO** (entrega, teste funcional em 220 V, tombamento patrimonial, garantia de 12 meses e logística reversa REEE/baterias).
4. **4 GESTÃO/ANÁLISE DE RISCOS** (síntese de riscos de bens permanentes e remissão ao Decreto Municipal nº 6.799/2023).
5. **5 REQUISITOS DA CONTRATAÇÃO** (vedação à marca; requisitos por família; garantia mínima de 12 meses e assistência técnica sem ônus logístico; ficha técnica + INMETRO + ANATEL sem ANVISA; amostras só se necessário em 5 dias úteis; habilitação proporcional; vedação à subcontratação do fornecimento; frete CIF; substituição em 5 dias úteis).
6. **6 DA EXECUÇÃO DO OBJETO** (endereço completo do Almoxarifado SMS com agendamento; prazo de entrega de 30 dias corridos; pedido de prorrogação com 5 dias de antecedência; acessórios/manuais/NF; operação "plug and play" com teste funcional e tombamento).
7. **7 DA GESTÃO E FISCALIZAÇÃO DA CONTRATAÇÃO** (execução vinculada à Nota de Empenho/Edital; Gestor Dariane Lima Portela e Fiscal Aline Kohler Geppert; comunicações formais; sanções da Lei nº 14.133/2021 e Decreto Municipal nº 6.799/2023).
8. **8 RECEBIMENTO DO OBJETO E INCORPORAÇÃO PATRIMONIAL** (recebimento provisório em até 7 dias corridos; rejeição/substituição em 5 dias úteis; recebimento definitivo em até 8 dias corridos após teste funcional em 220 V; tombamento patrimonial e Termo de Responsabilidade).
9. **9 DOS CRITÉRIOS DE MEDIÇÃO E PAGAMENTO** (liquidação por item após recebimento definitivo e ateste; requisitos da NF/DANFE com marca/modelo/série e Emenda nº 43530013; ordem cronológica art. 141; retenções tributárias).
10. **10 DA FORMA E CRITÉRIOS DE SELEÇÃO DO FORNECEDOR** (Pregão Eletrônico sem SRP; menor preço por item; exclusividade ME/EPP; pesquisa de preços art. 23).
11. **11 DAS OBRIGAÇÕES DO CONTRATANTE** (incluindo inspeção funcional e tombamento patrimonial).
12. **12 DAS OBRIGAÇÕES DO CONTRATADO** (incluindo tensão 220 V/bivolt, NBR 14136, INMETRO/ANATEL, garantia de 12 meses, assistência técnica, reserva de cargos art. 92, XVII, e NF detalhada para tombamento).
13. **13 DA ADEQUAÇÃO ORÇAMENTÁRIA** (quadro orçamentário com Emenda nº 43530013, elemento 4.4.90.52, valor R$ 11.600,00 e campos pendentes).
14. **14 EQUIPE DE PLANEJAMENTO** (certificação e assinaturas de Dariane Lima Portela, Aline Kohler Geppert e Angela Moreira Vitória).

---

## 6. Pendências de Preenchimento Antes do Protocolo

Os campos administrativos externos não constantes dos PDFs originais foram mantidos claramente sinalizados entre colchetes `[...]` para preenchimento antes da autuação/protocolo:

1. **Numeração administrativa:** número do **DFD nº `[___]/2026`** e do **Processo Administrativo / Memorando nº `[______]/2026`** (DFD, ETP e TR).
2. **Plano de Contratações Anual (PCA):** informar o número do item no PCA 2026 ou registrar justificativa formal de inclusão excepcional vinculada à Emenda Parlamentar nº 43530013 (DFD Seção 3, ETP Seção 3, TR Seção 2.3).
3. **Validação Contábil / Orçamentária (CCGOF) — CRÍTICO:**
   - Preencher **Projeto/Atividade**, **Fonte de Recurso** e **Dotação / Reduzido** (DFD Seção 5 e TR Seção 13).
   - **Confirmar a compatibilidade de empenho no elemento 4.4.90.52 (Despesa de Capital / GND 4)** dentro da Emenda Parlamentar nº 43530013 (ou indicar fonte complementar/remanejamento) e validar o enquadramento patrimonial municipal de todos os 5 itens.
   - Confirmar a **conciliação do saldo global da Emenda nº 43530013** frente ao ajuste aritmético (+R$ 400,00 neste processo) e ao valor real do Processo 1 (Materiais de Consumo).
4. **Códigos CATMAT (TR Seção 1.1):** consultar e preencher os 5 códigos `[A CONFIRMAR]` no catálogo do Compras.gov.br antes do cadastramento no sistema de licitações.
5. **Pesquisa formal de preços (art. 23 da Lei nº 14.133/2021):** anexar o mapa comparativo de preços oficial (PNCP/Painel de Preços, contratações similares e/ou cotações) validando ou atualizando os valores unitários referenciais e o total de R$ 11.600,00 (TR Seção 10.4).
6. **Consulta prévia a atas vigentes e verificação de ramo de atividade:** certificar a inexistência de atas vantajosas (Consórcio Público do Extremo Sul / outros órgãos) e, caso o Setor de Compras avalie eventual Dispensa Eletrônica (art. 75, II), certificar ausência de fracionamento no exercício (art. 75, § 1º).
7. **Data e assinaturas:** preencher a data de emissão (`Pelotas/RS, _____ de ______________ de 2026`) e colher as assinaturas da Gestora, da Fiscal e da Secretária Municipal de Saúde.

---

## 7. Reprodutibilidade

Para regerar e validar programaticamente os três documentos `.docx`:

```bash
python3 scripts/gerar_picspel_permanentes.py
```
