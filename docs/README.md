# Pacote documental — Oxigênio Medicinal (Pelotas/SMS)

Documentos elaborados para tramitação **simultânea** de dois processos, a partir do **Processo nº 1242/2026** (licitação deserta em 29/10/2026) e do Projeto O2 / DFD atualizado.

## Decisão estratégica

| Processo | Opção | Objetivo | Base legal |
|---|---|---|---|
| **1** | **C** — Dispensa emergencial | Não deixar faltar oxigênio agora | Art. 75, VIII, Lei nº 14.133/2021 |
| **2** | **A** — Pregão eletrônico | Resolver em definitivo | Lei nº 14.133/2021 (pregão / menor preço por lote) |

Prazo sugerido da emergência: **180 dias** (ou até homologar o pregão, o que ocorrer primeiro), **sem prorrogação ordinária**.

Consumo de referência (DFD): **206 m³/mês** (182 urbano + 24 rural) · **2.472 m³/ano**.

## Como usar (ordem sugerida)

1. Assinar o [`00-despacho-encaminhamento-opcao-a-e-c.docx`](00-despacho-encaminhamento-opcao-a-e-c.docx) e anexar o [`00-cronograma-paralelo.docx`](00-cronograma-paralelo.docx).
2. Preencher estoque atual (em dias) na justificativa da emergência.
3. Coletar cotações e completar os mapas de preços dos dois processos.
4. Indicar **gestor** e **fiscal** nos TRs/contratos.
5. Tramitar o **Processo 1** com urgência (autorização → jurídico → PNCP → empenho/assinatura).
6. Em paralelo, publicar o **Processo 2** (edital revisado + divulgação ampliada + mínimo 8 dias úteis).

## Processo 1 — Dispensa emergencial

Pasta: [`processo-1-dispensa-emergencial/`](processo-1-dispensa-emergencial/)

1. Capa do processo  
2. DFD emergencial  
3. ETP simplificado  
4. Termo de Referência emergencial  
5. Justificativa da emergência  
6. Mapa de pesquisa de preços  
7. Razão da escolha do fornecedor  
8. Checklist de habilitação  
9. Minuta do contrato emergencial (com rescisão automática)  
10. Despacho de autorização + extrato PNCP  
11. Roteiro para parecer jurídico  

## Processo 2 — Pregão regular (paralelo)

Pasta: [`processo-2-pregao-regular/`](processo-2-pregao-regular/)

1. Capa do processo  
2. DFD revisado  
3. ETP completo  
4. Termo de Referência revisado (preço, prazo, habilitação, comodato, 12 meses)  
5. Pesquisa de preços atualizada  
6. Minuta de edital  
7. Anexo II — Minuta do contrato  
8. Anexo III — Modelo de proposta  
9. Ofício de divulgação ampliada  

## Anexos comuns

- [`anexos/anexo-locais-de-entrega-ubs-ubai.docx`](anexos/anexo-locais-de-entrega-ubs-ubai.docx) — 41 pontos urbanos + 12 rurais  

## Regenerar os .docx

```bash
python3 scripts/gerar_documentos.py
```

## Campos que ainda precisam ser preenchidos à mão

- Estoque atual em **dias**
- Valores unitários das **cotações** / mapa de preços
- Razão social / CNPJ da empresa da dispensa
- Nome do **gestor** e do **fiscal**
- Número do novo processo / empenho / publicação PNCP
- Data e horário da sessão do pregão
