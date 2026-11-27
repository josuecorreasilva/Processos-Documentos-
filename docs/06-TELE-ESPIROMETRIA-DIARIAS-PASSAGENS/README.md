# Projeto Tele-espirometria — DFD e ETP (Diárias e Passagens)

Pacote documental no padrão oficial SARH/SMS Pelotas (Lei Federal nº 14.133/2021, Decreto Municipal nº 6.799/2023 e Lei Municipal nº 6.621/2018) para instrução do **Processo SIS nº 15912/2026**.

## Peças geradas

| Ordem | Documento | Arquivo |
|------:|-----------|---------|
| 1 | **DFD** | `TeleEspirometria_DFD_Diarias_Passagens.docx` |
| 2 | **ETP** | `TeleEspirometria_ETP_Diarias_Passagens.docx` |

## Síntese do processo

- **Expediente:** Processo SIS nº 15912/2026 (abertura em 05/11/2026, acordo da Secretária em 09/11/2026)
- **Unidade Demandante:** Secretaria Municipal de Saúde – SMS / Diretoria de Atenção Primária / Departamento de Saúde Digital
- **Servidora Designada:** Elem Lira Rodrigues (Enfermeira, Matrícula 44777)
- **Evento:** Treinamento Prático Presencial — Projeto Tele-espirometria / TeleRESP (CTS-HC/UFMG + SAPS/MS + Novo PAC)
- **Data e Local:** 15/12/2026, 08h00 às 17h45 (credenciamento 30 minutos antes — 07h30) — Sede da Secretaria de Saúde do Estado de Pernambuco (SES-PE) — Rua Vinte e Quatro de Agosto, nº 209, bairro Santo Amaro, Recife/PE, CEP 50040-190
- **Deslocamento:** 14/12/2026 (06h00) a 16/12/2026 (18h00) — 60 horas de afastamento
- **Passagens (Lei 14.133/2021 — art. 75, II — Elemento 3.3.90.33):** estimativa preliminar de R$ 2.612,70 (aérea POA↔REC ref. R$ 2.350,00 + 2 trechos rodoviários Pelotas↔POA ref. R$ 262,70)
- **Diárias (Lei Municipal nº 6.621/2018 — Elemento 3.3.90.14):** 2 diárias integrais (8 URMs) + 1 diária estada (2 URMs) = **10 URMs não tributárias** (ref. mínima R$ 1.112,70 pela URM de R$ 111,27 — atualizar pela tabela SEFAZ de nov–dez/2026)

## Pendências operacionais antes do protocolo final

1. Atribuir numeração interna ao DFD (`nº [___]/2026`) e informar matrículas dos signatários da equipe de planejamento.
2. Confirmar Projeto/Atividade e Fonte de Recurso com a CCGOF/Contabilidade (elementos `3.3.90.33` e `3.3.90.14`).
3. Confirmar junto à SEFAZ/Contabilidade o valor vigente da **URM não tributária** de novembro/dezembro de 2026 para o cálculo definitivo das 10 URMs.
4. Juntar mapa de pesquisa formal de preços (mínimo 3 cotações contemporâneas — art. 23, Lei nº 14.133/2021) para emissão imediata das passagens.
5. Verificar disponibilidade de veículo oficial da SMS para o trecho Pelotas ↔ Porto Alegre (14/12 e 16/12/2026).
6. Providenciar conclusão do módulo EaD (até 08/12/2026), e-mail de confirmação de presença ao CTS-HC/UFMG e termo de cautela do espirômetro digital do Novo PAC.

## Regenerar

```bash
pip install python-docx
python3 scripts/gerar_tele_espirometria.py
```
