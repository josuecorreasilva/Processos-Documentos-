# Processos-Documentos — Oxigênio Medicinal (SMS Pelotas)

Pacote documental para andamento **simultâneo** de:

1. **Dispensa emergencial** (Art. 75, VIII) — abastecimento imediato  
2. **Pregão eletrônico** — solução definitiva em paralelo  

Base: Processo nº **1242/2026** (licitação deserta), DFD O2 atualizado (2026) e Projeto O2.

## Início rápido

Abra [`docs/README.md`](docs/README.md) para o índice completo e a ordem de tramitação.

Principais peças de abertura:

- [`docs/00-despacho-encaminhamento-opcao-a-e-c.docx`](docs/00-despacho-encaminhamento-opcao-a-e-c.docx)
- [`docs/00-cronograma-paralelo.docx`](docs/00-cronograma-paralelo.docx)
- [`docs/processo-1-dispensa-emergencial/`](docs/processo-1-dispensa-emergencial/)
- [`docs/processo-2-pregao-regular/`](docs/processo-2-pregao-regular/)

## Regenerar documentos

```bash
pip install python-docx
python3 scripts/gerar_documentos.py
```
