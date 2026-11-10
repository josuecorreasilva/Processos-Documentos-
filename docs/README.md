# Pacote documental oficial — Oxigênio Medicinal (SMS Pelotas)

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
