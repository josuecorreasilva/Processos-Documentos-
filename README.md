# Queremos Viver — PET-Saúde Clima UCPel

Site institucional do projeto **Queremos Viver: Articulação Interprofissional para Adaptação e Resiliência do SUS frente às Emergências Climáticas em Pelotas-RS** (PET-Saúde: Clima, UCPel / SMS Pelotas).

## Páginas

- Início
- O projeto e os três eixos
- Equipe e **A equipe em números**
- Documentos, notícias, agenda e sobre

## Desenvolvimento

As páginas HTML são geradas por `tools/build.py`.

```bash
python3 tools/build.py
python3 -m http.server 8080
```

Abra `http://127.0.0.1:8080`.

## Publicação (GitHub Pages)

No repositório: **Settings → Pages → Deploy from a branch**, branch `main`, pasta `/ (root)`.

## Fontes dos números

Os totais de “A equipe em números” usam o Edital nº 044/2026 – PRAC/UCPel e o resultado final da seleção (08/07/2026), documentos públicos da UCPel.
