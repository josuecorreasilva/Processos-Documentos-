# Mapa Territorial — UBS Dunas (Distrito Areal • Pelotas/RS)

Mapa cartográfico em **formato A4 Paisagem (`297 × 210 mm`)** com **margens mínimas (`2,5 mm`)**, preenchido em toda a folha com **mapa cartográfico real em alta resolução** (quadras, edificações, áreas verdes e ruas do entorno) e rotulagem vetorial nítida de **todas as 47 ruas/avenidas do território da UBS Dunas** (43 internas + 4 limítrofes) mais **52 vias de referência ao redor**.

## Arquivos disponíveis

| Arquivo | Formato | Uso recomendado |
|---------|---------|-----------------|
| [`Mapa_Territorio_UBS_Dunas_A4.pdf`](./Mapa_Territorio_UBS_Dunas_A4.pdf) | **PDF A4 (300 DPI + texto vetorial)** — Escala `1 : 11,000` | **Impressão em cópias** para as equipes (colorida ou P&B) |
| [`Mapa_Territorio_UBS_Dunas_A4_300dpi.png`](./Mapa_Territorio_UBS_Dunas_A4_300dpi.png) | **PNG 300 DPI** (`3507 × 2480 px`) | Visualização em tela, celular/WhatsApp ou inserção em relatórios |
| [`ubs_dunas_territorio.geojson`](./ubs_dunas_territorio.geojson) | **GeoJSON (WGS 84)** | Uso em SIG (QGIS, Google Earth) |

## Orientações de impressão (A4)

1. Abra [`Mapa_Territorio_UBS_Dunas_A4.pdf`](./Mapa_Territorio_UBS_Dunas_A4.pdf).
2. Configure **Papel A4 — Orientação Paisagem** e **Tamanho Real (100%)** (ou *Ajustar às margens imprimíveis*).
3. O mapa foi projetado **todo claro e limpo** (sem áreas externas escurecidas), economizando toner/tinta e preservando excelente contraste tanto em impressão colorida quanto em cópias preto e branco.

## Bases utilizadas

- **Limite oficial do território e ponto da UBS Dunas:** *Mapa Território UBSs Pelotas* (SMS Pelotas / Google My Maps — `mid=1D6SDvi-BRD-GIhiKeRsM6oYAgty-jDQ`).
- **Mapa cartográfico base (sem rótulos):** Google Maps tiles `@2x` (zoom 17).
- **Toponímia viária e equipamentos públicos:** OpenStreetMap (ODbL).
- **Área oficial do território:** `1,32 km²` (`132,13 ha`) | **Perímetro:** `8.023 m`.

## Relação completa de logradouros do território da UBS Dunas (47)

| # | Logradouro | Classificação |
|---:|------------|---------------|
| 01 | Avenida 20 Dunas | Limítrofe |
| 02 | Avenida Domingos José de Almeida | Limítrofe |
| 03 | Avenida Doutor Roberto Frate Martins | Limítrofe |
| 04 | Avenida Engenheiro Ildefonso Simões Lopes | Limítrofe |
| 05 | Avenida Manoel Antônio Peres | Interna |
| 06 | Avenida Ulysses Silveira Guimarães | Interna |
| 07 | Passeio Dois | Interna |
| 08 | Passeio Dois Loteamento Dunas | Interna |
| 09 | Passeio Três Loteamento Dunas | Interna |
| 10 | Passeio Um | Interna |
| 11 | Rua Alzira Gularte Farias | Interna |
| 12 | Rua Cinco | Interna |
| 13 | Rua Cláudio Joaquim Piquet Coelho | Interna |
| 14 | Rua Dez | Interna |
| 15 | Rua Dezesseis | Interna |
| 16 | Rua Dezessete | Interna |
| 17 | Rua Dezoito | Interna |
| 18 | Rua Dois | Interna |
| 19 | Rua Doze | Interna |
| 20 | Rua Ilma Araújo Chagas | Interna |
| 21 | Rua João Baptista Marimon | Interna |
| 22 | Rua Jornalista Guerreiro Vitória | Interna |
| 23 | Rua Maria José Campos Domingues | Interna |
| 24 | Rua Miss Yolanda Pereira | Interna |
| 25 | Rua Nove | Interna |
| 26 | Rua Oito | Interna |
| 27 | Rua Onze | Interna |
| 28 | Rua Quatorze | Interna |
| 29 | Rua Quatro | Interna |
| 30 | Rua Quatro Lot Bom Jesus | Interna |
| 31 | Rua Quinze | Interna |
| 32 | Rua Seis | Interna |
| 33 | Rua Sete | Interna |
| 34 | Rua Tito Borges | Interna |
| 35 | Rua Três | Interna |
| 36 | Rua Treze | Interna |
| 37 | Rua Um | Interna |
| 38 | Rua Vereador Élbio Abreu | Interna |
| 39 | Rua Vinte e Cinco | Interna |
| 40 | Rua Vinte e Dois | Interna |
| 41 | Rua Vinte e Oito | Interna |
| 42 | Rua Vinte e Quatro | Interna |
| 43 | Rua Vinte e Seis | Interna |
| 44 | Rua Vinte e Sete | Interna |
| 45 | Rua Vinte e Três | Interna |
| 46 | Rua Vinte e Um | Interna |
| 47 | Travessa Manoel Antônio Peres | Interna |

## Como regenerar

```bash
pip install shapely pyproj matplotlib pillow requests
python3 scripts/gerar_mapa_ubs_dunas.py
# Para atualizar dados online (KML/OSM):
python3 scripts/gerar_mapa_ubs_dunas.py --refresh
```
