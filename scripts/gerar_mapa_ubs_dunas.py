#!/usr/bin/env python3
"""
Gera o Mapa Territorial Oficial da UBS Dunas (Distrito Areal — Pelotas/RS)
em formato A4 Paisagem pronto para impressão (margens mínimas de 2,5 mm),
com mapa cartográfico real de fundo (quadras, edificações, áreas verdes e ruas
em alta resolução, sem rótulos duplicados) preenchendo toda a folha A4 e
rotulagem vetorial nítida de todas as ruas do território e das principais vias
de referência do entorno.

Bases utilizadas:
  - Polígono oficial e ponto da UBS Dunas: Mapa Território UBSs Pelotas (SMS / Google My Maps)
    https://www.google.com/maps/d/edit?mid=1D6SDvi-BRD-GIhiKeRsM6oYAgty-jDQ
  - Mapa base cartográfico (sem rótulos): Google Maps tiles @2x (z17)
  - Malha viária, toponímia e equipamentos públicos: OpenStreetMap (ODbL)
"""

from __future__ import annotations

import argparse
import concurrent.futures
import io
import json
import math
import re
import time
import unicodedata
import xml.etree.ElementTree as ET
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patheffects as pe
from matplotlib.offsetbox import AnnotationBbox, OffsetImage
from matplotlib.patches import Circle, FancyBboxPatch, Polygon as MplPolygon, Rectangle
import numpy as np
from PIL import Image
from pyproj import Transformer
import requests
from shapely.affinity import rotate as shp_rotate
from shapely.geometry import LineString, MultiLineString, Point, Polygon, box, mapping
from shapely.ops import linemerge

REPO_ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = REPO_ROOT / "docs" / "07-MAPA-TERRITORIO-UBS-DUNAS"
BRASAO_PATH = REPO_ROOT / "docs" / "assets" / "brasao_pelotas.png"
DATA_CACHE = OUT_DIR / "ubs_dunas_dados.json"
GEOJSON_OUT = OUT_DIR / "ubs_dunas_territorio.geojson"
PDF_OUT = OUT_DIR / "Mapa_Territorio_UBS_Dunas_A4.pdf"
PNG_OUT = OUT_DIR / "Mapa_Territorio_UBS_Dunas_A4_300dpi.png"

KML_URL = "https://www.google.com/maps/d/kml?mid=1D6SDvi-BRD-GIhiKeRsM6oYAgty-jDQ&forcekml=1"
TMP_KML = Path("/tmp/ubs_pelotas.kml")
TMP_OSM_FULL = Path("/tmp/ubs_dunas_full_osm.json")
TILE_CACHE_DIR = Path("/tmp/ubs_dunas_tiles_z17")

TO_3857 = Transformer.from_crs("EPSG:4326", "EPSG:3857", always_xy=True)
TO_WGS = Transformer.from_crs("EPSG:3857", "EPSG:4326", always_xy=True)
TO_UTM = Transformer.from_crs("EPSG:4326", "EPSG:31982", always_xy=True)

# Fator de escala Web Mercator em lat = -31.74° (1 m real = 1 / cos(lat) unidades 3857)
LAT_REF = -31.7418
MERC_PER_METER = 1.0 / math.cos(math.radians(LAT_REF))

# ---------------------------------------------------------------------------
# 1. Utilitários de texto e projeção de tiles Web Mercator
# ---------------------------------------------------------------------------

def sort_key_pt(s: str) -> str:
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower()


def abbreviate_street(name: str) -> str:
    """Abreviação cartográfica padrão IBGE para legibilidade limpa na folha A4."""
    s = name.strip()
    replacements = [
        (r"^Avenida Engenheiro ", "Av. Eng. "),
        (r"^Avenida Doutor ", "Av. Dr. "),
        (r"^Avenida Prefeito ", "Av. Pref. "),
        (r"^Avenida ", "Av. "),
        (r"^Travessa ", "Tv. "),
        (r"^Passeio ", "Pass. "),
        (r"^Rua Vereador ", "R. Ver. "),
        (r"^Rua Jornalista ", "R. Jorn. "),
        (r"^Rua Marechal ", "R. Mal. "),
        (r"^Rua Monsenhor ", "R. Mons. "),
        (r"^Rua Reverendo ", "R. Rev. "),
        (r"^Rua Professor ", "R. Prof. "),
        (r"^Rua Doutor ", "R. Dr. "),
        (r"^Rua ", "R. "),
    ]
    for pat, rep in replacements:
        if re.match(pat, s):
            s = re.sub(pat, rep, s, count=1)
            break
    s = s.replace(" Loteamento Dunas", " (Lot. Dunas)")
    s = s.replace(" Lot Bom Jesus", " (Lot. Bom Jesus)")
    s = s.replace(" Lot Bom Jésus", " (Lot. Bom Jesus)")
    return s


def wrap_two_lines(text: str) -> str:
    words = text.split()
    if len(words) <= 2 or len(text) <= 14:
        return text
    best_split, best_diff = 1, 999
    for i in range(1, len(words)):
        diff = abs(len(" ".join(words[:i])) - len(" ".join(words[i:])))
        if diff < best_diff:
            best_diff, best_split = diff, i
    return " ".join(words[:best_split]) + "\n" + " ".join(words[best_split:])


def deg2tile(lat: float, lon: float, z: int) -> tuple[int, int]:
    n = 2 ** z
    xt = int((lon + 180.0) / 360.0 * n)
    yt = int((1.0 - math.asinh(math.tan(math.radians(lat))) / math.pi) / 2.0 * n)
    return xt, yt


def tile2lonlat(xt: int, yt: int, z: int) -> tuple[float, float]:
    n = 2 ** z
    lon = xt / n * 360.0 - 180.0
    lat = math.degrees(math.atan(math.sinh(math.pi * (1.0 - 2.0 * yt / n))))
    return lon, lat


def fetch_basemap_mosaic(wx0: float, wy0: float, wx1: float, wy1: float, zoom: int = 17) -> tuple[np.ndarray, tuple[float, float, float, float]]:
    """
    Baixa e costura tiles Google Maps sem rótulos (@2x, 512x512 px) cobrindo
    [wx0, wx1] x [wy0, wy1] em EPSG:3857.
    """
    TILE_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    lon_min, lat_min = TO_WGS.transform(wx0, wy0)
    lon_max, lat_max = TO_WGS.transform(wx1, wy1)

    x_min, y_min = deg2tile(lat_max, lon_min, zoom)  # topo-esquerda
    x_max, y_max = deg2tile(lat_min, lon_max, zoom)  # base-direita

    cols = x_max - x_min + 1
    rows = y_max - y_min + 1
    total_tiles = cols * rows
    print(f"Basemap mosaic z={zoom} @2x: {cols}x{rows} = {total_tiles} tiles...")

    hdrs = {
        "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0.0.0 Safari/537.36",
        "Accept": "image/png,image/*;q=0.8,*/*;q=0.5",
    }

    def get_one(xy: tuple[int, int]) -> tuple[int, int, Image.Image]:
        xt, yt = xy
        cache_f = TILE_CACHE_DIR / f"g_{zoom}_{xt}_{yt}.png"
        if cache_f.exists() and cache_f.stat().st_size > 200:
            return xt, yt, Image.open(cache_f).convert("RGB")
        sub = (xt + yt) % 4
        url = f"https://mt{sub}.google.com/vt/lyrs=m&x={xt}&y={yt}&z={zoom}&scale=2&apistyle=s.e:l|p.v:off"
        last_err = None
        for attempt in range(4):
            try:
                r = requests.get(url, timeout=20, headers=hdrs)
                if r.status_code == 200 and len(r.content) > 200:
                    im = Image.open(io.BytesIO(r.content)).convert("RGB")
                    im.save(cache_f)
                    return xt, yt, im
                last_err = f"HTTP {r.status_code} ({len(r.content)}B)"
            except Exception as exc:
                last_err = str(exc)
            time.sleep(0.5 * (attempt + 1))
        raise RuntimeError(f"Falha ao baixar tile ({xt},{yt}) z={zoom}: {last_err}")

    tile_list = [(xt, yt) for yt in range(y_min, y_max + 1) for xt in range(x_min, x_max + 1)]
    tile_sz = 512
    mosaic = Image.new("RGB", (cols * tile_sz, rows * tile_sz), (245, 245, 245))

    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as ex:
        for xt, yt, im in ex.map(get_one, tile_list):
            if im.size != (tile_sz, tile_sz):
                im = im.resize((tile_sz, tile_sz), Image.Resampling.LANCZOS)
            mosaic.paste(im, ((xt - x_min) * tile_sz, (yt - y_min) * tile_sz))

    # Extensão exata em EPSG:3857 do mosaico costurado
    lon_l, lat_t = tile2lonlat(x_min, y_min, zoom)
    lon_r, lat_b = tile2lonlat(x_max + 1, y_max + 1, zoom)
    mx0, my1 = TO_3857.transform(lon_l, lat_t)
    mx1, my0 = TO_3857.transform(lon_r, lat_b)

    # Recortar exatamente para a janela [wx0, wx1] x [wy0, wy1] em resolução ótima (~3600 px de largura)
    mw, mh = mosaic.size
    u0 = max(0.0, (wx0 - mx0) / (mx1 - mx0))
    u1 = min(1.0, (wx1 - mx0) / (mx1 - mx0))
    v0 = max(0.0, (my1 - wy1) / (my1 - my0))
    v1 = min(1.0, (my1 - wy0) / (my1 - my0))

    crop_box = (int(round(u0 * mw)), int(round(v0 * mh)), int(round(u1 * mw)), int(round(v1 * mh)))
    cropped = mosaic.crop(crop_box)
    # Suavizar para 3200 px de largura (~280 DPI no raster de fundo + textos 100% vetoriais no PDF)
    target_w = 3200
    target_h = int(round(target_w * cropped.height / cropped.width))
    if cropped.width > target_w:
        cropped = cropped.resize((target_w, target_h), Image.Resampling.LANCZOS)

    print(f"Basemap cropped: {cropped.size[0]}x{cropped.size[1]} px")
    return np.asarray(cropped), (wx0, wx1, wy0, wy1)


# ---------------------------------------------------------------------------
# 2. Extração e preparação dos dados (KML + OSM da janela completa A4)
# ---------------------------------------------------------------------------

def load_or_build_dataset(refresh: bool = False) -> dict:
    if DATA_CACHE.exists() and not refresh:
        with open(DATA_CACHE, encoding="utf-8") as f:
            cached = json.load(f)
        if cached.get("version") == 2:
            return cached

    if refresh or not TMP_KML.exists():
        r = requests.get(KML_URL, timeout=60, headers={"User-Agent": "SMSPelotas-MapaUBSDunas/2.0"})
        r.raise_for_status()
        TMP_KML.write_bytes(r.content)

    ns = {"kml": "http://www.opengis.net/kml/2.2"}
    tree = ET.parse(TMP_KML)
    root = tree.getroot()

    poly_wgs = None
    ubs_pt_wgs = None
    for pm in root.findall(".//kml:Placemark", ns):
        nm = (pm.findtext("kml:name", "", ns) or "").strip()
        if nm == "UBS Dunas" and pm.find(".//kml:Polygon", ns) is not None:
            txt = pm.find(".//kml:Polygon//kml:coordinates", ns).text.strip()
            poly_wgs = [[float(p.split(",")[0]), float(p.split(",")[1])] for p in txt.split()]
        if "duna" in nm.lower() and pm.find(".//kml:Point", ns) is not None:
            p = pm.find(".//kml:Point//kml:coordinates", ns).text.strip().split(",")
            ubs_pt_wgs = [float(p[0]), float(p[1])]

    assert poly_wgs and ubs_pt_wgs, "Não foi possível localizar UBS Dunas no KML."

    if refresh or not TMP_OSM_FULL.exists():
        bbox = "-52.3315,-31.7545,-52.2915,-31.7295"
        url = f"https://api.openstreetmap.org/api/0.6/map.json?bbox={bbox}"
        r = requests.get(url, timeout=60, headers={"User-Agent": "SMSPelotas-MapaUBSDunas/2.0"})
        r.raise_for_status()
        TMP_OSM_FULL.write_bytes(r.content)

    with open(TMP_OSM_FULL, encoding="utf-8") as f:
        osm = json.load(f)

    nodes = {e["id"]: (e["lon"], e["lat"]) for e in osm["elements"] if e["type"] == "node"}
    poly_utm = Polygon([TO_UTM.transform(x, y) for x, y in poly_wgs])
    poly_utm_border_18 = poly_utm.buffer(18.0)

    poly_3857 = Polygon([TO_3857.transform(x, y) for x, y in poly_wgs])
    ubs_3857 = list(TO_3857.transform(*ubs_pt_wgs))

    territory_roads = []
    surrounding_roads = []
    pois = []

    for e in osm["elements"]:
        tags = e.get("tags", {})
        if e["type"] != "way":
            continue
        pts_wgs = [nodes[nid] for nid in e.get("nodes", []) if nid in nodes]
        if len(pts_wgs) < 2:
            continue

        hw = tags.get("highway")
        if hw:
            pts_utm = [TO_UTM.transform(x, y) for x, y in pts_wgs]
            pts_3857 = [list(TO_3857.transform(x, y)) for x, y in pts_wgs]
            line_utm = LineString(pts_utm)
            inter_in = line_utm.intersection(poly_utm)
            inter_b = line_utm.intersection(poly_utm_border_18)
            is_internal = inter_in.length > 5.0
            is_border = (not is_internal) and (inter_b.length > 80.0 and line_utm.distance(poly_utm) < 15.0)
            nm = tags.get("name", "").strip()

            if is_internal or is_border:
                if nm:
                    # Recortar suavemente para manter o trecho no território + borda
                    line_3857 = LineString(pts_3857)
                    clipped = line_3857.intersection(poly_3857.buffer(28.0 * MERC_PER_METER))
                    segs = []
                    if isinstance(clipped, LineString):
                        segs = [list(clipped.coords)]
                    elif isinstance(clipped, MultiLineString):
                        segs = [list(g.coords) for g in clipped.geoms if g.length > 3.0]
                    elif hasattr(clipped, "geoms"):
                        segs = [list(g.coords) for g in clipped.geoms if isinstance(g, LineString) and g.length > 3.0]
                    if segs:
                        territory_roads.append({
                            "id": e["id"],
                            "name": nm,
                            "highway": hw,
                            "is_border": is_border,
                            "segments_3857": segs,
                        })
            elif nm and hw in ("trunk", "primary", "secondary", "tertiary", "residential", "unclassified"):
                surrounding_roads.append({
                    "id": e["id"],
                    "name": nm,
                    "highway": hw,
                    "segments_3857": [pts_3857],
                })
            continue

        am = tags.get("amenity", "")
        nm = tags.get("name", "").strip()
        if am in ("school", "social_centre", "community_centre", "hospital") and nm:
            pts_utm = [TO_UTM.transform(x, y) for x, y in pts_wgs]
            c_utm = LineString(pts_utm).centroid
            if poly_utm_border_18.contains(c_utm):
                pts_3857 = [TO_3857.transform(x, y) for x, y in pts_wgs]
                c_3857 = LineString(pts_3857).centroid
                pois.append({"name": nm, "amenity": am, "x": c_3857.x, "y": c_3857.y})

    # Deduplicar POIs por nome
    pois_dedup = {}
    for p in pois:
        rec = pois_dedup.setdefault(p["name"], {"amenity": p["amenity"], "xs": [], "ys": []})
        rec["xs"].append(p["x"])
        rec["ys"].append(p["y"])
    pois_final = [
        {"name": nm, "amenity": r["amenity"], "x": float(np.mean(r["xs"])), "y": float(np.mean(r["ys"]))}
        for nm, r in pois_dedup.items()
    ]

    dataset = {
        "version": 2,
        "source_mymaps": "https://www.google.com/maps/d/edit?mid=1D6SDvi-BRD-GIhiKeRsM6oYAgty-jDQ",
        "area_km2": round(poly_utm.area / 1e6, 4),
        "perimeter_m": round(poly_utm.length, 1),
        "polygon_wgs84": poly_wgs,
        "polygon_3857": [list(c) for c in poly_3857.exterior.coords],
        "ubs_point_wgs84": ubs_pt_wgs,
        "ubs_point_3857": ubs_3857,
        "territory_roads": territory_roads,
        "surrounding_roads": surrounding_roads,
        "pois": pois_final,
    }

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    with open(DATA_CACHE, "w", encoding="utf-8") as f:
        json.dump(dataset, f, ensure_ascii=False)

    # Atualizar também GeoJSON auditável em WGS84
    features = [
        {
            "type": "Feature",
            "properties": {
                "nome": "Território UBS Dunas",
                "distrito": "Distrito Areal",
                "municipio": "Pelotas/RS",
                "area_km2": dataset["area_km2"],
                "perimetro_m": dataset["perimeter_m"],
            },
            "geometry": mapping(Polygon(poly_wgs)),
        },
        {
            "type": "Feature",
            "properties": {"nome": "UBS Dunas (Distrito Areal)", "tipo": "UBS"},
            "geometry": mapping(Point(ubs_pt_wgs)),
        },
    ]
    for r in territory_roads:
        for seg in r["segments_3857"]:
            seg_wgs = [list(TO_WGS.transform(x, y)) for x, y in seg]
            features.append({
                "type": "Feature",
                "properties": {
                    "nome": r["name"],
                    "highway": r["highway"],
                    "limitrofe": r["is_border"],
                },
                "geometry": mapping(LineString(seg_wgs)),
            })
    with open(GEOJSON_OUT, "w", encoding="utf-8") as f:
        json.dump({"type": "FeatureCollection", "features": features}, f, ensure_ascii=False, indent=2)

    return dataset


# ---------------------------------------------------------------------------
# 3. Motor de rotulagem vetorial (zero duplicação + zero sobreposição)
# ---------------------------------------------------------------------------

def merge_roads_by_name(roads: list[dict], clip_box: Polygon | None = None) -> dict[str, dict]:
    grouped: dict[str, dict] = {}
    for r in roads:
        nm = r["name"]
        if not nm:
            continue
        for seg in r["segments_3857"]:
            if len(seg) < 2:
                continue
            ln = LineString(seg)
            if clip_box is not None:
                ln = ln.intersection(clip_box)
                if ln.is_empty:
                    continue
            parts_in = [ln] if isinstance(ln, LineString) else [g for g in getattr(ln, "geoms", []) if isinstance(g, LineString) and g.length > 5.0]
            if not parts_in:
                continue
            entry = grouped.setdefault(nm, {"lines": [], "hw": set(), "is_border": r.get("is_border", False)})
            entry["hw"].add(r["highway"])
            if not r.get("is_border", False):
                entry["is_border"] = False
            entry["lines"].extend(parts_in)

    for nm, entry in grouped.items():
        merged = linemerge(entry["lines"])
        if isinstance(merged, LineString):
            parts = [merged]
        elif isinstance(merged, MultiLineString):
            parts = sorted(list(merged.geoms), key=lambda g: g.length, reverse=True)
        else:
            parts = sorted(entry["lines"], key=lambda g: g.length, reverse=True)
        entry["merged"] = parts
        entry["total_len_m"] = sum(g.length for g in parts) / MERC_PER_METER
    return grouped


def point_and_angle_along(line: LineString, frac: float, span_units: float = 40.0) -> tuple[float, float, float]:
    L = line.length
    d = max(0.0, min(L, L * frac))
    p = line.interpolate(d)
    d0 = max(0.0, d - span_units / 2.0)
    d1 = min(L, d + span_units / 2.0)
    if d1 - d0 < 1.0:
        p0, p1 = Point(line.coords[0]), Point(line.coords[-1])
    else:
        p0, p1 = line.interpolate(d0), line.interpolate(d1)
    ang = math.degrees(math.atan2(p1.y - p0.y, p1.x - p0.x))
    if ang > 90.0:
        ang -= 180.0
    elif ang < -90.0:
        ang += 180.0
    return p.x, p.y, ang


def make_label_box_mm(x_mm: float, y_mm: float, text: str, fontsize_pt: float, angle_deg: float, pad_x: float = 0.48, pad_y: float = 0.14) -> Polygon:
    lines = text.split("\n")
    max_chars = max(len(ln) for ln in lines)
    char_w_mm = fontsize_pt * 0.35278 * 0.50
    text_w = max(3.3, max_chars * char_w_mm) + 2 * pad_x
    text_h = fontsize_pt * 0.35278 * (1.05 * len(lines) + 0.10 * (len(lines) - 1)) + 2 * pad_y
    b = box(x_mm - text_w / 2.0, y_mm - text_h / 2.0, x_mm + text_w / 2.0, y_mm + text_h / 2.0)
    return shp_rotate(b, angle_deg, origin=(x_mm, y_mm))


def territory_street_tier(info: dict) -> tuple[int, float]:
    hws = info["hw"]
    if info["is_border"] and "secondary" in hws:
        return (1, 8.2)
    if "secondary" in hws:
        return (1, 8.2)
    if info["is_border"]:
        return (1, 7.6)
    if "tertiary" in hws:
        return (2, 7.1)
    return (3, 6.35)


def surrounding_street_tier(info: dict) -> tuple[int, float]:
    hws = info["hw"]
    if "trunk" in hws or "primary" in hws or "secondary" in hws:
        return (1, 7.4)
    if "tertiary" in hws:
        return (2, 6.5)
    return (3, 5.7)


# ---------------------------------------------------------------------------
# 4. Renderização A4 Paisagem — Mapa Cartográfico Real Completo + Margens Mínimas
# ---------------------------------------------------------------------------

def render_map(dataset: dict) -> dict:
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = ["Liberation Sans", "DejaVu Sans", "Noto Sans"]
    plt.rcParams["pdf.fonttype"] = 42
    plt.rcParams["ps.fonttype"] = 42

    poly_3857 = Polygon(dataset["polygon_3857"])
    ux, uy = dataset["ubs_point_3857"]
    minx, miny, maxx, maxy = poly_3857.bounds

    # Folha A4 Paisagem exata (297 x 210 mm) com margens mínimas de 2,5 mm!
    PAGE_W_MM, PAGE_H_MM = 297.0, 210.0
    MARGIN_MM = 2.5
    MAP_W_MM = PAGE_W_MM - 2 * MARGIN_MM  # 292.0 mm
    MAP_H_MM = PAGE_H_MM - 2 * MARGIN_MM  # 205.0 mm
    MM_TO_IN = 1.0 / 25.4

    # Escala real máxima com Norte para cima: 1 : 11.000 (1 cm = 110 m reais)
    SCALE = 11000.0
    WORLD_W_M = (MAP_W_MM / 1000.0) * SCALE  # 3.212,0 m reais
    WORLD_H_M = (MAP_H_MM / 1000.0) * SCALE  # 2.255,0 m reais
    WORLD_W_3857 = WORLD_W_M * MERC_PER_METER
    WORLD_H_3857 = WORLD_H_M * MERC_PER_METER

    # Centralizar o território da UBS Dunas confortavelmente na folha A4
    cx = (minx + maxx) / 2.0
    cy = (miny + maxy) / 2.0
    WX0 = cx - WORLD_W_3857 / 2.0
    WX1 = cx + WORLD_W_3857 / 2.0
    WY0 = cy - WORLD_H_3857 / 2.0
    WY1 = cy + WORLD_H_3857 / 2.0

    def m3857_to_mm(x: float, y: float) -> tuple[float, float]:
        return ((x - WX0) / WORLD_W_3857 * MAP_W_MM, (y - WY0) / WORLD_H_3857 * MAP_H_MM)

    def mm_to_3857(xm: float, ym: float) -> tuple[float, float]:
        return (WX0 + (xm / MAP_W_MM) * WORLD_W_3857, WY0 + (ym / MAP_H_MM) * WORLD_H_3857)

    # 1. Baixar/carregar o mosaico de tiles cartográficos reais (Google Maps sem rótulos, z17 @2x)
    basemap_img, basemap_extent = fetch_basemap_mosaic(WX0, WY0, WX1, WY1, zoom=17)

    fig = plt.figure(figsize=(PAGE_W_MM * MM_TO_IN, PAGE_H_MM * MM_TO_IN), dpi=300, facecolor="#FFFFFF")
    ax = fig.add_axes([
        MARGIN_MM / PAGE_W_MM,
        MARGIN_MM / PAGE_H_MM,
        MAP_W_MM / PAGE_W_MM,
        MAP_H_MM / PAGE_H_MM,
    ])
    ax.set_xlim(WX0, WX1)
    ax.set_ylim(WY0, WY1)
    ax.set_aspect("equal", adjustable="box")
    ax.set_xticks([])
    ax.set_yticks([])
    for sp in ax.spines.values():
        sp.set_edgecolor("#334155")
        sp.set_linewidth(0.8)

    # Camada 1: Mapa base cartográfico real em toda a folha A4 (bem claro e nítido)
    ax.imshow(basemap_img, extent=basemap_extent, origin="upper", interpolation="lanczos", zorder=1)

    # Camada 2: Suavização BEM CLARA fora do polígono (mantém todo o entorno claro e legível,
    # sem escurecer nem gastar toner!) + destaque quente suave dentro da UBS Dunas
    outer_rect = box(WX0 - 500, WY0 - 500, WX1 + 500, WY1 + 500)
    outside_geom = outer_rect.difference(poly_3857)
    # Desenhar máscara externa usando Path com buraco (polígono externo - polígono UBS Dunas)
    from matplotlib.path import Path as MplPath
    from matplotlib.patches import PathPatch

    ext_coords = list(outer_rect.exterior.coords)
    hole_coords = list(poly_3857.exterior.coords)[::-1]  # orientação oposta para furo
    verts = ext_coords + hole_coords
    codes = (
        [MplPath.MOVETO] + [MplPath.LINETO] * (len(ext_coords) - 2) + [MplPath.CLOSEPOLY]
        + [MplPath.MOVETO] + [MplPath.LINETO] * (len(hole_coords) - 2) + [MplPath.CLOSEPOLY]
    )
    # Véu branco claríssimo (alpha=0.22): suaviza o entorno SEM escurecer nada!
    ax.add_patch(PathPatch(MplPath(verts, codes), facecolor="#FFFFFF", edgecolor="none", alpha=0.24, zorder=2))

    # Destaque suave dentro do território da UBS Dunas (preserva quadras, prédios e áreas verdes)
    px, py = poly_3857.exterior.xy
    ax.fill(px, py, facecolor="#FEF08A", edgecolor="none", alpha=0.16, zorder=3)

    # Camada 3: Contorno Oficial do Território da UBS Dunas (muito nítido, alto contraste)
    ax.plot(px, py, color="#FFFFFF", lw=4.6, solid_joinstyle="round", alpha=0.92, zorder=8)
    ax.plot(px, py, color="#DC2626", lw=2.5, solid_joinstyle="round", zorder=9)

    # -----------------------------------------------------------------------
    # Camada 4: Cartela Ultra Compacta no canto SW (sem tabela lateral!)
    # -----------------------------------------------------------------------
    CARD_W_MM, CARD_H_MM = 84.0, 35.0
    CARD_X_MM, CARD_Y_MM = 2.5, 2.5

    occupied_boxes: list[Polygon] = [
        box(CARD_X_MM - 1.8, CARD_Y_MM - 1.8, CARD_X_MM + CARD_W_MM + 2.0, CARD_Y_MM + CARD_H_MM + 2.0),
    ]

    # -----------------------------------------------------------------------
    # Camada 5: Marcador UBS Dunas e Principais Equipamentos Públicos
    # -----------------------------------------------------------------------
    ux_mm, uy_mm = m3857_to_mm(ux, uy)
    ax.scatter([ux], [uy], s=175, c="#DC2626", edgecolors="#FFFFFF", linewidths=2.0, zorder=22)
    ax.text(ux, uy, "+", ha="center", va="center", fontsize=10.5, fontweight="bold", color="#FFFFFF", zorder=23)

    # Badge UBS DUNAS posicionado com seta curta em vão aberto próximo (sem cobrir ruas)
    ubs_tx = ux - 210 * MERC_PER_METER
    ubs_ty = uy - 125 * MERC_PER_METER
    ax.annotate(
        "UBS DUNAS",
        xy=(ux, uy),
        xytext=(ubs_tx, ubs_ty),
        ha="center", va="center",
        fontsize=8.2, fontweight="bold", color="#991B1B",
        bbox=dict(boxstyle="round,pad=0.28", facecolor="#FEF2F2", edgecolor="#DC2626", lw=1.3),
        arrowprops=dict(arrowstyle="-|>", color="#DC2626", lw=1.3, shrinkA=2, shrinkB=7, mutation_scale=10),
        zorder=24,
    )
    ubs_bx_mm, ubs_by_mm = m3857_to_mm(ubs_tx, ubs_ty)
    occupied_boxes.append(box(ux_mm - 3.6, uy_mm - 3.6, ux_mm + 3.6, uy_mm + 3.6))
    occupied_boxes.append(box(ubs_bx_mm - 12.2, ubs_by_mm - 3.3, ubs_bx_mm + 12.2, ubs_by_mm + 3.3))

    # Principais equipamentos públicos no território
    poi_spec = {
        "EMEF Núcleo Habitacional Dunas": ("EMEF Núc. Hab. Dunas", "#1D4ED8", (0, -34)),
        "EMEF Jornalista Deogar Soares": ("EMEF Deogar Soares", "#1D4ED8", (-78, -28)),
        "CRAS - Centro de Referência de Assistência Social": ("CRAS Dunas", "#047857", (58, 26)),
        "Dunas Clube": ("Dunas Clube", "#6D28D9", (0, 36)),
    }
    for p in dataset["pois"]:
        if p["name"] not in poi_spec:
            continue
        lbl, col, (dx_m, dy_m) = poi_spec[p["name"]]
        ax.scatter([p["x"]], [p["y"]], s=32, marker="s", c=col, edgecolors="#FFFFFF", linewidths=0.9, zorder=19)
        tx = p["x"] + dx_m * MERC_PER_METER
        ty = p["y"] + dy_m * MERC_PER_METER
        ax.text(
            tx, ty, lbl,
            ha="center", va="center",
            fontsize=5.8, fontweight="bold", color=col,
            path_effects=[pe.withStroke(linewidth=2.4, foreground="#FFFFFF")],
            zorder=20,
        )
        pxm, pym = m3857_to_mm(tx, ty)
        occupied_boxes.append(make_label_box_mm(pxm, pym, lbl, 5.8, 0.0, pad_x=0.6, pad_y=0.35))

    # -----------------------------------------------------------------------
    # Camada 6A: TODAS as 47 ruas do território da UBS Dunas (prioridade absoluta!)
    # -----------------------------------------------------------------------
    view_box = box(WX0, WY0, WX1, WY1)
    terr_streets = merge_roads_by_name(dataset["territory_roads"], clip_box=view_box)

    stagger_a = {
        "Rua Seis", "Rua Oito", "Rua Dez", "Rua Doze", "Rua Quatorze", "Rua Dezesseis", "Rua Dezoito",
        "Rua Vinte e Um", "Rua Vinte e Três", "Rua Vinte e Cinco", "Rua Vinte e Sete",
    }
    stagger_b = {
        "Rua Sete", "Rua Nove", "Rua Onze", "Rua Treze", "Rua Quinze", "Rua Dezessete",
        "Rua Vinte e Dois", "Rua Vinte e Quatro", "Rua Vinte e Seis", "Rua Vinte e Oito",
    }

    ordered_terr = sorted(
        terr_streets.keys(),
        key=lambda nm: (
            territory_street_tier(terr_streets[nm])[0],
            -terr_streets[nm]["total_len_m"] if territory_street_tier(terr_streets[nm])[0] <= 2 else terr_streets[nm]["total_len_m"],
            sort_key_pt(nm),
        ),
    )

    def fits_in_map(lbox: Polygon, margin: float = 1.4) -> bool:
        bx0, by0, bx1, by1 = lbox.bounds
        return bx0 >= margin and by0 >= margin and bx1 <= MAP_W_MM - margin and by1 <= MAP_H_MM - margin

    terr_placed = []
    for nm in ordered_terr:
        info = terr_streets[nm]
        tier, base_fs = territory_street_tier(info)
        parts = info["merged"]
        main_line = parts[0]
        short_lbl = abbreviate_street(nm)

        if nm in ("Rua Vereador Élbio Abreu", "Rua João Baptista Marimon", "Rua Tito Borges", "Passeio Dois Loteamento Dunas", "Passeio Dois"):
            fracs = [0.22, 0.78, 0.32, 0.68, 0.50, 0.14, 0.86, 0.42, 0.58]
        elif nm == "Avenida Engenheiro Ildefonso Simões Lopes":
            fracs = [0.48, 0.58, 0.38, 0.68, 0.28, 0.78]
        elif nm in stagger_a:
            fracs = [0.36, 0.44, 0.28, 0.56, 0.68, 0.20, 0.78]
        elif nm in stagger_b:
            fracs = [0.64, 0.56, 0.72, 0.44, 0.32, 0.80, 0.22]
        elif nm == "Avenida Ulysses Silveira Guimarães":
            fracs = [0.24, 0.76, 0.50, 0.35, 0.65]
        else:
            fracs = [0.50, 0.40, 0.60, 0.30, 0.70, 0.22, 0.78, 0.15, 0.85]

        offsets_mm = [0.0, 1.3, -1.3, 2.4, -2.4, 3.5, -3.5]
        fs_candidates = [base_fs, max(5.9, base_fs - 0.4), max(5.5, base_fs - 0.8)]
        lbl_variants = [short_lbl]
        wrapped = wrap_two_lines(short_lbl)
        if "\n" in wrapped:
            lbl_variants.append(wrapped)
        on_road_variants = lbl_variants if nm == "Avenida Engenheiro Ildefonso Simões Lopes" else lbl_variants[:1]

        best = None
        for fs in fs_candidates:
            for lbl_txt in on_road_variants:
                for part in parts[:3]:
                    if part.length < 30.0 * MERC_PER_METER and len(parts) > 1:
                        continue
                    for off_mm in offsets_mm:
                        for fr in fracs:
                            x, y, ang = point_and_angle_along(part, fr, span_units=min(45.0 * MERC_PER_METER, max(14.0, part.length * 0.3)))
                            xm, ym = m3857_to_mm(x, y)
                            rad = math.radians(ang + 90.0)
                            xm_off = xm + off_mm * math.cos(rad)
                            ym_off = ym + off_mm * math.sin(rad)
                            lbox = make_label_box_mm(xm_off, ym_off, lbl_txt, fs, ang)
                            if not fits_in_map(lbox):
                                continue
                            if any(lbox.intersects(ob) for ob in occupied_boxes):
                                continue
                            ux_f, uy_f = mm_to_3857(xm_off, ym_off)
                            best = (ux_f, uy_f, ang, fs, lbox, False, x, y, lbl_txt)
                            break
                        if best:
                            break
                    if best:
                        break
                if best:
                    break
            if best:
                break

        if best is None:
            x0_s, y0_s, _ = point_and_angle_along(main_line, 0.5, span_units=22.0 * MERC_PER_METER)
            xm0, ym0 = m3857_to_mm(x0_s, y0_s)
            if nm == "Rua Maria José Campos Domingues":
                pref_degs = [45, 35, 55, 25, 65, 15, 75, 0, 90]
            elif nm == "Rua Alzira Gularte Farias":
                pref_degs = [110, 125, 95, 140, 80, 155, 170]
            else:
                pref_degs = [180, 165, 195, 150, 210, 135, 225, 120, 240, 90, 270, 60, 300, 30, 330, 0]
            found_c = False
            for fs in [6.0, 5.6]:
                for lbl_txt in lbl_variants[::-1]:
                    for radius in [6.0, 9.0, 12.5, 16.5, 21.0, 26.0, 32.0]:
                        for deg in pref_degs:
                            rad = math.radians(deg)
                            xm_c = xm0 + radius * math.cos(rad)
                            ym_c = ym0 + radius * math.sin(rad)
                            lbox = make_label_box_mm(xm_c, ym_c, lbl_txt, fs, 0.0, pad_x=0.75, pad_y=0.4)
                            if not fits_in_map(lbox):
                                continue
                            if any(lbox.intersects(ob) for ob in occupied_boxes):
                                continue
                            ux_f, uy_f = mm_to_3857(xm_c, ym_c)
                            best = (ux_f, uy_f, 0.0, fs, lbox, True, x0_s, y0_s, lbl_txt)
                            found_c = True
                            break
                        if found_c:
                            break
                    if found_c:
                        break
                if found_c:
                    break

        assert best is not None, f"Falha ao posicionar rótulo de território: {nm}"
        ux_f, uy_f, ang, fs, lbox, is_callout, ax_x, ax_y, final_lbl = best
        occupied_boxes.append(lbox)

        txt_color = "#7F1D1D" if info["is_border"] else "#0F172A"
        if is_callout:
            ax.annotate(
                final_lbl,
                xy=(ax_x, ax_y),
                xytext=(ux_f, uy_f),
                ha="center", va="center",
                fontsize=fs, fontweight="bold", color=txt_color,
                linespacing=1.05,
                bbox=dict(boxstyle="round,pad=0.20", facecolor="#FFFFFF", edgecolor="#64748B", lw=0.55, alpha=0.96),
                arrowprops=dict(arrowstyle="-", color="#475569", lw=0.7, shrinkA=1, shrinkB=1),
                zorder=18,
            )
        else:
            ax.text(
                ux_f, uy_f, final_lbl,
                ha="center", va="center",
                rotation=ang, rotation_mode="anchor",
                fontsize=fs, fontweight="bold", color=txt_color,
                path_effects=[pe.withStroke(linewidth=2.8, foreground="#FFFFFF")],
                zorder=17 if tier <= 2 else 16,
            )
        terr_placed.append((nm, final_lbl, fs, is_callout))

        # Segundo rótulo em avenidas muito longas (> 1.050 m) para leitura imediata nas duas pontas
        if info["total_len_m"] > 1050.0:
            for fr2 in [0.16, 0.84, 0.22, 0.78]:
                x2, y2, ang2 = point_and_angle_along(main_line, fr2, span_units=48.0 * MERC_PER_METER)
                xm2, ym2 = m3857_to_mm(x2, y2)
                lbox2 = make_label_box_mm(xm2, ym2, short_lbl, fs, ang2)
                if not fits_in_map(lbox2, margin=4.0):
                    continue
                if not any(lbox2.intersects(ob) for ob in occupied_boxes):
                    occupied_boxes.append(lbox2)
                    ax.text(
                        x2, y2, short_lbl,
                        ha="center", va="center",
                        rotation=ang2, rotation_mode="anchor",
                        fontsize=fs, fontweight="bold", color=txt_color,
                        path_effects=[pe.withStroke(linewidth=2.8, foreground="#FFFFFF")],
                        zorder=17,
                    )
                    break

    # -----------------------------------------------------------------------
    # Camada 6B: Principais Vias de Referência ao Redor (em toda a folha A4)
    # -----------------------------------------------------------------------
    # Excluir nomes já rotulados dentro/na divisa do território para nunca duplicar perto da borda
    # (exceto avenidas longas externas se houver trecho distante, já cobertas por anti-colisão!)
    inner_clip = box(
        WX0 + 8.0 / MAP_W_MM * WORLD_W_3857,
        WY0 + 6.0 / MAP_H_MM * WORLD_H_3857,
        WX1 - 8.0 / MAP_W_MM * WORLD_W_3857,
        WY1 - 6.0 / MAP_H_MM * WORLD_H_3857,
    )
    surr_streets = merge_roads_by_name(dataset["surrounding_roads"], clip_box=inner_clip)

    # Prioridade: grandes avenidas/rodovias primeiro, depois coletoras, depois residenciais longas do entorno
    ordered_surr = sorted(
        [nm for nm in surr_streets.keys() if nm not in terr_streets],
        key=lambda nm: (
            surrounding_street_tier(surr_streets[nm])[0],
            -surr_streets[nm]["total_len_m"],
            sort_key_pt(nm),
        ),
    )

    surr_placed = 0
    for nm in ordered_surr:
        info = surr_streets[nm]
        tier, base_fs = surrounding_street_tier(info)
        # Filtrar ruazinhas externas muito curtinhas para manter o entorno limpo e bem legível
        if tier == 3 and info["total_len_m"] < 150.0:
            continue
        if tier == 2 and info["total_len_m"] < 80.0:
            continue

        short_lbl = abbreviate_street(nm)
        parts = info["merged"]
        placed_this = False
        # Maior folga (pad_x/pad_y maiores) para as ruas externas não ficarem amontoadas!
        pad_x = 1.6 if tier == 3 else 1.1
        pad_y = 1.1 if tier == 3 else 0.7

        for part in parts[:2]:
            if part.length < 70.0 * MERC_PER_METER:
                continue
            for fr in [0.50, 0.35, 0.65, 0.22, 0.78]:
                x, y, ang = point_and_angle_along(part, fr, span_units=min(55.0 * MERC_PER_METER, max(18.0, part.length * 0.35)))
                # Evitar colocar rótulo externo muito colado no interior do polígono
                if poly_3857.buffer(15.0 * MERC_PER_METER).contains(Point(x, y)):
                    continue
                xm, ym = m3857_to_mm(x, y)
                lbox = make_label_box_mm(xm, ym, short_lbl, base_fs, ang, pad_x=pad_x, pad_y=pad_y)
                if not fits_in_map(lbox, margin=2.8):
                    continue
                if any(lbox.intersects(ob) for ob in occupied_boxes):
                    continue
                occupied_boxes.append(lbox)
                col = "#1E293B" if tier == 1 else ("#334155" if tier == 2 else "#475569")
                fw = "bold" if tier <= 2 else "normal"
                ax.text(
                    x, y, short_lbl,
                    ha="center", va="center",
                    rotation=ang, rotation_mode="anchor",
                    fontsize=base_fs, fontweight=fw, color=col,
                    path_effects=[pe.withStroke(linewidth=2.5, foreground="#FFFFFF")],
                    zorder=13 if tier <= 2 else 12,
                )
                surr_placed += 1
                placed_this = True
                break
            if placed_this:
                break

        # Avenidas externas muito extensas (> 900 m): 2º rótulo se houver espaço amplo
        if placed_this and tier <= 2 and info["total_len_m"] > 900.0:
            part = parts[0]
            for fr2 in [0.18, 0.82]:
                x2, y2, ang2 = point_and_angle_along(part, fr2, span_units=55.0 * MERC_PER_METER)
                if poly_3857.buffer(20.0 * MERC_PER_METER).contains(Point(x2, y2)):
                    continue
                xm2, ym2 = m3857_to_mm(x2, y2)
                lbox2 = make_label_box_mm(xm2, ym2, short_lbl, base_fs, ang2, pad_x=2.0, pad_y=1.2)
                if not fits_in_map(lbox2, margin=4.0):
                    continue
                if not any(lbox2.intersects(ob) for ob in occupied_boxes):
                    occupied_boxes.append(lbox2)
                    ax.text(
                        x2, y2, short_lbl,
                        ha="center", va="center",
                        rotation=ang2, rotation_mode="anchor",
                        fontsize=base_fs, fontweight="bold", color="#1E293B" if tier == 1 else "#334155",
                        path_effects=[pe.withStroke(linewidth=2.5, foreground="#FFFFFF")],
                        zorder=13,
                    )
                    break

    # -----------------------------------------------------------------------
    # Camada 7: Cartela Compacta (Título + Legenda + Norte + Escala — SW)
    # -----------------------------------------------------------------------
    ax_card = fig.add_axes([
        (MARGIN_MM + CARD_X_MM) / PAGE_W_MM,
        (MARGIN_MM + CARD_Y_MM) / PAGE_H_MM,
        CARD_W_MM / PAGE_W_MM,
        CARD_H_MM / PAGE_H_MM,
    ], zorder=25)
    ax_card.set_xlim(0, CARD_W_MM)
    ax_card.set_ylim(0, CARD_H_MM)
    ax_card.axis("off")

    ax_card.add_patch(FancyBboxPatch(
        (0.4, 0.4), CARD_W_MM - 0.8, CARD_H_MM - 0.8,
        boxstyle="round,pad=0.25,rounding_size=1.5",
        facecolor="#FFFFFF", edgecolor="#1E293B", linewidth=0.85, alpha=0.95,
    ))

    # Faixa superior fina (10,5 mm)
    HDR_H = 10.5
    ax_card.add_patch(FancyBboxPatch(
        (0.4, CARD_H_MM - HDR_H - 0.4), CARD_W_MM - 0.8, HDR_H,
        boxstyle="round,pad=0.25,rounding_size=1.5",
        facecolor="#0F2942", edgecolor="#0F2942", linewidth=0.4,
    ))
    ax_card.add_patch(Rectangle((0.2, CARD_H_MM - HDR_H - 0.6), CARD_W_MM - 0.4, 2.5, facecolor="#0F2942", edgecolor="none"))

    tx0 = 2.8
    if BRASAO_PATH.exists():
        img = Image.open(BRASAO_PATH).convert("RGBA")
        im_box = OffsetImage(img, zoom=0.036)
        ab = AnnotationBbox(im_box, (6.4, CARD_H_MM - HDR_H / 2.0 - 0.4), frameon=False)
        ax_card.add_artist(ab)
        tx0 = 12.2

    ax_card.text(tx0, CARD_H_MM - 3.6, "MAPA TERRITORIAL — UBS DUNAS", fontsize=7.5, fontweight="bold", color="#FFFFFF", va="center")
    ax_card.text(tx0, CARD_H_MM - 8.0, "SMS Pelotas/RS  •  Distrito Areal  •  Área: 1,32 km²  •  47 ruas", fontsize=5.1, color="#E2E8F0", va="center")

    # Legenda compacta (2 linhas x 2 colunas na metade esquerda)
    ly1, ly2 = 19.6, 14.4
    # Item 1: Território UBS Dunas
    ax_card.add_patch(Rectangle((2.5, ly1 - 1.3), 5.6, 2.6, facecolor="#FEF08A", edgecolor="#DC2626", lw=1.3))
    ax_card.text(9.5, ly1, "Território UBS Dunas", fontsize=5.3, fontweight="bold", color="#0F172A", va="center")
    # Item 2: UBS Dunas
    ax_card.add_patch(Circle((39.5, ly1), radius=1.55, facecolor="#DC2626", edgecolor="#FFFFFF", lw=0.6))
    ax_card.text(39.5, ly1, "+", ha="center", va="center", fontsize=5.8, fontweight="bold", color="#FFFFFF")
    ax_card.text(42.2, ly1, "UBS Dunas", fontsize=5.3, fontweight="bold", color="#991B1B", va="center")
    # Item 3: Vias limítrofes
    ax_card.text(2.5, ly2, "Av.", fontsize=5.3, fontweight="bold", color="#7F1D1D", va="center")
    ax_card.text(9.5, ly2, "Via limítrofe do território", fontsize=5.1, color="#334155", va="center")
    # Item 4: Escolas / CRAS
    ax_card.add_patch(Rectangle((38.2, ly2 - 0.9), 1.8, 1.8, facecolor="#1D4ED8", edgecolor="#FFFFFF", lw=0.3))
    ax_card.add_patch(Rectangle((40.4, ly2 - 0.9), 1.8, 1.8, facecolor="#047857", edgecolor="#FFFFFF", lw=0.3))
    ax_card.text(43.5, ly2, "EMEF / CRAS", fontsize=5.1, color="#334155", va="center")

    # Seta Norte compacta (canto direito da cartela)
    nx, ny = CARD_W_MM - 8.2, 14.5
    ax_card.add_patch(Circle((nx, ny), radius=4.8, facecolor="#F8FAFC", edgecolor="#CBD5E1", lw=0.55))
    ax_card.add_patch(MplPolygon([[nx, ny + 3.8], [nx + 1.8, ny - 2.6], [nx, ny - 1.2]], facecolor="#0F2942", edgecolor="#0F2942", lw=0.4))
    ax_card.add_patch(MplPolygon([[nx, ny + 3.8], [nx - 1.8, ny - 2.6], [nx, ny - 1.2]], facecolor="#FFFFFF", edgecolor="#0F2942", lw=0.5))
    ax_card.text(nx, ny + 6.4, "N", ha="center", va="center", fontsize=6.4, fontweight="bold", color="#0F2942")

    # Escala gráfica métrica exata (500 m = 500 / 11000 * 1000 = 45.45 mm)
    bar_x0, bar_y0 = 2.6, 6.2
    bar_400_mm = (400.0 / SCALE) * 1000.0  # 36.36 mm (4 blocos de 100 m)
    seg_w = bar_400_mm / 4.0
    for si in range(4):
        fc = "#0F2942" if si % 2 == 0 else "#FFFFFF"
        ax_card.add_patch(Rectangle((bar_x0 + si * seg_w, bar_y0), seg_w, 1.8, facecolor=fc, edgecolor="#0F2942", lw=0.55))
    for lbl, pos in [("0", 0), ("200 m", 2), ("400 m", 4)]:
        ax_card.text(bar_x0 + pos * seg_w, bar_y0 + 3.1, lbl, ha="center", va="center", fontsize=4.9, fontweight="bold", color="#0F172A")
    ax_card.text(bar_x0 + bar_400_mm + 3.5, bar_y0 + 1.0, f"Escala 1 : {int(SCALE):,}".replace(",", "."), ha="left", va="center", fontsize=5.4, fontweight="bold", color="#0F2942")

    # Linha final de fontes (ultra enxuta)
    ax_card.text(2.6, 2.3, "A4 Paisagem (imprimir em 100%)  •  Bases: Mapa UBSs Pelotas (SMS) • Google Maps • OSM", fontsize=4.4, color="#64748B", va="center")

    # Salvar PDF vetorial (A4 exato) e PNG 300 DPI
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    fig.savefig(PDF_OUT, format="pdf", dpi=300)
    fig.savefig(PNG_OUT, format="png", dpi=300)
    plt.close(fig)

    all_terr_sorted = sorted(terr_streets.keys(), key=sort_key_pt)
    return {
        "scale": int(SCALE),
        "total_territory_streets": len(all_terr_sorted),
        "internal_streets": sum(1 for k in all_terr_sorted if not terr_streets[k]["is_border"]),
        "border_streets": sum(1 for k in all_terr_sorted if terr_streets[k]["is_border"]),
        "surrounding_labeled": surr_placed,
        "callouts": [x[0] for x in terr_placed if x[3]],
        "streets": [(nm, terr_streets[nm]["is_border"]) for nm in all_terr_sorted],
    }


def write_readme(stats: dict) -> None:
    rows_md = "\n".join(
        f"| {i:02d} | {nm} | {'Limítrofe' if is_b else 'Interna'} |"
        for i, (nm, is_b) in enumerate(stats["streets"], 1)
    )
    readme = f"""# Mapa Territorial — UBS Dunas (Distrito Areal • Pelotas/RS)

Mapa cartográfico em **formato A4 Paisagem (`297 × 210 mm`)** com **margens mínimas (`2,5 mm`)**, preenchido em toda a folha com **mapa cartográfico real em alta resolução** (quadras, edificações, áreas verdes e ruas do entorno) e rotulagem vetorial nítida de **todas as {stats['total_territory_streets']} ruas/avenidas do território da UBS Dunas** ({stats['internal_streets']} internas + {stats['border_streets']} limítrofes) mais **{stats['surrounding_labeled']} vias de referência ao redor**.

## Arquivos disponíveis

| Arquivo | Formato | Uso recomendado |
|---------|---------|-----------------|
| [`Mapa_Territorio_UBS_Dunas_A4.pdf`](./Mapa_Territorio_UBS_Dunas_A4.pdf) | **PDF A4 (300 DPI + texto vetorial)** — Escala `1 : {stats['scale']:,}` | **Impressão em cópias** para as equipes (colorida ou P&B) |
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

## Relação completa de logradouros do território da UBS Dunas ({stats['total_territory_streets']})

| # | Logradouro | Classificação |
|---:|------------|---------------|
{rows_md}

## Como regenerar

```bash
pip install shapely pyproj matplotlib pillow requests
python3 scripts/gerar_mapa_ubs_dunas.py
# Para atualizar dados online (KML/OSM):
python3 scripts/gerar_mapa_ubs_dunas.py --refresh
```
"""
    (OUT_DIR / "README.md").write_text(readme, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description="Gera mapa A4 cartográfico real da UBS Dunas (Pelotas/RS).")
    parser.add_argument("--refresh", action="store_true", help="Rebaixa KML e OSM da internet.")
    args = parser.parse_args()

    dataset = load_or_build_dataset(refresh=args.refresh)
    stats = render_map(dataset)
    write_readme(stats)

    print(f"OK PDF: {PDF_OUT} ({PDF_OUT.stat().st_size / 1024:.1f} KB)")
    print(f"OK PNG: {PNG_OUT} ({PNG_OUT.stat().st_size / 1024:.1f} KB)")
    print(f"Território UBS Dunas: {stats['total_territory_streets']} ruas ({stats['internal_streets']} internas + {stats['border_streets']} limítrofes)")
    print(f"Vias de referência do entorno rotuladas: {stats['surrounding_labeled']}")
    print(f"Callouts utilizados ({len(stats['callouts'])}): {stats['callouts']}")


if __name__ == "__main__":
    main()
