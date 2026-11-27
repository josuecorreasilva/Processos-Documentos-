#!/usr/bin/env python3
"""
Gera o Mapa Territorial Oficial da UBS Dunas (Distrito Areal — Pelotas/RS)
em formato A4 Paisagem pronto para impressão (margens mínimas, PDF vetorial + PNG 300 DPI).

Base territorial:
  Mapa Território UBSs Pelotas (Google My Maps)
  https://www.google.com/maps/d/edit?mid=1D6SDvi-BRD-GIhiKeRsM6oYAgty-jDQ

Base viária e equipamentos públicos:
  OpenStreetMap (API v0.6 / ODbL) — projeção métrica SIRGAS 2000 / UTM 22S (EPSG:31982)
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import unicodedata
import xml.etree.ElementTree as ET
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patheffects as pe
from matplotlib.patches import FancyBboxPatch, Polygon as MplPolygon, Rectangle, Circle
from matplotlib.offsetbox import OffsetImage, AnnotationBbox
import numpy as np
from PIL import Image
from pyproj import Transformer
from shapely.affinity import rotate as shp_rotate
from shapely.geometry import LineString, MultiLineString, Point, Polygon, box, mapping
from shapely.ops import linemerge, unary_union

REPO_ROOT = Path(__file__).resolve().parent.parent
OUT_DIR = REPO_ROOT / "docs" / "07-MAPA-TERRITORIO-UBS-DUNAS"
BRASAO_PATH = REPO_ROOT / "docs" / "assets" / "brasao_pelotas.png"
DATA_CACHE = OUT_DIR / "ubs_dunas_dados.json"
GEOJSON_OUT = OUT_DIR / "ubs_dunas_territorio.geojson"
PDF_OUT = OUT_DIR / "Mapa_Territorio_UBS_Dunas_A4.pdf"
PNG_OUT = OUT_DIR / "Mapa_Territorio_UBS_Dunas_A4_300dpi.png"

KML_URL = "https://www.google.com/maps/d/kml?mid=1D6SDvi-BRD-GIhiKeRsM6oYAgty-jDQ&forcekml=1"
TMP_KML = Path("/tmp/ubs_pelotas.kml")
TMP_OSM = Path("/tmp/ubs_dunas_osm.json")

TO_UTM = Transformer.from_crs("EPSG:4326", "EPSG:31982", always_xy=True)
TO_WGS = Transformer.from_crs("EPSG:31982", "EPSG:4326", always_xy=True)

# ---------------------------------------------------------------------------
# 1. Extração e preparação dos dados (KML + OSM)
# ---------------------------------------------------------------------------

def sort_key_pt(s: str) -> str:
    nfkd = unicodedata.normalize("NFKD", s)
    return "".join(c for c in nfkd if not unicodedata.combining(c)).lower()


def abbreviate_street(name: str) -> str:
    """Abreviação cartográfica padrão IBGE para melhorar legibilidade no traçado."""
    s = name.strip()
    replacements = [
        (r"^Avenida Engenheiro ", "Av. Eng. "),
        (r"^Avenida Doutor ", "Av. Dr. "),
        (r"^Avenida ", "Av. "),
        (r"^Travessa ", "Tv. "),
        (r"^Passeio ", "Pass. "),
        (r"^Rua Vereador ", "R. Ver. "),
        (r"^Rua Jornalista ", "R. Jorn. "),
        (r"^Rua ", "R. "),
    ]
    for pat, rep in replacements:
        if re.match(pat, s):
            s = re.sub(pat, rep, s, count=1)
            break
    s = s.replace(" Loteamento Dunas", " (Lot. Dunas)")
    s = s.replace(" Lot Bom Jesus", " (Lot. Bom Jesus)")
    return s


def load_or_build_dataset(refresh: bool = False) -> dict:
    if DATA_CACHE.exists() and not refresh:
        with open(DATA_CACHE, encoding="utf-8") as f:
            return json.load(f)

    import requests

    if refresh or not TMP_KML.exists():
        r = requests.get(KML_URL, timeout=60, headers={"User-Agent": "SMSPelotas-MapaUBSDunas/1.0"})
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
    poly_geom_wgs = Polygon(poly_wgs)
    minx, miny, maxx, maxy = poly_geom_wgs.bounds
    pad = 0.0025

    if refresh or not TMP_OSM.exists():
        bbox_osm = f"{minx - pad:.6f},{miny - pad:.6f},{maxx + pad:.6f},{maxy + pad:.6f}"
        url = f"https://api.openstreetmap.org/api/0.6/map.json?bbox={bbox_osm}"
        r = requests.get(url, timeout=60, headers={"User-Agent": "SMSPelotas-MapaUBSDunas/1.0"})
        r.raise_for_status()
        TMP_OSM.write_bytes(r.content)

    with open(TMP_OSM, encoding="utf-8") as f:
        osm = json.load(f)

    nodes = {e["id"]: (e["lon"], e["lat"]) for e in osm["elements"] if e["type"] == "node"}
    poly_utm = Polygon([TO_UTM.transform(x, y) for x, y in poly_wgs])
    poly_border_18 = poly_utm.buffer(18.0)
    poly_clip_24 = poly_utm.buffer(24.0)

    # Coletar vias dentro ou limítrofes ao polígono
    roads = []
    parks = []
    pois = []

    for e in osm["elements"]:
        tags = e.get("tags", {})
        if e["type"] == "way":
            pts_wgs = [nodes[nid] for nid in e.get("nodes", []) if nid in nodes]
            if len(pts_wgs) < 2:
                continue
            pts_utm = [TO_UTM.transform(x, y) for x, y in pts_wgs]

            hw = tags.get("highway")
            if hw:
                line = LineString(pts_utm)
                inter_in = line.intersection(poly_utm)
                inter_b = line.intersection(poly_border_18)
                is_internal = inter_in.length > 5.0
                # Limítrofe real: acompanha o contorno por pelo menos 80 m e fica a <= 15 m do polígono
                is_border = (not is_internal) and (inter_b.length > 80.0 and line.distance(poly_utm) < 15.0)
                if not (is_internal or is_border):
                    continue
                clipped = line.intersection(poly_clip_24)
                segs = []
                if isinstance(clipped, LineString):
                    segs = [list(clipped.coords)]
                elif isinstance(clipped, MultiLineString):
                    segs = [list(g.coords) for g in clipped.geoms if g.length > 2.0]
                elif hasattr(clipped, "geoms"):
                    segs = [list(g.coords) for g in clipped.geoms if isinstance(g, LineString) and g.length > 2.0]
                if not segs:
                    continue
                name = tags.get("name", "").strip()
                # Pequeno trecho (way/860893278) sem tag name que é a continuidade direta da Travessa Manoel Antônio Peres
                roads.append({
                    "id": e["id"],
                    "name": name,
                    "highway": hw,
                    "is_border": is_border,
                    "segments_utm": segs,
                })
                continue

            # Parques / quadras esportivas dentro do polígono
            if (tags.get("leisure") in ("park", "pitch", "playground", "garden") or tags.get("natural") == "water") and len(pts_utm) >= 3:
                if pts_utm[0] == pts_utm[-1]:
                    p = Polygon(pts_utm)
                    if p.is_valid and p.intersects(poly_utm):
                        clipped_p = p.intersection(poly_utm)
                        if not clipped_p.is_empty and clipped_p.area > 30:
                            geoms = [clipped_p] if isinstance(clipped_p, Polygon) else [g for g in getattr(clipped_p, "geoms", []) if isinstance(g, Polygon)]
                            for g in geoms:
                                parks.append({
                                    "kind": tags.get("leisure") or tags.get("natural"),
                                    "coords_utm": list(g.exterior.coords),
                                })

            # Equipamentos públicos relevantes (escolas, CRAS, clube)
            am = tags.get("amenity", "")
            nm = tags.get("name", "").strip()
            if am in ("school", "social_centre", "community_centre") and nm:
                c = LineString(pts_utm).centroid
                if poly_border_18.contains(c):
                    pois.append({"name": nm, "amenity": am, "x": c.x, "y": c.y})

    # Deduplicar POIs por nome (ex.: CRAS tem 2 blocos contíguos)
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
        "source_mymaps": "https://www.google.com/maps/d/edit?mid=1D6SDvi-BRD-GIhiKeRsM6oYAgty-jDQ",
        "crs": "EPSG:31982",
        "polygon_wgs84": poly_wgs,
        "polygon_utm": [list(TO_UTM.transform(x, y)) for x, y in poly_wgs],
        "ubs_point_wgs84": ubs_pt_wgs,
        "ubs_point_utm": list(TO_UTM.transform(*ubs_pt_wgs)),
        "roads": roads,
        "parks": parks,
        "pois": pois_final,
    }

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    with open(DATA_CACHE, "w", encoding="utf-8") as f:
        json.dump(dataset, f, ensure_ascii=False)

    # Exportar também GeoJSON padrão WGS84 (auditável / QGIS / Google Earth)
    features = [
        {
            "type": "Feature",
            "properties": {
                "nome": "Território UBS Dunas",
                "distrito": "Distrito Areal",
                "municipio": "Pelotas/RS",
                "area_km2": round(poly_utm.area / 1e6, 4),
                "perimetro_m": round(poly_utm.length, 1),
            },
            "geometry": mapping(Polygon(poly_wgs)),
        },
        {
            "type": "Feature",
            "properties": {"nome": "UBS Dunas (Distrito Areal)", "tipo": "UBS"},
            "geometry": mapping(Point(ubs_pt_wgs)),
        },
    ]
    for r in roads:
        if not r["name"]:
            continue
        for seg in r["segments_utm"]:
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
# 2. Motor de rotulagem inteligente (alinhado à via + anti-colisão em mm)
# ---------------------------------------------------------------------------

def merge_named_streets(roads: list[dict]) -> dict[str, dict]:
    """Agrupa e mescla segmentos por nome oficial de rua."""
    grouped: dict[str, dict] = {}
    for r in roads:
        nm = r["name"]
        if not nm:
            continue
        entry = grouped.setdefault(nm, {"lines": [], "hw": set(), "is_border": True})
        entry["hw"].add(r["highway"])
        if not r["is_border"]:
            entry["is_border"] = False
        for seg in r["segments_utm"]:
            if len(seg) >= 2:
                entry["lines"].append(LineString(seg))

    for nm, entry in grouped.items():
        merged = linemerge(entry["lines"])
        if isinstance(merged, LineString):
            parts = [merged]
        elif isinstance(merged, MultiLineString):
            parts = sorted(list(merged.geoms), key=lambda g: g.length, reverse=True)
        else:
            parts = sorted(entry["lines"], key=lambda g: g.length, reverse=True)
        entry["merged"] = parts
        entry["total_len"] = sum(g.length for g in parts)
    return grouped


def point_and_angle_along(line: LineString, frac: float, span_m: float = 35.0) -> tuple[float, float, float]:
    """Retorna (x, y, angulo_graus) suavizado ao redor da fração `frac` da linha."""
    L = line.length
    d = max(0.0, min(L, L * frac))
    p = line.interpolate(d)
    d0 = max(0.0, d - span_m / 2.0)
    d1 = min(L, d + span_m / 2.0)
    if d1 - d0 < 1.0:
        p0, p1 = Point(line.coords[0]), Point(line.coords[-1])
    else:
        p0, p1 = line.interpolate(d0), line.interpolate(d1)
    ang = math.degrees(math.atan2(p1.y - p0.y, p1.x - p0.x))
    # Manter texto sempre legível (nunca de cabeça para baixo)
    if ang > 90.0:
        ang -= 180.0
    elif ang < -90.0:
        ang += 180.0
    return p.x, p.y, ang


def wrap_two_lines(text: str) -> str:
    """Quebra um nome longo em 2 linhas equilibradas para uso em rótulos compactos."""
    words = text.split()
    if len(words) <= 2 or len(text) <= 14:
        return text
    best_split = 1
    best_diff = 999
    for i in range(1, len(words)):
        l1 = " ".join(words[:i])
        l2 = " ".join(words[i:])
        diff = abs(len(l1) - len(l2))
        if diff < best_diff:
            best_diff = diff
            best_split = i
    return " ".join(words[:best_split]) + "\n" + " ".join(words[best_split:])


def make_label_box_mm(x_mm: float, y_mm: float, text: str, fontsize_pt: float, angle_deg: float, pad_x: float = 0.45, pad_y: float = 0.16) -> Polygon:
    """Cria polígono rotacionado aproximado (em mm) ocupado pelo texto (1 ou 2 linhas) na página."""
    lines = text.split("\n")
    max_chars = max(len(ln) for ln in lines)
    char_w_mm = fontsize_pt * 0.35278 * 0.50
    text_w = max(3.2, max_chars * char_w_mm) + 2 * pad_x
    text_h = fontsize_pt * 0.35278 * (1.05 * len(lines) + 0.10 * (len(lines) - 1)) + 2 * pad_y
    b = box(x_mm - text_w / 2.0, y_mm - text_h / 2.0, x_mm + text_w / 2.0, y_mm + text_h / 2.0)
    return shp_rotate(b, angle_deg, origin=(x_mm, y_mm))


def street_tier(info: dict) -> tuple[int, float]:
    """Retorna (prioridade, fontsize_pt) pela hierarquia viária."""
    hws = info["hw"]
    if info["is_border"] and "secondary" in hws:
        return (1, 7.6)
    if "secondary" in hws:
        return (1, 7.6)
    if info["is_border"]:
        return (1, 7.2)
    if "tertiary" in hws:
        return (2, 6.8)
    return (3, 6.2)


# ---------------------------------------------------------------------------
# 3. Renderização do Mapa Cartográfico A4 Paisagem (margens mínimas)
# ---------------------------------------------------------------------------

def render_map(dataset: dict) -> dict:
    plt.rcParams["font.family"] = "sans-serif"
    plt.rcParams["font.sans-serif"] = ["Liberation Sans", "DejaVu Sans", "Noto Sans"]
    plt.rcParams["pdf.fonttype"] = 42  # TrueType embutido no PDF (100% vetorial e pesquisável)
    plt.rcParams["ps.fonttype"] = 42

    poly_utm = Polygon(dataset["polygon_utm"])
    ux, uy = dataset["ubs_point_utm"]
    minx, miny, maxx, maxy = poly_utm.bounds

    # Dimensões exatas A4 Paisagem (mm -> polegadas)
    PAGE_W_MM, PAGE_H_MM = 297.0, 210.0
    MM_TO_IN = 1.0 / 25.4
    fig = plt.figure(figsize=(PAGE_W_MM * MM_TO_IN, PAGE_H_MM * MM_TO_IN), dpi=300, facecolor="#FFFFFF")

    # Margens mínimas de impressora (3.0 mm) + faixa fina da grade alfanumérica (3.3 mm)
    OUTER_M = 3.0
    GRID_STRIP = 3.2
    MAP_X0_MM = OUTER_M + GRID_STRIP
    MAP_Y0_MM = OUTER_M + GRID_STRIP
    MAP_W_MM = PAGE_W_MM - 2 * (OUTER_M + GRID_STRIP)   # 284.6 mm
    MAP_H_MM = PAGE_H_MM - 2 * (OUTER_M + GRID_STRIP)   # 197.6 mm

    # Escala exata 1:11.500 (1 cm = 115 m) — máximo aproveitamento da altura útil A4!
    SCALE = 11500.0
    WORLD_W = (MAP_W_MM / 1000.0) * SCALE  # 3272.9 m
    WORLD_H = (MAP_H_MM / 1000.0) * SCALE  # 2272.4 m

    pad_y = (WORLD_H - (maxy - miny)) / 2.0
    WX0 = minx - 85.0
    WX1 = WX0 + WORLD_W
    WY0 = miny - pad_y
    WY1 = maxy + pad_y

    def utm_to_mm(x: float, y: float) -> tuple[float, float]:
        return ((x - WX0) / WORLD_W * MAP_W_MM, (y - WY0) / WORLD_H * MAP_H_MM)

    def mm_to_utm(xm: float, ym: float) -> tuple[float, float]:
        return (WX0 + (xm / MAP_W_MM) * WORLD_W, WY0 + (ym / MAP_H_MM) * WORLD_H)

    # Eixo principal do mapa (ocupando praticamente toda a folha A4)
    ax = fig.add_axes([
        MAP_X0_MM / PAGE_W_MM,
        MAP_Y0_MM / PAGE_H_MM,
        MAP_W_MM / PAGE_W_MM,
        MAP_H_MM / PAGE_H_MM,
    ])
    ax.set_xlim(WX0, WX1)
    ax.set_ylim(WY0, WY1)
    ax.set_aspect("equal", adjustable="box")
    ax.set_facecolor("#FFFFFF")
    ax.set_xticks([])
    ax.set_yticks([])
    for sp in ax.spines.values():
        sp.set_edgecolor("#1E293B")
        sp.set_linewidth(0.9)

    # Eixo auxiliar da moldura externa + rótulos da grade alfanumérica (A-G / 1-5)
    ax_frame = fig.add_axes([0, 0, 1, 1], zorder=1)
    ax_frame.set_xlim(0, PAGE_W_MM)
    ax_frame.set_ylim(0, PAGE_H_MM)
    ax_frame.axis("off")
    ax_frame.add_patch(Rectangle(
        (OUTER_M, OUTER_M),
        PAGE_W_MM - 2 * OUTER_M,
        PAGE_H_MM - 2 * OUTER_M,
        facecolor="#F8FAFC",
        edgecolor="#1E293B",
        linewidth=0.9,
        zorder=0,
    ))

    # Grade alfanumérica discreta: 7 colunas (A-G) x 5 linhas (1-5)
    COLS = ["A", "B", "C", "D", "E", "F", "G"]
    ROWS = ["1", "2", "3", "4", "5"]  # 1 no topo, 5 na base
    col_edges_x = np.linspace(WX0, WX1, len(COLS) + 1)
    row_edges_y = np.linspace(WY1, WY0, len(ROWS) + 1)  # decrescente (1 no topo)

    for cx in col_edges_x[1:-1]:
        ax.axvline(cx, color="#E2E8F0", linewidth=0.45, linestyle="-", zorder=1)
    for ry in row_edges_y[1:-1]:
        ax.axhline(ry, color="#E2E8F0", linewidth=0.45, linestyle="-", zorder=1)

    for i, lbl in enumerate(COLS):
        xm = MAP_X0_MM + (i + 0.5) * (MAP_W_MM / len(COLS))
        for ym in (OUTER_M + GRID_STRIP / 2.0, PAGE_H_MM - OUTER_M - GRID_STRIP / 2.0):
            ax_frame.text(xm, ym, lbl, ha="center", va="center", fontsize=6.5, fontweight="bold", color="#334155")
        if i > 0:
            xsep = MAP_X0_MM + i * (MAP_W_MM / len(COLS))
            ax_frame.plot([xsep, xsep], [OUTER_M, MAP_Y0_MM], color="#94A3B8", lw=0.4)
            ax_frame.plot([xsep, xsep], [MAP_Y0_MM + MAP_H_MM, PAGE_H_MM - OUTER_M], color="#94A3B8", lw=0.4)

    for j, lbl in enumerate(ROWS):
        ym = MAP_Y0_MM + MAP_H_MM - (j + 0.5) * (MAP_H_MM / len(ROWS))
        for xm in (OUTER_M + GRID_STRIP / 2.0, PAGE_W_MM - OUTER_M - GRID_STRIP / 2.0):
            ax_frame.text(xm, ym, lbl, ha="center", va="center", fontsize=6.5, fontweight="bold", color="#334155")
        if j > 0:
            ysep = MAP_Y0_MM + MAP_H_MM - j * (MAP_H_MM / len(ROWS))
            ax_frame.plot([OUTER_M, MAP_X0_MM], [ysep, ysep], color="#94A3B8", lw=0.4)
            ax_frame.plot([MAP_X0_MM + MAP_W_MM, PAGE_W_MM - OUTER_M], [ysep, ysep], color="#94A3B8", lw=0.4)

    def get_grid_code(x: float, y: float) -> str:
        ci = int(np.clip(np.searchsorted(col_edges_x, x, side="right") - 1, 0, len(COLS) - 1))
        ri = int(np.clip(np.floor((WY1 - y) / (WY1 - WY0) * len(ROWS)), 0, len(ROWS) - 1))
        return f"{COLS[ci]}{ROWS[ri]}"

    # -----------------------------------------------------------------------
    # Camada 1: Polígono do Território da UBS Dunas (preenchimento alto contraste)
    # -----------------------------------------------------------------------
    px, py = poly_utm.exterior.xy
    # Sombra/halo suave externo
    ax.fill(px, py, facecolor="#FEFCE8", edgecolor="none", zorder=2)

    # Camada 1b: Áreas verdes / quadras dentro do território
    for pk in dataset["parks"]:
        coords = pk["coords_utm"]
        c_arr = np.array(coords)
        fc = "#D1FAE5" if pk["kind"] in ("park", "garden") else "#DCFCE7"
        ax.fill(c_arr[:, 0], c_arr[:, 1], facecolor=fc, edgecolor="#86EFAC", linewidth=0.4, zorder=3)

    # -----------------------------------------------------------------------
    # Camada 2: Malha Viária Hierarquizada (estilo casing escuro + núcleo branco)
    # -----------------------------------------------------------------------
    roads = dataset["roads"]
    # Ordem de desenho: service -> residential -> tertiary -> secondary
    # Passada A: casing (contorno escuro)
    for r in roads:
        hw = r["highway"]
        for seg in r["segments_utm"]:
            xs, ys = zip(*seg)
            if hw in ("service", "cycleway", "footway", "path"):
                ax.plot(xs, ys, color="#94A3B8", lw=0.75, ls="--" if hw == "cycleway" else "-", solid_capstyle="round", zorder=4)
            elif hw in ("residential", "unclassified"):
                ax.plot(xs, ys, color="#475569", lw=2.2, solid_capstyle="round", solid_joinstyle="round", zorder=5)
            elif hw == "tertiary":
                ax.plot(xs, ys, color="#1E293B", lw=3.2, solid_capstyle="round", solid_joinstyle="round", zorder=6)
            elif hw == "secondary":
                ax.plot(xs, ys, color="#0F172A", lw=4.1, solid_capstyle="round", solid_joinstyle="round", zorder=7)

    # Passada B: núcleo (branco / âmbar suave para estruturais)
    for r in roads:
        hw = r["highway"]
        for seg in r["segments_utm"]:
            xs, ys = zip(*seg)
            if hw in ("residential", "unclassified"):
                ax.plot(xs, ys, color="#FFFFFF", lw=1.25, solid_capstyle="round", solid_joinstyle="round", zorder=8)
            elif hw == "tertiary":
                ax.plot(xs, ys, color="#FFFBEB", lw=1.95, solid_capstyle="round", solid_joinstyle="round", zorder=9)
            elif hw == "secondary":
                ax.plot(xs, ys, color="#FDE68A", lw=2.55, solid_capstyle="round", solid_joinstyle="round", zorder=10)

    # -----------------------------------------------------------------------
    # Camada 3: Contorno Oficial do Território (UBS Dunas) — bem destacado
    # -----------------------------------------------------------------------
    ax.plot(px, py, color="#FFFFFF", lw=4.0, solid_joinstyle="round", zorder=11, alpha=0.85)
    ax.plot(px, py, color="#B91C1C", lw=2.3, solid_joinstyle="round", zorder=12)

    # -----------------------------------------------------------------------
    # Camada 4: Caixas de painéis nos cantos livres (para bloquear colisão de rótulos)
    # -----------------------------------------------------------------------
    # Painel NE (canto superior direito): Cabeçalho compacto + Índice completo das 47 ruas
    NE_W_MM, NE_H_MM = 113.0, 136.0
    NE_X_MM = MAP_W_MM - NE_W_MM - 1.8
    NE_Y_MM = MAP_H_MM - NE_H_MM - 1.8

    # Painel SW (canto inferior esquerdo): Legenda + Escala + Norte + Metadados
    SW_W_MM, SW_H_MM = 90.0, 59.0
    SW_X_MM = 2.0
    SW_Y_MM = 2.0

    # Áreas proibidas para rótulos de rua (em mm relativos ao eixo do mapa)
    occupied_boxes: list[Polygon] = [
        box(NE_X_MM - 2.0, NE_Y_MM - 2.0, MAP_W_MM + 5.0, MAP_H_MM + 5.0),
        box(-5.0, -5.0, SW_X_MM + SW_W_MM + 2.0, SW_Y_MM + SW_H_MM + 2.0),
    ]

    # -----------------------------------------------------------------------
    # Camada 5: Equipamentos Públicos & Marcador Principal da UBS Dunas
    # -----------------------------------------------------------------------
    ux_mm, uy_mm = utm_to_mm(ux, uy)
    # Cruz da UBS Dunas (marcador oficial destacado)
    ax.scatter([ux], [uy], s=145, c="#DC2626", edgecolors="#FFFFFF", linewidths=1.8, zorder=22)
    ax.text(ux, uy, "+", ha="center", va="center", fontsize=9.5, fontweight="bold", color="#FFFFFF", zorder=23)

    # Etiqueta UBS DUNAS posicionada no vão livre a sudoeste da unidade (sem cobrir ruas)
    ubs_tx, ubs_ty = ux - 195, uy - 125
    ax.annotate(
        "UBS DUNAS",
        xy=(ux, uy),
        xytext=(ubs_tx, ubs_ty),
        ha="center",
        va="center",
        fontsize=7.6,
        fontweight="bold",
        color="#991B1B",
        bbox=dict(boxstyle="round,pad=0.25", facecolor="#FEF2F2", edgecolor="#DC2626", lw=1.15),
        arrowprops=dict(arrowstyle="-|>", color="#DC2626", lw=1.15, shrinkA=2, shrinkB=6, mutation_scale=9),
        zorder=24,
    )
    ubs_bx_mm, ubs_by_mm = utm_to_mm(ubs_tx, ubs_ty)
    occupied_boxes.append(box(ux_mm - 3.2, uy_mm - 3.2, ux_mm + 3.2, uy_mm + 3.2))
    occupied_boxes.append(box(ubs_bx_mm - 11.0, ubs_by_mm - 3.0, ubs_bx_mm + 11.0, ubs_by_mm + 3.0))

    # Outros POIs (Escolas, CRAS, Dunas Clube) — ícones discretos + rótulos curtos
    poi_short = {
        "EMEF Núcleo Habitacional Dunas": ("EMEF Núc. Hab. Dunas", "#1D4ED8", (-10, -28)),
        "EMEF Jornalista Deogar Soares": ("EMEF Deogar Soares", "#1D4ED8", (-65, -24)),
        "CRAS - Centro de Referência de Assistência Social": ("CRAS Dunas", "#047857", (52, 22)),
        "Dunas Clube": ("Dunas Clube", "#6D28D9", (0, 32)),
    }
    for p in dataset["pois"]:
        if p["name"] not in poi_short:
            continue
        lbl, col, (dx, dy) = poi_short[p["name"]]
        ax.scatter([p["x"]], [p["y"]], s=24, marker="s", c=col, edgecolors="#FFFFFF", linewidths=0.7, zorder=19)
        tx, ty = p["x"] + dx, p["y"] + dy
        ax.text(
            tx, ty, lbl,
            ha="center", va="center",
            fontsize=5.3, fontweight="bold", color=col,
            path_effects=[pe.withStroke(linewidth=2.0, foreground="#FFFFFF")],
            zorder=20,
        )
        pxm, pym = utm_to_mm(tx, ty)
        occupied_boxes.append(make_label_box_mm(pxm, pym, lbl, 5.3, 0.0, pad_x=0.5, pad_y=0.3))

    # -----------------------------------------------------------------------
    # Camada 6: Rotulagem Inteligente de TODAS as 47 Ruas (Zero Sobreposição)
    # -----------------------------------------------------------------------
    named_streets = merge_named_streets(roads)

    # Escalonamento alternado para as séries paralelas do Lot. Dunas (Ruas 6..18 e 21..28),
    # garantindo máximo arejamento e fonte grande diretamente sobre cada rua!
    stagger_a = {
        "Rua Seis", "Rua Oito", "Rua Dez", "Rua Doze", "Rua Quatorze", "Rua Dezesseis", "Rua Dezoito",
        "Rua Vinte e Um", "Rua Vinte e Três", "Rua Vinte e Cinco", "Rua Vinte e Sete",
    }
    stagger_b = {
        "Rua Sete", "Rua Nove", "Rua Onze", "Rua Treze", "Rua Quinze", "Rua Dezessete",
        "Rua Vinte e Dois", "Rua Vinte e Quatro", "Rua Vinte e Seis", "Rua Vinte e Oito",
    }

    # Prioridade: grandes avenidas/eixos estruturais primeiro, depois coletoras, depois residenciais por comprimento
    ordered_names = sorted(
        named_streets.keys(),
        key=lambda nm: (
            street_tier(named_streets[nm])[0],
            -named_streets[nm]["total_len"] if street_tier(named_streets[nm])[0] <= 2 else named_streets[nm]["total_len"],
            sort_key_pt(nm),
        ),
    )

    placed_labels = []
    street_grid_codes: dict[str, str] = {}

    for nm in ordered_names:
        info = named_streets[nm]
        tier, base_fs = street_tier(info)
        parts = info["merged"]
        main_line = parts[0]

        # Código da grade pelo ponto médio do trecho principal
        mid_pt = main_line.interpolate(main_line.length * 0.5)
        street_grid_codes[nm] = get_grid_code(mid_pt.x, mid_pt.y)

        short_lbl = abbreviate_street(nm)

        # Para vias residenciais/coletoras longas que cruzam muitas transversais, dar prioridade no posicionamento na via
        if nm in ("Rua Vereador Élbio Abreu", "Rua João Baptista Marimon", "Passeio Dois Loteamento Dunas"):
            fracs = [0.22, 0.78, 0.30, 0.70, 0.50, 0.15, 0.85, 0.40, 0.60]
        elif nm == "Avenida Engenheiro Ildefonso Simões Lopes":
            fracs = [0.45, 0.55, 0.35, 0.65, 0.25, 0.75]
        elif nm in stagger_a:
            fracs = [0.36, 0.44, 0.28, 0.56, 0.68, 0.20, 0.78]
        elif nm in stagger_b:
            fracs = [0.64, 0.56, 0.72, 0.44, 0.32, 0.80, 0.22]
        elif nm == "Avenida Ulysses Silveira Guimarães":
            fracs = [0.24, 0.76, 0.50, 0.35, 0.65, 0.16, 0.84]
        else:
            fracs = [0.50, 0.40, 0.60, 0.30, 0.70, 0.22, 0.78, 0.15, 0.85]

        offsets_mm = [0.0, 1.4, -1.4, 2.5, -2.5, 3.6, -3.6]
        fs_candidates = [base_fs, max(5.6, base_fs - 0.4), max(5.2, base_fs - 0.8)]
        lbl_variants = [short_lbl]
        wrapped_lbl = wrap_two_lines(short_lbl)
        if "\n" in wrapped_lbl:
            lbl_variants.append(wrapped_lbl)

        # Permitir 2 linhas na própria via para avenidas longas em trechos curtos (como Av. Eng. Ildefonso Simões Lopes)
        on_road_variants = lbl_variants if nm == "Avenida Engenheiro Ildefonso Simões Lopes" else lbl_variants[:1]

        def fits_in_map(lbox: Polygon) -> bool:
            bx0, by0, bx1, by1 = lbox.bounds
            return bx0 >= 1.2 and by0 >= 1.2 and bx1 <= MAP_W_MM - 1.2 and by1 <= MAP_H_MM - 1.2

        best = None
        for fs in fs_candidates:
            for lbl_txt in on_road_variants:
                for part in parts[:3]:
                    if part.length < 30.0 and len(parts) > 1:
                        continue
                    for off_mm in offsets_mm:
                        for fr in fracs:
                            x, y, ang = point_and_angle_along(part, fr, span_m=min(40.0, max(12.0, part.length * 0.3)))
                            xm, ym = utm_to_mm(x, y)
                            rad = math.radians(ang + 90.0)
                            xm_off = xm + off_mm * math.cos(rad)
                            ym_off = ym + off_mm * math.sin(rad)
                            lbox = make_label_box_mm(xm_off, ym_off, lbl_txt, fs, ang)
                            if not fits_in_map(lbox):
                                continue
                            if any(lbox.intersects(ob) for ob in occupied_boxes):
                                continue
                            ux_f, uy_f = mm_to_utm(xm_off, ym_off)
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

        # Fallback com linha de chamada (callout/leader line) somente se a rua for muito curta/espremida
        if best is None:
            x0_s, y0_s, _ = point_and_angle_along(main_line, 0.5, span_m=20.0)
            xm0, ym0 = utm_to_mm(x0_s, y0_s)
            found_callout = False
            # Preferir ângulos apontando para o espaço externo livre sem cruzar outras ruas/setas
            if nm == "Rua Maria José Campos Domingues":
                pref_degs = [45, 35, 55, 25, 65, 15, 75, 0, 90]
            elif nm == "Rua Alzira Gularte Farias":
                pref_degs = [110, 125, 95, 140, 80, 155, 170]
            else:
                pref_degs = [180, 165, 195, 150, 210, 135, 225, 120, 240, 90, 270, 60, 300, 30, 330, 0]
            for fs in [5.6, 5.2]:
                for lbl_txt in (lbl_variants[::-1]):
                    for radius in [6.0, 9.0, 12.5, 16.5, 21.0, 26.0, 32.0]:
                        for deg in pref_degs:
                            rad = math.radians(deg)
                            xm_c = xm0 + radius * math.cos(rad)
                            ym_c = ym0 + radius * math.sin(rad)
                            lbox = make_label_box_mm(xm_c, ym_c, lbl_txt, fs, 0.0, pad_x=0.7, pad_y=0.4)
                            if not fits_in_map(lbox):
                                continue
                            if any(lbox.intersects(ob) for ob in occupied_boxes):
                                continue
                            ux_f, uy_f = mm_to_utm(xm_c, ym_c)
                            best = (ux_f, uy_f, 0.0, fs, lbox, True, x0_s, y0_s, lbl_txt)
                            found_callout = True
                            break
                        if found_callout:
                            break
                    if found_callout:
                        break
                if found_callout:
                    break

        assert best is not None, f"Não foi possível posicionar o rótulo de: {nm}"
        ux_f, uy_f, ang, fs, lbox, is_callout, ax_x, ax_y, final_lbl = best
        occupied_boxes.append(lbox)

        txt_color = "#7C2D12" if info["is_border"] else ("#0F172A" if tier <= 2 else "#1E293B")
        if is_callout:
            ax.annotate(
                final_lbl,
                xy=(ax_x, ax_y),
                xytext=(ux_f, uy_f),
                ha="center",
                va="center",
                fontsize=fs,
                fontweight="bold",
                color=txt_color,
                linespacing=1.05,
                bbox=dict(boxstyle="round,pad=0.18", facecolor="#FFFFFF", edgecolor="#64748B", lw=0.5, alpha=0.96),
                arrowprops=dict(arrowstyle="-", color="#475569", lw=0.65, shrinkA=1, shrinkB=1),
                zorder=18,
            )
        else:
            ax.text(
                ux_f, uy_f, final_lbl,
                ha="center", va="center",
                rotation=ang,
                rotation_mode="anchor",
                fontsize=fs,
                fontweight="bold",
                color=txt_color,
                path_effects=[pe.withStroke(linewidth=2.3, foreground="#FFFFFF")],
                zorder=17 if tier <= 2 else 16,
            )
        placed_labels.append((nm, final_lbl, fs, is_callout, street_grid_codes[nm]))

        # Para vias muito longas (> 1.100 m), tentar um 2º rótulo na outra extremidade para leitura imediata
        if main_line.length > 1100.0:
            for fr2 in [0.16, 0.84, 0.22, 0.78]:
                x2, y2, ang2 = point_and_angle_along(main_line, fr2, span_m=45.0)
                xm2, ym2 = utm_to_mm(x2, y2)
                if not (6.0 < xm2 < MAP_W_MM - 6.0 and 6.0 < ym2 < MAP_H_MM - 6.0):
                    continue
                lbox2 = make_label_box_mm(xm2, ym2, short_lbl, fs, ang2)
                if not any(lbox2.intersects(ob) for ob in occupied_boxes):
                    occupied_boxes.append(lbox2)
                    ax.text(
                        x2, y2, short_lbl,
                        ha="center", va="center",
                        rotation=ang2,
                        rotation_mode="anchor",
                        fontsize=fs,
                        fontweight="bold",
                        color=txt_color,
                        path_effects=[pe.withStroke(linewidth=2.3, foreground="#FFFFFF")],
                        zorder=17,
                    )
                    break

    # -----------------------------------------------------------------------
    # Camada 7: Painel NE (Cabeçalho Institucional + Índice Completo das 47 Ruas)
    # -----------------------------------------------------------------------
    ne_left = (MAP_X0_MM + NE_X_MM) / PAGE_W_MM
    ne_bottom = (MAP_Y0_MM + NE_Y_MM) / PAGE_H_MM
    ax_ne = fig.add_axes([ne_left, ne_bottom, NE_W_MM / PAGE_W_MM, NE_H_MM / PAGE_H_MM], zorder=25)
    ax_ne.set_xlim(0, NE_W_MM)
    ax_ne.set_ylim(0, NE_H_MM)
    ax_ne.axis("off")

    # Fundo do card NE
    ax_ne.add_patch(FancyBboxPatch(
        (0.5, 0.5), NE_W_MM - 1.0, NE_H_MM - 1.0,
        boxstyle="round,pad=0.3,rounding_size=1.8",
        facecolor="#FFFFFF", edgecolor="#1E293B", linewidth=0.9, alpha=0.97,
    ))

    # Faixa superior azul-marinho institucional (compacta — 18 mm)
    HDR_H = 18.5
    ax_ne.add_patch(FancyBboxPatch(
        (0.5, NE_H_MM - HDR_H - 0.5), NE_W_MM - 1.0, HDR_H,
        boxstyle="round,pad=0.3,rounding_size=1.8",
        facecolor="#0F2942", edgecolor="#0F2942", linewidth=0.5,
    ))
    # Retângulo para deixar a borda inferior da faixa reta
    ax_ne.add_patch(Rectangle((0.2, NE_H_MM - HDR_H - 0.8), NE_W_MM - 0.4, 3.0, facecolor="#0F2942", edgecolor="none"))

    # Brasão de Pelotas (se disponível)
    tx_start = 3.5
    if BRASAO_PATH.exists():
        img = Image.open(BRASAO_PATH).convert("RGBA")
        # Fundo circular branco discreto atrás do brasão
        im_box = OffsetImage(img, zoom=0.055)
        ab = AnnotationBbox(im_box, (9.8, NE_H_MM - HDR_H / 2.0 - 0.6), frameon=False)
        ax_ne.add_artist(ab)
        tx_start = 18.5

    ax_ne.text(tx_start, NE_H_MM - 5.2, "MAPA TERRITORIAL — UBS DUNAS", fontsize=9.6, fontweight="bold", color="#FFFFFF", va="center")
    ax_ne.text(tx_start, NE_H_MM - 10.6, "Secretaria Municipal de Saúde de Pelotas/RS  •  Distrito Areal", fontsize=6.4, color="#CBD5E1", va="center")
    ax_ne.text(tx_start, NE_H_MM - 15.2, "Malha Viária Completa do Território de Abrangência  •  Área: 1,32 km²", fontsize=6.0, color="#FDE68A", fontweight="bold", va="center")

    # Título da lista de ruas
    list_top = NE_H_MM - HDR_H - 3.6
    all_sorted = sorted(named_streets.keys(), key=sort_key_pt)
    ax_ne.text(3.2, list_top, f"RELAÇÃO COMPLETA DE LOGRADOUROS ({len(all_sorted)} RUAS / AVENIDAS)", fontsize=6.8, fontweight="bold", color="#0F2942", va="center")
    ax_ne.text(NE_W_MM - 3.2, list_top, "Grade", fontsize=6.0, fontweight="bold", color="#64748B", va="center", ha="right")
    ax_ne.plot([2.8, NE_W_MM - 2.8], [list_top - 2.3, list_top - 2.3], color="#CBD5E1", lw=0.6)

    # 2 colunas perfeitamente alinhadas para as 47 ruas
    n_total = len(all_sorted)
    n_left = math.ceil(n_total / 2)
    y_start = list_top - 5.0
    y_end = 5.8
    row_step = (y_start - y_end) / max(1, n_left - 1)
    col_x = [3.0, NE_W_MM / 2.0 + 1.2]
    col_w = NE_W_MM / 2.0 - 4.2

    # Linha divisória vertical entre as 2 colunas
    ax_ne.plot([NE_W_MM / 2.0, NE_W_MM / 2.0], [4.5, list_top - 2.8], color="#E2E8F0", lw=0.5)

    for idx, nm in enumerate(all_sorted):
        col = 0 if idx < n_left else 1
        row = idx if col == 0 else idx - n_left
        yy = y_start - row * row_step
        xx = col_x[col]
        is_lim = named_streets[nm]["is_border"]
        code = street_grid_codes[nm]

        # Fundo zebrado bem sutil para facilitar leitura horizontal
        if row % 2 == 1:
            ax_ne.add_patch(Rectangle((xx - 0.8, yy - row_step * 0.46), col_w + 1.4, row_step * 0.92, facecolor="#F8FAFC", edgecolor="none"))

        disp_name = nm
        # Abreviações leves se nome extenso para manter folga confortável antes do código da grade
        if len(disp_name) > 27:
            disp_name = disp_name.replace("Avenida Engenheiro ", "Av. Eng. ")
            disp_name = disp_name.replace("Avenida Doutor ", "Av. Dr. ")
            disp_name = disp_name.replace("Avenida ", "Av. ")
            disp_name = disp_name.replace("Rua Jornalista ", "R. Jorn. ")
            disp_name = disp_name.replace("Rua Vereador ", "R. Ver. ")
            disp_name = disp_name.replace("Rua ", "R. ")
            disp_name = disp_name.replace("Passeio ", "Pass. ")
            disp_name = disp_name.replace("Travessa ", "Tv. ")
            disp_name = disp_name.replace("Loteamento ", "Lot. ")
        if is_lim:
            disp_name = f"{disp_name} *"

        ax_ne.text(
            xx, yy, disp_name,
            fontsize=5.7,
            fontweight="bold" if is_lim else "normal",
            color="#7C2D12" if is_lim else "#0F172A",
            va="center", ha="left",
        )
        ax_ne.text(
            xx + col_w, yy, code,
            fontsize=5.55, fontweight="bold", color="#334155",
            va="center", ha="right",
        )

    ax_ne.text(
        3.0, 2.3,
        "* Via limítrofe do território (divisa oficial conforme Mapa UBSs Pelotas).",
        fontsize=5.0, style="italic", color="#7C2D12", va="center",
    )

    # -----------------------------------------------------------------------
    # Camada 8: Painel SW (Legenda + Rosa dos Ventos + Escala Gráfica + Metadados)
    # -----------------------------------------------------------------------
    sw_left = (MAP_X0_MM + SW_X_MM) / PAGE_W_MM
    sw_bottom = (MAP_Y0_MM + SW_Y_MM) / PAGE_H_MM
    ax_sw = fig.add_axes([sw_left, sw_bottom, SW_W_MM / PAGE_W_MM, SW_H_MM / PAGE_H_MM], zorder=25)
    ax_sw.set_xlim(0, SW_W_MM)
    ax_sw.set_ylim(0, SW_H_MM)
    ax_sw.axis("off")

    ax_sw.add_patch(FancyBboxPatch(
        (0.5, 0.5), SW_W_MM - 1.0, SW_H_MM - 1.0,
        boxstyle="round,pad=0.3,rounding_size=1.8",
        facecolor="#FFFFFF", edgecolor="#1E293B", linewidth=0.9, alpha=0.97,
    ))

    ax_sw.text(3.2, SW_H_MM - 4.2, "LEGENDA E CONVENÇÕES", fontsize=7.0, fontweight="bold", color="#0F2942", va="center")
    ax_sw.plot([2.8, SW_W_MM - 2.8], [SW_H_MM - 6.6, SW_H_MM - 6.6], color="#CBD5E1", lw=0.6)

    # Itens da legenda (lado esquerdo do painel SW)
    leg_y = SW_H_MM - 10.5
    leg_step = 5.4
    # 1. Território UBS Dunas
    ax_sw.add_patch(Rectangle((3.2, leg_y - 1.5), 7.5, 3.0, facecolor="#FEFCE8", edgecolor="#B91C1C", lw=1.4))
    ax_sw.text(12.5, leg_y, "Território oficial da UBS Dunas", fontsize=5.8, fontweight="bold", color="#0F172A", va="center")

    # 2. UBS Dunas
    leg_y -= leg_step
    ax_sw.add_patch(Circle((6.95, leg_y), radius=1.9, facecolor="#DC2626", edgecolor="#FFFFFF", lw=0.7))
    ax_sw.text(6.95, leg_y, "+", ha="center", va="center", fontsize=6.5, fontweight="bold", color="#FFFFFF")
    ax_sw.text(12.5, leg_y, "UBS Dunas (ponto de referência)", fontsize=5.8, color="#0F172A", va="center")

    # 3. Via estrutural / limítrofe
    leg_y -= leg_step
    ax_sw.plot([3.2, 10.7], [leg_y, leg_y], color="#0F172A", lw=3.0, solid_capstyle="butt")
    ax_sw.plot([3.2, 10.7], [leg_y, leg_y], color="#FDE68A", lw=1.7, solid_capstyle="butt")
    ax_sw.text(12.5, leg_y, "Avenida estrutural / via limítrofe", fontsize=5.8, color="#0F172A", va="center")

    # 4. Via coletora
    leg_y -= leg_step
    ax_sw.plot([3.2, 10.7], [leg_y, leg_y], color="#1E293B", lw=2.3, solid_capstyle="butt")
    ax_sw.plot([3.2, 10.7], [leg_y, leg_y], color="#FFFBEB", lw=1.2, solid_capstyle="butt")
    ax_sw.text(12.5, leg_y, "Via coletora / principal interna", fontsize=5.8, color="#0F172A", va="center")

    # 5. Rua local / residencial
    leg_y -= leg_step
    ax_sw.plot([3.2, 10.7], [leg_y, leg_y], color="#475569", lw=1.8, solid_capstyle="butt")
    ax_sw.plot([3.2, 10.7], [leg_y, leg_y], color="#FFFFFF", lw=0.9, solid_capstyle="butt")
    ax_sw.text(12.5, leg_y, "Rua residencial / passeio", fontsize=5.8, color="#0F172A", va="center")

    # 6. Acesso / servidão & Equipamentos públicos
    leg_y -= leg_step
    ax_sw.plot([3.2, 10.7], [leg_y, leg_y], color="#94A3B8", lw=0.8)
    ax_sw.text(12.5, leg_y, "Acesso interno / via sem denominação", fontsize=5.6, color="#334155", va="center")

    leg_y -= leg_step
    ax_sw.add_patch(Rectangle((4.2, leg_y - 1.0), 2.0, 2.0, facecolor="#1D4ED8", edgecolor="#FFFFFF", lw=0.4))
    ax_sw.add_patch(Rectangle((7.2, leg_y - 1.0), 2.0, 2.0, facecolor="#047857", edgecolor="#FFFFFF", lw=0.4))
    ax_sw.text(12.5, leg_y, "Escolas (EMEF) / CRAS / Equipamentos", fontsize=5.6, color="#334155", va="center")

    # Rosa dos Ventos (canto superior direito do painel SW — com folga abaixo do divisor)
    nx, ny = SW_W_MM - 11.5, SW_H_MM - 21.8
    ax_sw.add_patch(Circle((nx, ny), radius=6.4, facecolor="#F8FAFC", edgecolor="#CBD5E1", lw=0.6))
    # Metade escura (leste) e clara (oeste) da seta Norte
    ax_sw.add_patch(MplPolygon([[nx, ny + 5.0], [nx + 2.3, ny - 3.4], [nx, ny - 1.6]], facecolor="#0F2942", edgecolor="#0F2942", lw=0.5))
    ax_sw.add_patch(MplPolygon([[nx, ny + 5.0], [nx - 2.3, ny - 3.4], [nx, ny - 1.6]], facecolor="#FFFFFF", edgecolor="#0F2942", lw=0.6))
    ax_sw.text(nx, ny + 8.3, "N", ha="center", va="center", fontsize=7.6, fontweight="bold", color="#0F2942")

    # Escala Gráfica exata na base do painel SW (500 m = 500/11500*1000 = 43.48 mm)
    bar_x0 = 3.5
    bar_y0 = 10.8
    bar_500_mm = (500.0 / SCALE) * 1000.0  # 43.48 mm
    seg_w = bar_500_mm / 4.0  # segmentos de 125 m
    for si in range(4):
        fc = "#0F2942" if si % 2 == 0 else "#FFFFFF"
        ax_sw.add_patch(Rectangle((bar_x0 + si * seg_w, bar_y0), seg_w, 2.1, facecolor=fc, edgecolor="#0F2942", lw=0.6))
    for lbl, pos in [("0", 0), ("250 m", 2), ("500 m", 4)]:
        ax_sw.text(bar_x0 + pos * seg_w, bar_y0 + 3.6, lbl, ha="center", va="center", fontsize=5.3, fontweight="bold", color="#0F172A")
    ax_sw.text(bar_x0 + bar_500_mm + 4.5, bar_y0 + 1.1, f"Escala 1 : {int(SCALE):,}".replace(",", "."), ha="left", va="center", fontsize=6.0, fontweight="bold", color="#0F2942")

    # Metadados cartográficos compactos (rodapé interno do painel SW)
    ax_sw.plot([2.8, SW_W_MM - 2.8], [8.2, 8.2], color="#E2E8F0", lw=0.5)
    ax_sw.text(3.2, 5.8, "SIRGAS 2000 / UTM 22S  •  Formato A4 Paisagem (imprimir em 100% / tamanho real)", fontsize=4.8, fontweight="bold", color="#334155", va="center")
    ax_sw.text(3.2, 2.9, "Fontes: Mapa Território UBSs Pelotas (SMS/Google My Maps) • OpenStreetMap (ODbL)", fontsize=4.6, color="#64748B", va="center")

    # Salvar PDF vetorial e PNG 300 DPI
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    fig.savefig(PDF_OUT, format="pdf")
    fig.savefig(PNG_OUT, format="png", dpi=300)
    plt.close(fig)

    return {
        "total_streets": len(all_sorted),
        "internal_streets": sum(1 for k in all_sorted if not named_streets[k]["is_border"]),
        "border_streets": sum(1 for k in all_sorted if named_streets[k]["is_border"]),
        "callouts": [x for x in placed_labels if x[3]],
        "streets": [(nm, street_grid_codes[nm], named_streets[nm]["is_border"]) for nm in all_sorted],
    }


def write_readme(stats: dict) -> None:
    rows_md = "\n".join(
        f"| {i:02d} | {nm} | {'Limítrofe' if is_b else 'Interna'} | `{code}` |"
        for i, (nm, code, is_b) in enumerate(stats["streets"], 1)
    )
    readme = f"""# Mapa Territorial — UBS Dunas (Distrito Areal • Pelotas/RS)

Mapa cartográfico exclusivo do território da **UBS Dunas**, contendo **todas as {stats['total_streets']} ruas e avenidas** ({stats['internal_streets']} internas + {stats['border_streets']} vias limítrofes), elaborado em **formato A4 Paisagem (`297 × 210 mm`)** com **margens mínimas (`3 mm`)** para máximo aproveitamento da área de impressão e alta legibilidade na distribuição às equipes.

## Arquivos disponíveis

| Arquivo | Formato | Uso recomendado |
|---------|---------|-----------------|
| [`Mapa_Territorio_UBS_Dunas_A4.pdf`](./Mapa_Territorio_UBS_Dunas_A4.pdf) | **PDF vetorial A4** (`1 : 11.500`) | **Impressão em cópias** (qualidade tipográfica máxima em impressoras laser/jato de tinta coloridas ou P&B) |
| [`Mapa_Territorio_UBS_Dunas_A4_300dpi.png`](./Mapa_Territorio_UBS_Dunas_A4_300dpi.png) | **PNG 300 DPI** (`3508 × 2480 px`) | Visualização rápida em tela, celular/WhatsApp ou inserção em documentos |
| [`ubs_dunas_territorio.geojson`](./ubs_dunas_territorio.geojson) | **GeoJSON (WGS 84)** | Uso em SIG (QGIS, Google Earth, mapas digitais) |

## Orientações de impressão (A4)

1. Abra [`Mapa_Territorio_UBS_Dunas_A4.pdf`](./Mapa_Territorio_UBS_Dunas_A4.pdf).
2. Configure a impressora para **Papel A4 — Orientação Paisagem**.
3. Em dimensionamento de página, selecione **Tamanho Real (100%)** ou **Ajustar às margens imprimíveis** (o mapa já utiliza margem de segurança mínima de `3 mm`, preservando o aproveitamento máximo da folha).
4. O traçado viário foi projetado com **alto contraste (contorno escuro + miolo claro)**, garantindo excelente legibilidade tanto em **impressão colorida** quanto em **fotocópias preto e branco**.

## Bases utilizadas

- **Limite oficial do território e ponto da UBS Dunas:** *Mapa Território UBSs Pelotas* (SMS Pelotas / Google My Maps — `mid=1D6SDvi-BRD-GIhiKeRsM6oYAgty-jDQ`).
- **Malha viária, toponímia e equipamentos públicos:** OpenStreetMap (ODbL), projetados em **SIRGAS 2000 / UTM Zona 22S (`EPSG:31982`)**.
- **Área do território:** `1,32 km²` (`132,13 ha`) | **Perímetro:** `8.023 m`.

## Relação completa de logradouros ({stats['total_streets']})

| # | Logradouro | Classificação | Grade |
|---:|------------|---------------|-------|
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
    parser = argparse.ArgumentParser(description="Gera mapa A4 do território da UBS Dunas (Pelotas/RS).")
    parser.add_argument("--refresh", action="store_true", help="Rebaixa KML e OSM da internet.")
    args = parser.parse_args()

    dataset = load_or_build_dataset(refresh=args.refresh)
    stats = render_map(dataset)
    write_readme(stats)

    print(f"OK PDF: {PDF_OUT} ({PDF_OUT.stat().st_size / 1024:.1f} KB)")
    print(f"OK PNG: {PNG_OUT} ({PNG_OUT.stat().st_size / 1024:.1f} KB)")
    print(f"OK GeoJSON: {GEOJSON_OUT} ({GEOJSON_OUT.stat().st_size / 1024:.1f} KB)")
    print(f"Logradouros rotulados: {stats['total_streets']} ({stats['internal_streets']} internas + {stats['border_streets']} limítrofes)")
    print(f"Callouts utilizados ({len(stats['callouts'])}):", [c[0] for c in stats["callouts"]])


if __name__ == "__main__":
    main()
