# -*- coding: utf-8 -*-
"""
Генерирует docs/data.json и копирует фото для статического сайта (GitHub Pages).
Переиспользует данные о товарах/клеймах/таксономии из build_report.py, чтобы
не дублировать вручную одни и те же списки в двух местах.
"""
import json
import os
import shutil
import sys
from collections import defaultdict

sys.path.insert(0, os.path.dirname(__file__))
import build_report as br  # noqa: E402  (импорт заодно пересоберёт xlsx — это ок)

PROJECT_DIR = br.PROJECT_DIR
DOCS_DIR = os.path.join(PROJECT_DIR, "docs")
DOCS_PHOTOS = os.path.join(DOCS_DIR, "photos")

os.makedirs(DOCS_PHOTOS, exist_ok=True)

# ---- копируем фото ----
for oid, (date, fname) in br.PHOTO_FILES.items():
    src = br.photo_path(oid)
    dst_dir = os.path.join(DOCS_PHOTOS, date)
    os.makedirs(dst_dir, exist_ok=True)
    dst = os.path.join(dst_dir, fname)
    shutil.copyfile(src, dst)

# ---- клеймы по товару ----
claims_by_oid = defaultdict(list)
names_by_code = {t[0]: t[1] for t in br.taxonomy}
for oid, pname, animal, ptype, code, cat_name, text in br.claim_rows:
    claims_by_oid[oid].append({"code": code, "category": cat_name, "text": text})

# ---- товары ----
products_out = []
for p in br.products:
    date, fname = br.PHOTO_FILES[p["oid"]]
    method = "sort" if date == "2026-07-31" else "diff"
    products_out.append({
        "oid": p["oid"],
        "date": date,
        "method": method,
        "animal": p["animal"],
        "ptype": p["ptype"],
        "brand": p["brand"],
        "name": p["name"],
        "url": p["url"],
        "photo": f"photos/{date}/{fname}",
        "claims_short": p["claims_short"],
        "claims": claims_by_oid.get(p["oid"], []),
        "protein": p["protein"],
        "ingredients": p["ingredients"],
        "plant": p["plant"],
    })

# ---- аналитика: количество товаров на категорию ----
n_products = len(br.products)
counts = {code: 0 for code in names_by_code}
for oid, claims in claims_by_oid.items():
    for code in {c["code"] for c in claims}:
        counts[code] += 1

analytics = [
    {"code": code, "name": names_by_code[code], "count": counts[code],
     "pct": round(100 * counts[code] / n_products)}
    for code in names_by_code
]
analytics.sort(key=lambda x: -x["count"])

taxonomy_out = [
    {"code": t[0], "name": t[1], "description": t[2], "examples": t[3]}
    for t in br.taxonomy
]

conclusions = [c for c in br.conclusions if c]

# ---- протеин: по ВСЕМ 49 подтверждённым SKU, не только по 12 разобранным ----
protein_counts = defaultdict(int)
for oid, tags in br.PROTEIN_TAGS.items():
    for tag in tags:
        protein_counts[tag] += 1
for oid, (_, tags) in br.EXTRA_PROTEIN_ONLY.items():
    for tag in tags:
        protein_counts[tag] += 1
protein_base = br.protein_known
protein_analytics = [
    {"name": tag, "count": c, "pct": round(100 * c / protein_base)}
    for tag, c in sorted(protein_counts.items(), key=lambda x: -x[1])
]

# ---- ингредиенты-клеймы (фасад): база — вся выборка (27) ----
ingredient_counts = defaultdict(int)
for oid, tags in br.INGREDIENT_TAGS.items():
    for tag in tags:
        ingredient_counts[tag] += 1
ingredient_analytics = [
    {"name": tag, "count": c, "pct": round(100 * c / n_products)}
    for tag, c in sorted(ingredient_counts.items(), key=lambda x: -x[1])
]

# ---- производитель: имя юрлица для РФ, страна для остальных; неизвестные исключены ----
producer_counts = defaultdict(int)
for oid, region in br.PLANT_REGION.items():
    if region:
        producer_counts[region] += 1
producer_base = br.region_known
producer_analytics = [
    {"name": tag, "count": c, "pct": round(100 * c / producer_base)}
    for tag, c in sorted(producer_counts.items(), key=lambda x: -x[1])
]

data = {
    "generated_at": "2026-08-03",
    "n_products": n_products,
    "days_tracked": 3,
    "history_ids_total": 159,
    "products": products_out,
    "taxonomy": taxonomy_out,
    "analytics": analytics,
    "protein_analytics": protein_analytics,
    "protein_base": protein_base,
    "protein_total_skus": br.PROTEIN_TOTAL_SKUS,
    "ingredient_analytics": ingredient_analytics,
    "ingredient_base": n_products,
    "producer_analytics": producer_analytics,
    "producer_base": producer_base,
    "conclusions": conclusions,
}

with open(os.path.join(DOCS_DIR, "data.json"), "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("site data.json written:", len(products_out), "products,", len(analytics), "categories")
