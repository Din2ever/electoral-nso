"""
Электоральная карта НСО.
Источник данных — один Excel-файл с тремя листами:
  • «УИК»   — результаты по каждому участку
  • «ТИК»   — агрегаты по территориальным комиссиям (районы/города)
  • «ОблИК» — агрегат по всей области

Запуск: python server.py
Открыть: http://localhost:8000
"""
from pathlib import Path
from typing import Optional, List, Dict, Any

import pandas as pd
import numpy as np
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import HTMLResponse
from fastapi.middleware.cors import CORSMiddleware
from sklearn.cluster import KMeans

BASE_DIR = Path(__file__).parent
EXCEL_FILE = BASE_DIR / "база результатов ГД и ЗС.xlsx"

PARTY_NAMES = {
    "ER": "Единая Россия",
    "LDPR": "ЛДПР",
    "SRZP": "Справедливая Россия",
    "KPRF": "КПРФ",
    "NEWPEOPLE": "Новые люди",
}
def _normalize_pct(x):
    """
    Приводит значение из базы к «человеческому» проценту:
      • 0.5191  → 51.91
      • 0.51910000000000006 → 51.91
      • 51.91   → 51.91
      • None    → None
    То есть если число < 2 — считаем, что это доля (0..1), умножаем на 100.
    Иначе считаем, что это уже процент (0..100).
    Всегда округляем до 2 знаков.
    """
    if x is None:
        return None
    try:
        v = float(x)
    except (ValueError, TypeError):
        return None
    if v <= 2:                      # доля 0..1 (с запасом до 2)
        v = v * 100.0
    return round(v, 2)
PARTY_SHORT = {"ER": "ЕР", "LDPR": "ЛДПР", "SRZP": "СРЗП",
               "KPRF": "КПРФ", "NEWPEOPLE": "НЛ"}
PARTY_COLORS = {"ER": "#1e6bb8", "LDPR": "#0a3d91", "SRZP": "#f5a623",
                "KPRF": "#c1121f", "NEWPEOPLE": "#0b8a3a"}
PARTY_ORDER = ["ER", "LDPR", "SRZP", "KPRF", "NEWPEOPLE"]

# Коды выборов = уникальные значения столбца A
ELECTIONS: Dict[str, Dict[str, Any]] = {}


def _parse_pct(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).replace("%", "").replace(",", ".").strip()
    try:
        return float(s)
    except ValueError:
        return None


def _parse_int(x):
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return None
    try:
        return int(float(str(x).replace(" ", "").replace(",", ".")))
    except (ValueError, TypeError):
        return None


def _party_dict(votes, pct):
    return {"votes": votes, "pct": pct}


def load_uik_sheet(xl: pd.ExcelFile) -> List[dict]:
    """Лист «УИК». Строки — участки."""
    df = xl.parse("УИК")
    records = []
    for _, row in df.iterrows():
        code = row.iloc[0]
        year = row.iloc[1]
        territory = row.iloc[2]
        uik = row.iloc[4]
        if pd.isna(code) or pd.isna(uik):
            continue
        try:
            uik_num = int(float(uik))
        except (ValueError, TypeError):
            continue
        rec = {
            "level": str(code).strip(),
            "year": _parse_int(year),
            "territory": str(territory).strip() if not pd.isna(territory) else None,
            "uik": uik_num,
            "turnout_abs": _parse_int(row.iloc[5]),
            "turnout_pct": _parse_pct(row.iloc[6]),
            "voters": _parse_int(row.iloc[7]),
            "parties": {
                "ER": _party_dict(_parse_int(row.iloc[8]), _parse_pct(row.iloc[9])),
                "LDPR": _party_dict(_parse_int(row.iloc[10]), _parse_pct(row.iloc[11])),
                "SRZP": _party_dict(_parse_int(row.iloc[12]), _parse_pct(row.iloc[13])),
                "KPRF": _party_dict(_parse_int(row.iloc[14]), _parse_pct(row.iloc[15])),
                "NEWPEOPLE": _party_dict(_parse_int(row.iloc[16]), _parse_pct(row.iloc[17])),
            },
        }
        records.append(rec)
    return records


def load_tik_sheet(xl: pd.ExcelFile) -> List[dict]:
    """Лист «ТИК». Строки — территории."""
    df = xl.parse("ТИК")
    records = []
    for _, row in df.iterrows():
        code = row.iloc[0]
        year = row.iloc[1]
        territory = row.iloc[2]
        if pd.isna(code) or pd.isna(territory):
            continue
        rec = {
            "level": str(code).strip(),
            "year": _parse_int(year),
            "territory": str(territory).strip(),
            "uik_count": _parse_int(row.iloc[3]),
            "turnout_abs": _parse_int(row.iloc[5]),
            "turnout_pct": _parse_pct(row.iloc[6]),
            "voters": _parse_int(row.iloc[7]),
            "parties": {
                "ER": _party_dict(_parse_int(row.iloc[8]), _parse_pct(row.iloc[9])),
                "LDPR": _party_dict(_parse_int(row.iloc[10]), _parse_pct(row.iloc[11])),
                "SRZP": _party_dict(_parse_int(row.iloc[12]), _parse_pct(row.iloc[13])),
                "KPRF": _party_dict(_parse_int(row.iloc[14]), _parse_pct(row.iloc[15])),
                "NEWPEOPLE": _party_dict(_parse_int(row.iloc[16]), _parse_pct(row.iloc[17])),
            },
        }
        records.append(rec)
    return records


def load_oblik_sheet(xl: pd.ExcelFile) -> List[dict]:
    """Лист «ОблИК». Одна строка на выборы (вся область)."""
    df = xl.parse("ОблИК")
    records = []
    for _, row in df.iterrows():
        code = row.iloc[0]
        if pd.isna(code):
            continue
        rec = {
            "level": str(code).strip(),
            "year": _parse_int(row.iloc[1]),
            "territory": str(row.iloc[2]).strip(),
            "turnout_abs": _parse_int(row.iloc[5]),
            "turnout_pct": _parse_pct(row.iloc[6]),
            "voters": _parse_int(row.iloc[7]),
            "parties": {
                "ER": _party_dict(_parse_int(row.iloc[8]), _parse_pct(row.iloc[9])),
                "LDPR": _party_dict(_parse_int(row.iloc[10]), _parse_pct(row.iloc[11])),
                "SRZP": _party_dict(_parse_int(row.iloc[12]), _parse_pct(row.iloc[13])),
                "KPRF": _party_dict(_parse_int(row.iloc[14]), _parse_pct(row.iloc[15])),
                "NEWPEOPLE": _party_dict(_parse_int(row.iloc[16]), _parse_pct(row.iloc[17])),
            },
        }
        records.append(rec)
    return records


print("Загрузка Excel...")
if not EXCEL_FILE.exists():
    raise FileNotFoundError(f"Не найден файл: {EXCEL_FILE}")

xl = pd.ExcelFile(EXCEL_FILE)
UIK_DATA = load_uik_sheet(xl)
TIK_DATA = load_tik_sheet(xl)
OBLIK_DATA = load_oblik_sheet(xl)
print(f"  УИК:   {len(UIK_DATA)}")
print(f"  ТИК:   {len(TIK_DATA)}")
print(f"  ОблИК: {len(OBLIK_DATA)}")

# ---- Реестр выборов ----
# (код, год) -> {code, year, title, date}
def _election_key(level: str, year: int) -> str:
    return f"{level}|{year}"

for r in TIK_DATA + OBLIK_DATA + UIK_DATA:
    if not r.get("year"):
        continue
    key = _election_key(r["level"], r["year"])
    if key not in ELECTIONS:
        ELECTIONS[key] = {
            "code": key,
            "level": r["level"],
            "year": r["year"],
            "title": f'{r["level"]} — {r["year"]}',
            "date": f'{r["year"]}-01-01',
        }

ELECTIONS = dict(sorted(ELECTIONS.items(), key=lambda kv: (kv[1]["year"], kv[1]["level"])))


def find_code(code: str) -> dict:
    if code not in ELECTIONS:
        raise HTTPException(404, f"Неизвестный код выборов: {code}")
    return ELECTIONS[code]


def get_tik(code: str) -> List[dict]:
    """Только лист ТИК."""
    meta = find_code(code)
    return [r for r in TIK_DATA
            if r["level"] == meta["level"] and r["year"] == meta["year"]]


def get_uik(code: str, territory: Optional[str] = None) -> List[dict]:
    """Только лист УИК."""
    meta = find_code(code)
    rows = [r for r in UIK_DATA
            if r["level"] == meta["level"] and r["year"] == meta["year"]]
    if territory:
        rows = [r for r in rows if r["territory"] == territory]
    return rows


def get_oblik(code: str) -> Optional[dict]:
    """Только лист ОблИК."""
    meta = find_code(code)
    for r in OBLIK_DATA:
        if r["level"] == meta["level"] and r["year"] == meta["year"]:
            return r
    return None


# ---------- FastAPI ----------
app = FastAPI(title="Электоральная карта НСО")
app.add_middleware(CORSMiddleware, allow_origins=["*"],
                   allow_methods=["*"], allow_headers=["*"])


@app.get("/api/elections")
def list_elections():
    return [{"code": v["code"], "title": v["title"],
             "date": v["date"], "year": v["year"], "level": v["level"]}
            for v in ELECTIONS.values()]


@app.get("/api/territories/{code}")
def territories(code: str):
    """Список территорий — из ТИК (для фильтров)."""
    return sorted({r["territory"] for r in get_tik(code)})


# ---------- /api/multi-results ----------
@app.get("/api/multi-results")
def multi_results(codes: str = Query(...),
                  territory: Optional[str] = None,
                  source: str = Query("TIK", pattern="^(TIK|UIK)$")):
    """
    Свод по территориям для выбранных выборов.
    source = "TIK"  → строки-территории (для страницы «Таблица»)
    source = "UIK"  → строки-участки (для режима «один район»)
    """
    code_list = [c.strip() for c in codes.split(",") if c.strip()]
    out: Dict[str, Any] = {}
    for c in code_list:
        if c not in ELECTIONS:
            continue
        if source == "UIK":
            rows = get_uik(c, territory)  # только УИК
            # агрегат по выбранному району (сумма по УИК)
            agg = _aggregate_rows(rows)
            out[c] = [{"territory": territory or "—", **agg}]
        else:
            rows = get_tik(c)  # только ТИК
            out[c] = [
                {
                    "territory": r["territory"],
                    "turnout_abs": r["turnout_abs"],
                    "turnout_pct": r["turnout_pct"],
                    "voters": r["voters"],
                    "parties": [
                        {"party_code": pc, "party_name": PARTY_NAMES[pc],
                         "color": PARTY_COLORS[pc],
                         "votes": r["parties"][pc]["votes"],
                         "pct": r["parties"][pc]["pct"]}
                        for pc in PARTY_ORDER
                    ],
                } for r in rows
            ]
    return out

def _aggregate_rows(rows: List[dict]) -> dict:
    if not rows:
        return {"turnout_abs": 0, "turnout_pct": 0, "voters": 0,
                "parties": [{"party_code": pc, "party_name": PARTY_NAMES[pc],
                             "color": PARTY_COLORS[pc], "votes": 0, "pct": 0}
                            for pc in PARTY_ORDER]}
    t_abs = sum(r["turnout_abs"] or 0 for r in rows)
    voters = sum(r["voters"] or 0 for r in rows)
    totals = {pc: 0 for pc in PARTY_ORDER}
    pct_sums = {pc: 0.0 for pc in PARTY_ORDER}
    pct_n = {pc: 0 for pc in PARTY_ORDER}
    for r in rows:
        for pc in PARTY_ORDER:
            totals[pc] += r["parties"][pc]["votes"] or 0
            p = _normalize_pct(r["parties"][pc]["pct"])
            if p is not None:
                pct_sums[pc] += p
                pct_n[pc] += 1
    return {
        "turnout_abs": t_abs,
        "turnout_pct": round(t_abs / voters * 100, 2) if voters else None,
        "voters": voters,
        "parties": [
            {"party_code": pc, "party_name": PARTY_NAMES[pc],
             "color": PARTY_COLORS[pc],
             "votes": totals[pc],
             "pct": round(pct_sums[pc] / pct_n[pc], 2) if pct_n[pc] else None}
            for pc in PARTY_ORDER
        ],
    }



# ---------- /api/results/{code} (агрегат по ТИК или ОблИК) ----------
@app.get("/api/results/{code}")
def results(code: str, territory: Optional[str] = None):
    if territory:
        rows = get_uik(code, territory)
        return [_aggregate_rows(rows)]
    rows = get_tik(code)
    return [
        {
            "territory": r["territory"],
            "turnout_abs": r["turnout_abs"],
            "turnout_pct": r["turnout_pct"],
            "voters": r["voters"],
            "parties": [
                {"party_code": pc, "party_name": PARTY_NAMES[pc],
                 "color": PARTY_COLORS[pc],
                 "votes": r["parties"][pc]["votes"],
                 "pct": r["parties"][pc]["pct"]}
                for pc in PARTY_ORDER
            ],
        } for r in rows
    ]


# ---------- /api/uik-list/{code}/{territory} ----------
@app.get("/api/uik-list/{code}")
def uik_list(code: str, territory: str):
    """Все УИК конкретного района — только лист УИК."""
    rows = get_uik(code, territory)
    return [
        {
            "uik": r["uik"],
            "turnout_abs": r["turnout_abs"],
            "turnout_pct": r["turnout_pct"],
            "voters": r["voters"],
            "parties": [
                {"party_code": pc, "party_name": PARTY_NAMES[pc],
                 "color": PARTY_COLORS[pc],
                 "votes": r["parties"][pc]["votes"],
                 "pct": r["parties"][pc]["pct"]}
                for pc in PARTY_ORDER
            ],
        } for r in rows
    ]


# ---------- /api/dynamics (ОблИК или ТИК) ----------
@app.get("/api/dynamics")
def dynamics(territory: Optional[str] = None):
    """
    territory == None → источник ОблИК (вся область).
    territory != None → источник ТИК (конкретный район).
    Проценты берутся из базы (столбцы J, L, N, P, R),
    нормализуются к 0..100 и округляются до 2 знаков.
    """
    out = []
    for code, meta in ELECTIONS.items():
        if territory:
            rows = [r for r in get_tik(code) if r["territory"] == territory]
            if not rows:
                continue
            r = rows[0]
        else:
            r = get_oblik(code)
            if not r:
                continue
        for pc in PARTY_ORDER:
            votes = r["parties"][pc]["votes"] or 0
            pct = _normalize_pct(r["parties"][pc]["pct"])
            out.append({
                "election_code": code,
                "election_date": str(meta["year"]),
                "party_code": pc,
                "party_name": PARTY_NAMES[pc],
                "color": PARTY_COLORS[pc],
                "votes": votes,
                "pct": pct,
            })
    out.sort(key=lambda x: (x["election_date"], x["party_code"]))
    return out
@app.get("/api/dynamics-by-territory")
def dynamics_by_territory(party: str = "ER", territories: Optional[str] = None):
    """Динамика одной партии по нескольким районам — только лист ТИК."""
    terr_list = [t.strip() for t in (territories or "").split(",") if t.strip()]
    out = []
    for code, meta in ELECTIONS.items():
        for r in get_tik(code):
            if terr_list and r["territory"] not in terr_list:
                continue
            p = r["parties"].get(party, {})
            out.append({
                "election_code": code,
                "election_date": str(meta["year"]),
                "territory": r["territory"],
                "votes": p.get("votes") or 0,
                "pct": p.get("pct"),
            })
    out.sort(key=lambda x: (x["territory"], x["election_date"]))
    return out


# ---------- /api/turnout-dynamics (ОблИК или ТИК) ----------
# ---------- /api/turnout-dynamics (ОблИК или ТИК) ----------
@app.get("/api/turnout-dynamics")
def turnout_dynamics(territory: Optional[str] = None):
    out = []
    for code, meta in ELECTIONS.items():
        if territory:
            rows = [r for r in get_tik(code) if r["territory"] == territory]
            if not rows:
                continue
            r = rows[0]
            out.append({
                "election_code": code,
                "election_date": str(meta["year"]),
                "turnout_pct": _normalize_pct(r["turnout_pct"]),
                "turnout_abs": r["turnout_abs"],
                "voters": r["voters"],
            })
        else:
            r = get_oblik(code)
            if not r:
                continue
            out.append({
                "election_code": code,
                "election_date": str(meta["year"]),
                "turnout_pct": _normalize_pct(r["turnout_pct"]),
                "turnout_abs": r["turnout_abs"],
                "voters": r["voters"],
            })
    out.sort(key=lambda x: x["election_date"])
    return out


# ---------- /api/clusters (ТИК, 4 фикс. кластера по явке × партии) ----------
@app.get("/api/clusters/{code}")
def clusters(code: str, party: str = "KPRF", k: int = 4):
    """
    Кластеры: ось X = явка %, ось Y = поддержка выбранной партии %.
    Источник — лист ТИК.
    """
    rows = get_tik(code)
    if not rows:
        return []
    points = []
    for r in rows:
        p = r["parties"].get(party) or {}
        points.append({
            "territory": r["territory"],
            "turnout": r["turnout_pct"],
            "support": p.get("pct"),
            "votes": p.get("votes"),
        })
    valid = [p for p in points if p["turnout"] is not None and p["support"] is not None]
    if len(valid) < 4:
        return [{"territory": p["territory"], "turnout": p["turnout"],
                 "support": p["support"], "votes": p["votes"],
                 "cluster": 0, "cluster_label": "—",
                 "cluster_color": "#888"} for p in points]

    X = np.array([[p["turnout"], p["support"]] for p in valid])
    km = KMeans(n_clusters=4, n_init=10, random_state=42)
    labels = km.fit_predict(X)

    # Центры кластеров → подписи
    centers = km.cluster_centers_
    med_x = float(np.median(X[:, 0]))
    med_y = float(np.median(X[:, 1]))

    CLUSTER_META = {
        "low_low":   {"label": "🟦 Низкая явка + низкая поддержка", "color": "#1e6bb8"},
        "low_high":  {"label": "🟩 Низкая явка + высокая поддержка", "color": "#0b8a3a"},
        "high_low":  {"label": "🟨 Высокая явка + низкая поддержка", "color": "#f5a623"},
        "high_high": {"label": "🟥 Высокая явка + высокая поддержка", "color": "#c1121f"},
    }

    def bucket(cx, cy):
        low_x = cx < med_x
        low_y = cy < med_y
        if low_x and low_y: return "low_low"
        if low_x and not low_y: return "low_high"
        if not low_x and low_y: return "high_low"
        return "high_high"

    cluster_meta = {}
    for ci, (cx, cy) in enumerate(centers):
        key = bucket(cx, cy)
        # на всякий случай — если два центра попали в один бакет,
        # разводим по значениям
        while key in cluster_meta.values():
            key = ("low_high" if key == "low_low" else
                   "high_low" if key == "low_high" else
                   "high_high" if key == "high_low" else "low_low")
        cluster_meta[ci] = key

    out = []
    for p, lbl in zip(valid, labels):
        key = cluster_meta[int(lbl)]
        meta = CLUSTER_META[key]
        out.append({
            "territory": p["territory"],
            "turnout": p["turnout"],
            "support": p["support"],
            "votes": p["votes"],
            "cluster": int(lbl),
            "cluster_key": key,
            "cluster_label": meta["label"],
            "cluster_color": meta["color"],
        })
    # добавим точки без данных — нейтральные
    for p in points:
        if p["turnout"] is None or p["support"] is None:
            out.append({**p, "cluster": -1, "cluster_key": "na",
                        "cluster_label": "Нет данных", "cluster_color": "#999"})
    return out


# ---------- /api/sankey (ОблИК) ----------
@app.get("/api/sankey")
def sankey(from_code: str, to_code: str):
    """Переток голосов между выборами — источник ОблИК."""
    a = get_oblik(from_code)
    b = get_oblik(to_code)
    if not a or not b:
        raise HTTPException(404, "Нет данных ОблИК для выбранных выборов")

    nodes = ([f'{PARTY_SHORT[pc]} ({a["year"]})' for pc in PARTY_ORDER] +
             [f'{PARTY_SHORT[pc]} ({b["year"]})' for pc in PARTY_ORDER])
    links = []
    for src in PARTY_ORDER:
        sv = a["parties"][src]["votes"] or 0
        for dst in PARTY_ORDER:
            dv = b["parties"][dst]["votes"] or 0
            val = int(min(sv, dv))
            if val <= 0:
                continue
            links.append({
                "source": f'{PARTY_SHORT[src]} ({a["year"]})',
                "target": f'{PARTY_SHORT[dst]} ({b["year"]})',
                "value": max(1, val // 1000),
                "real_value": val,
                "source_party": src,
                "target_party": dst,
            })
    return {"nodes": nodes, "links": links,
            "from_year": a["year"], "to_year": b["year"]}


# ---------- /api/top10 (ТИК) ----------
@app.get("/api/top10/{code}/{party}")
def top10(code: str, party: str, by: str = "pct"):
    rows = get_tik(code)  # только ТИК
    rows = [r for r in rows if r["parties"].get(party, {}).get(by) is not None]
    rows.sort(key=lambda r: r["parties"][party][by], reverse=True)
    top = rows[:10]
    return [
        {
            "territory": r["territory"],
            "parties": [
                {"party_code": pc, "party_name": PARTY_NAMES[pc],
                 "color": PARTY_COLORS[pc],
                 "votes": r["parties"][pc]["votes"],
                 "pct": r["parties"][pc]["pct"]}
                for pc in PARTY_ORDER
            ],
        } for r in top
    ]


# ---------- /api/heatmap-data ----------
@app.get("/api/heatmap-data")
def heatmap_data(codes: str = Query(...),
                 territory: Optional[str] = None):
    """
    territory == None → все районы, источник ТИК.
    territory != None → все УИК района, источник УИК.
    """
    code_list = [c.strip() for c in codes.split(",") if c.strip()]
    if territory:
        # режим «район» — строки: УИК
        uik_set = set()
        data: Dict[str, Dict[str, Optional[float]]] = {}
        for c in code_list:
            if c not in ELECTIONS:
                continue
            for r in get_uik(c, territory):
                key = f'УИК {r["uik"]}'
                uik_set.add(key)
                data.setdefault(key, {})[c] = _normalize_pct(r["turnout_pct"])
        return {"rows": sorted(uik_set),
                "elections": code_list,
                "data": data,
                "mode": "uik"}
    else:
        # режим «вся область» — строки: районы, источник ТИК
        terr_set = set()
        data = {}
        for c in code_list:
            if c not in ELECTIONS:
                continue
            for r in get_tik(c):
                terr_set.add(r["territory"])
                data.setdefault(r["territory"], {})[c] = _normalize_pct(r["turnout_pct"])
        return {"rows": sorted(terr_set),
                "elections": code_list,
                "data": data,
                "mode": "tik"}


# ---------- /api/radar-data (ОблИК или ТИК) ----------
@app.get("/api/radar-data")
def radar_data(codes: str = Query(...),
               territory: Optional[str] = None):
    """
    Источник данных:
      • territory == None → лист ОблИК (вся область);
      • territory != None → лист ТИК (конкретный район).
    Проценты берутся из базы (столбцы J, L, N, P, R),
    а не пересчитываются из голосов.
    """
    code_list = [c.strip() for c in codes.split(",") if c.strip()]
    out = []
    for c in code_list:
        if c not in ELECTIONS:
            continue
        if territory:
            rows = [r for r in get_tik(c) if r["territory"] == territory]
            if not rows:
                continue
            r = rows[0]
            label = f'{territory} · {ELECTIONS[c]["year"]}'
            source = "ТИК"
        else:
            r = get_oblik(c)
            if not r:
                continue
            label = f'Область · {ELECTIONS[c]["year"]}'
            source = "ОблИК"

        out.append({
            "code": c,
            "label": label,
            "source": source,
            # значения — прямо из базы, без пересчёта
            "values": [
                _normalize_pct(r["parties"][pc]["pct"]) or 0
                for pc in PARTY_ORDER
            ],
        })

    return {
        "indicators": [PARTY_SHORT[pc] for pc in PARTY_ORDER],
        "series": out,
    }


@app.get("/", response_class=HTMLResponse)
def index():
    return (BASE_DIR / "index.html").read_text(encoding="utf-8")


if __name__ == "__main__":
    import uvicorn
    import os
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))