#!/usr/bin/env python3
"""
Конвертация Excel-базы (лист «Software comparables table») в data/fundamentals.json.

Берутся сохранённые (кешированные) значения ячеек, поэтому файл должен быть
сохранён в Excel после пересчёта. Формулы Cbonds и внешние ссылки не пересчитываются.

Использование:
    python3 scripts/build_fundamentals.py base/software_comps.xlsx [--check]

--check  — сверяет мультипликаторы, посчитанные из JSON, с блоком C4:J9 в Excel.
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, date
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "data" / "fundamentals.json"

YEARS = [2024, 2025, 2026, 2027, 2028, 2029, 2030]
YCOLS = ["C", "D", "E", "F", "G", "H", "I"]

# Расположение блоков в листе. Порядок — порядок вывода на дашборде.
# rows: строки выручки, EBITDAC, NIC; extra: дополнительная строка (отгрузки / оборот)
COMPANIES = [
    {"key": "POSI", "isin": "RU000A103X66", "name": "Позитив", "row13": 14,
     "rows": {"rev": 26, "ebitdac": 27, "nic": 29}, "extra": (24, "Отгрузки"),
     "source": "наш прогноз", "color": "#2a78d6", "mult_row": 6},
    {"key": "ASTR", "isin": "RU000A106T36", "name": "Астра", "row13": 15,
     "rows": {"rev": 75, "ebitdac": 76, "nic": 78}, "extra": None,
     "source": "наш прогноз", "color": "#eb6834", "mult_row": 7},
    {"key": "BAZA", "isin": "RU000A10CTQ0", "name": "Базис", "row13": 16,
     "rows": {"rev": 35, "ebitdac": 36, "nic": 38}, "extra": None,
     "source": "наш прогноз", "color": "#1baf7a", "mult_row": 8},
    {"key": "DATA", "isin": "RU000A108ZR8", "name": "Аренадата", "row13": 13,
     "rows": {"rev": 44, "ebitdac": 45, "nic": 47}, "extra": None,
     "source": "наш прогноз", "color": "#eda100", "mult_row": 4},
    {"key": "DIAS", "isin": "RU000A107ER5", "name": "Диасофт", "row13": 17,
     "rows": {"rev": 53, "ebitdac": 54, "nic": 56}, "extra": None,
     "source": "Сбер", "color": "#e87ba4", "mult_row": 9},
    {"key": "SOFL", "isin": "RU000A0ZZBC2", "name": "Софтлайн", "row13": 18,
     "rows": {"rev": 64, "ebitdac": 65, "nic": 67}, "extra": (62, "Оборот"),
     "source": "Альфа-Банк", "color": "#008300", "mult_row": 5},
]

# Ручные пометки о данных, требующих проверки (показываются значком ⚠ на дашборде)
FLAGS = {
    "SOFL": {"net_debt_bn": "Чистый долг 0,03 млрд руб (Cbonds) не согласуется с ЧД/EBITDAC 1,8x за 2025 "
                            "и 3,3x LTM в той же базе: вероятна ошибка единиц. Требует проверки — от него зависит EV."},
}

# Наш прогноз по Астре (млн руб, 2025–2030), заменяет консенсус брокеров в "fin".
# Консенсус сохраняется в поле "consensus" для сравнения. 2024 — факт из Excel-базы.
ASTRA_OWN = {
    "rev": [20369.24, 23610.037711885056, 27396.20137174215, 33324.548293828935, 39808.78915408612, 46934.07391795832],
    "ebitdac": [4433.811, 5624.063130260126, 7524.005068118789, 11063.155907762144, 14994.103075268074, 19280.500251583144],
    "nic": [3928.287, 5522.5750027538825, 7182.425048100133, 10532.555405924773, 14254.33451895691, 18313.148018976957],
}

# Прогнозы брокеров по Астре (из них строится консенсус). Эйлер — в млрд руб.
ASTRA_BROKERS = [
    ("Эйлер", {"rev": 84, "ebitdac": 85, "nic": 86}, 1000.0),
    ("SberCIB", {"rev": 91, "ebitdac": 92, "nic": 93}, 1.0),
    ("Sovcom", {"rev": 98, "ebitdac": 99, "nic": 100}, 1.0),
    ("D8", {"rev": 105, "ebitdac": 106, "nic": 107}, 1.0),
]


def num(v):
    if v is None or isinstance(v, str):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def series(ws, row, scale=1.0):
    out = {}
    for y, c in zip(YEARS, YCOLS):
        v = num(ws[f"{c}{row}"].value)
        out[str(y)] = None if v is None else round(v * scale, 3)
    return out


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--out", default=str(OUT))
    a = ap.parse_args()

    wb = openpyxl.load_workbook(a.xlsx, data_only=True)
    ws = wb["Software comparables table"] if "Software comparables table" in wb.sheetnames else wb.worksheets[0]

    asof = ws["B1"].value
    asof_s = asof.date().isoformat() if isinstance(asof, (datetime, date)) else None

    warnings: list[str] = []
    comps = []
    for c in COMPANIES:
        r = c["row13"]
        isin = ws[f"A{r}"].value
        if isin != c["isin"]:
            warnings.append(f"{c['name']}: ISIN в A{r} = {isin}, ожидался {c['isin']}")
        fin = {k: series(ws, row) for k, row in c["rows"].items()}
        item = {
            "key": c["key"], "isin": c["isin"], "name": c["name"], "color": c["color"],
            "source": c["source"],
            "shares_mn": num(ws[f"D{r}"].value),
            "net_debt_bn": num(ws[f"F{r}"].value),
            "xls_price": num(ws[f"C{r}"].value),
            "xls_mcap_bn": num(ws[f"E{r}"].value),
            "nd_ebitdac_xls": {"2025": num(ws[f"K{c['mult_row']}"].value),
                                "ltm_9m26": num(ws[f"L{c['mult_row']}"].value)},
            "fin": fin,
        }
        if c["key"] in FLAGS:
            item["flags"] = FLAGS[c["key"]]
        if c["extra"]:
            item["extra"] = {"label": c["extra"][1], "values": series(ws, c["extra"][0])}

        if c["key"] == "ASTR":
            brokers = {}
            for nm, rows, sc in ASTRA_BROKERS:
                brokers[nm] = {k: series(ws, row, sc) for k, row in rows.items()}
            item["brokers"] = brokers
            # Консенсус = среднее по доступным прогнозам (пустые ячейки не учитываются).
            # В Excel формула AVERAGE(I84*1000; …) превращает пустую ячейку Эйлера в 0,
            # из-за чего консенсус 2030 занижен. Пересчитываем и фиксируем расхождения.
            for k in ("rev", "ebitdac", "nic"):
                for y in map(str, YEARS):
                    if y in ("2024", "2025"):
                        continue  # 2024–2025: значения из Excel как есть
                    vals = [b[k][y] for b in brokers.values() if b[k][y] is not None]
                    if not vals:
                        continue
                    avg = round(sum(vals) / len(vals), 3)
                    xv = fin[k][y]
                    if xv is None or abs(avg - xv) > 0.5:
                        lab = {"rev": "выручка", "ebitdac": "EBITDAC", "nic": "NIC"}[k]
                        f0 = lambda v: "—" if v is None else f"{v:,.0f}".replace(",", " ")
                        warnings.append(
                            f"Астра, {lab} {y}: консенсус в Excel {f0(xv)} млн руб (формула AVERAGE считает пустую "
                            f"ячейку Эйлера нулём); использовано среднее {len(vals)} доступных прогнозов — {f0(avg)} млн руб")
                        fin[k][y] = avg
            item["consensus"] = json.loads(json.dumps(fin))
            for k, vals in ASTRA_OWN.items():
                for y, v in zip(range(2025, 2031), vals):
                    fin[k][str(y)] = round(v, 3)
            warnings = [w for w in warnings if not w.startswith("Астра")]
        comps.append(item)

    out = {
        "asof_excel": asof_s,
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "years": [str(y) for y in YEARS],
        "units": "млн руб; чистый долг — млрд руб; акции — млн шт.",
        "companies": comps,
        "warnings": warnings,
    }

    if a.check:
        bad = 0
        for c in COMPANIES:
            if c["key"] == "ASTR":
                continue  # мультипликаторы Excel посчитаны на консенсусе, а не на нашем прогнозе
            item = next(x for x in comps if x["key"] == c["key"])
            ev = item["xls_mcap_bn"] + item["net_debt_bn"]
            mc = item["xls_mcap_bn"]
            for j, y in enumerate(["2025", "2026", "2027", "2028"]):
                for col0, k, base in (("C", "ebitdac", ev), ("G", "nic", mc)):
                    col = chr(ord(col0) + j)
                    xv = ws[f"{col}{c['mult_row']}"].value
                    den = item["fin"][k][y]
                    mine = base * 1000 / den if den and den > 0 else None
                    if isinstance(xv, (int, float)) and mine is not None:
                        if abs(xv - mine) > 1e-6 * max(1, abs(xv)):
                            print(f"MISMATCH {c['name']} {k} {y}: excel {xv:.4f} vs {mine:.4f}")
                            bad += 1
                    else:
                        print(f"  {c['name']:10s} {k:7s} {y}: excel={xv!r:>8} calc={'отр.' if mine is None else round(mine,2)}")
        print("check:", "OK" if bad == 0 else f"{bad} mismatches")

    Path(a.out).parent.mkdir(parents=True, exist_ok=True)
    Path(a.out).write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    for w in warnings:
        print("WARN:", w)
    print("written", a.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
