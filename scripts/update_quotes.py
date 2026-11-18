#!/usr/bin/env python3
"""
Обновление котировок из MOEX ISS (без ключей) → data/prices.json.

Каждый запуск:
1. Находит SECID на TQBR по ISIN из data/fundamentals.json.
2. Догружает дневную историю (официальные цены закрытия) начиная с последней
   сохранённой даты минус 10 дней; при пустом файле — всю историю с 2021 г.
   Последние дни перезаписываются официальными значениями.
3. Берёт текущую цену (LAST / LCURRENTPRICE) и записывает точку «сегодня»,
   если за сегодня ещё нет официального закрытия.
4. Пишет data/prices.json и data/status.json.

Запуск: python3 scripts/update_quotes.py [--full] [--dry-run]
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import datetime, timedelta, timezone, date
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen

MSK = timezone(timedelta(hours=3))
ROOT = Path(__file__).resolve().parent.parent
FUND = ROOT / "data" / "fundamentals.json"
PRICES = ROOT / "data" / "prices.json"
STATUS = ROOT / "data" / "status.json"
ISS = "https://iss.moex.com/iss"
HIST_START = "2021-01-01"

# Резервное соответствие ISIN → SECID, если поиск по ISIN не сработал
FALLBACK_SECID = {
    "RU000A103X66": "POSI", "RU000A106T36": "ASTR", "RU000A10CTQ0": "BAZA",
    "RU000A108ZR8": "DATA", "RU000A107ER5": "DIAS", "RU000A0ZZBC2": "SOFL",
}
INDEXES = ["IMOEX"]


def get(url: str, retries: int = 4) -> dict:
    err = None
    for i in range(retries):
        try:
            req = Request(url, headers={"User-Agent": "ru-it-dashboard/1.0"})
            with urlopen(req, timeout=30) as r:
                return json.loads(r.read())
        except (URLError, HTTPError, TimeoutError, json.JSONDecodeError) as e:
            err = e
            time.sleep(2 * (i + 1))
    raise RuntimeError(f"ISS недоступен: {url} — {err}")


def table(d: dict, name: str) -> list[dict]:
    blk = d.get(name) or {}
    cols = blk.get("columns", [])
    return [dict(zip(cols, row)) for row in blk.get("data", [])]


def board_snapshot() -> tuple[dict, dict]:
    """Все бумаги TQBR: securities и marketdata по SECID."""
    d = get(f"{ISS}/engines/stock/markets/shares/boards/TQBR/securities.json"
            "?iss.meta=off&iss.only=securities,marketdata")
    sec = {r["SECID"]: r for r in table(d, "securities")}
    md = {r["SECID"]: r for r in table(d, "marketdata")}
    return sec, md


def history(path: str, start: str) -> list[tuple[str, float, float | None]]:
    """[(date, close, value_mln)] — постранично через history.cursor."""
    out, pos = [], 0
    while True:
        d = get(f"{ISS}/history/{path}.json?iss.meta=off&from={start}&start={pos}"
                "&iss.only=history,history.cursor")
        rows = table(d, "history")
        for r in rows:
            px = r.get("LEGALCLOSEPRICE") or r.get("CLOSE")
            if px is None or r.get("TRADEDATE") is None:
                continue
            val = r.get("VALUE")
            out.append((r["TRADEDATE"], float(px), round(val / 1e6, 2) if val else None))
        cur = table(d, "history.cursor")
        if not cur:
            break
        c = cur[0]
        pos = c["INDEX"] + c["PAGESIZE"]
        if pos >= c["TOTAL"] or not rows:
            break
    return out


def merge(series: dict, rows: list[tuple]) -> None:
    """Сливает ряды по дате; новые значения перезаписывают старые."""
    m = {x: (y, v) for x, y, v in zip(series.get("x", []), series.get("y", []),
                                       series.get("v", [None] * len(series.get("x", []))))}
    for x, y, v in rows:
        m[x] = (y, v)
    xs = sorted(m)
    series["x"] = xs
    series["y"] = [m[x][0] for x in xs]
    series["v"] = [m[x][1] for x in xs]


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--full", action="store_true", help="перезагрузить всю историю")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    now = datetime.now(MSK)
    today = now.date().isoformat()
    fund = json.loads(FUND.read_text(encoding="utf-8"))
    P = {} if a.full or not PRICES.exists() else json.loads(PRICES.read_text(encoding="utf-8"))
    P.setdefault("tickers", {})
    P.setdefault("index", {})
    status = {"run_at": now.strftime("%Y-%m-%d %H:%M МСК"), "ok": [], "warn": []}

    sec, md = board_snapshot()
    by_isin = {r.get("ISIN"): sid for sid, r in sec.items()}

    for c in fund["companies"]:
        key, isin = c["key"], c["isin"]
        secid = by_isin.get(isin) or (FALLBACK_SECID.get(isin) if FALLBACK_SECID.get(isin) in sec else None)
        if not secid:
            status["warn"].append(f"{c['name']}: ISIN {isin} не найден на TQBR")
            continue
        if secid != by_isin.get(isin):
            status["warn"].append(f"{c['name']}: SECID {secid} взят из резервного списка")
        t = P["tickers"].setdefault(key, {"x": [], "y": [], "v": []})
        start = HIST_START
        if t["x"]:
            start = (date.fromisoformat(t["x"][-1]) - timedelta(days=10)).isoformat()
        rows = history(f"engines/stock/markets/shares/boards/TQBR/securities/{secid}", start)
        # убираем внутридневную точку «сегодня», если пришло официальное закрытие
        merge(t, rows)

        s, m = sec[secid], md.get(secid, {})
        last = m.get("LAST") or m.get("LCURRENTPRICE") or m.get("MARKETPRICE")
        prev = s.get("PREVPRICE")
        hist_today = bool(rows) and rows[-1][0] == today
        if last and not hist_today:
            merge(t, [(today, float(last), round(m["VALTODAY"] / 1e6, 2) if m.get("VALTODAY") else None)])
        px = float(last) if last else (t["y"][-1] if t["y"] else None)
        t.update({
            "secid": secid, "isin": isin,
            "issuesize": s.get("ISSUESIZE"),
            "last": px, "prev": prev,
            "chg": round((px / prev - 1) * 100, 2) if px and prev else None,
            "time": m.get("UPDATETIME") or m.get("TIME"),
            "tradedate": m.get("TRADEDATE") or (t["x"][-1] if t["x"] else None),
            "status": m.get("TRADINGSTATUS"),
        })
        status["ok"].append(f"{key}/{secid}: {px} ({len(t['x'])} дн.)")

    for ix in INDEXES:
        try:
            t = P["index"].setdefault(ix, {"x": [], "y": [], "v": []})
            start = HIST_START if not t["x"] else (date.fromisoformat(t["x"][-1]) - timedelta(days=10)).isoformat()
            rows = [(x, y, None) for x, y, _ in history(f"engines/stock/markets/index/boards/SNDX/securities/{ix}", start)]
            merge(t, rows)
            d = get(f"{ISS}/engines/stock/markets/index/boards/SNDX/securities/{ix}.json?iss.meta=off&iss.only=marketdata")
            mdx = table(d, "marketdata")
            cur = mdx[0].get("CURRENTVALUE") if mdx else None
            if cur and not (rows and rows[-1][0] == today):
                merge(t, [(today, float(cur), None)])
            t["last"] = cur or (t["y"][-1] if t["y"] else None)
            t["prev"] = mdx[0].get("LASTVALUE") if mdx else None
            status["ok"].append(f"{ix}: {t['last']}")
        except Exception as e:  # индекс не критичен
            status["warn"].append(f"{ix}: {e}")

    P["updated"] = now.strftime("%d.%m.%Y %H:%M МСК")
    P["updated_iso"] = now.isoformat(timespec="minutes")

    print(json.dumps(status, ensure_ascii=False, indent=1))
    if a.dry_run:
        return 0
    PRICES.write_text(json.dumps(P, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    STATUS.write_text(json.dumps(status, ensure_ascii=False, indent=1), encoding="utf-8")
    return 0 if status["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
