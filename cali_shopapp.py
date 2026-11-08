"""
Cali Off-Road wheel prices on shop.app.

Reads cali/cali_catalog.json (built from the TWG dealer pages), searches shop.app,
matches each competitor listing to a Cali part number, and writes:
  cali_cache/shopapp_hits.json   every Cali wheel listing found (raw)
  cali_prices.csv                every matched listing, per-wheel price
  cali_best.csv                  cheapest competitor per part number (+ store, link)

Matching: part number in the title/handle first. If there is none, the listing must
agree on model number, diameter x width, bolt pattern and offset, and the finish
must not conflict.

Usage:  python cali_shopapp.py            (search + match)
        python cali_shopapp.py --skip-search   (re-match cached results)
"""
import argparse
import csv
import json
import os
import re
from collections import defaultdict

import amp_price_scraper as base

base.set_brand("cali")
CATALOG = os.path.join("cali", "cali_catalog.json")

MODELS = {  # model number -> names stores use
    "9110": ["Summit"], "9110D": ["Summit Dually"], "9114": ["Purge"], "9115": ["Invader"],
    "9115D": ["Invader Dually"], "9116": ["Vertex"], "9117": ["Auburn"], "9118": ["Crusher"],
    "9119": ["Evoke"], "9111": ["Sevenfold"], "9112": ["Gemini"],
}
FINISH_WORDS = {
    "polished": ["polish"], "chrome": ["chrome"], "matte black": ["matte"],
    "gloss black milled": ["milled"], "gloss black - milled spokes": ["milled"],
    "black milled": ["milled"], "brushed milled": ["brushed"], "gloss black": ["gloss"],
}


def norm(s):
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def load_catalog():
    cat = json.load(open(CATALOG))
    for c in cat:
        c["_sku"] = norm(c["sku"])
    return cat


def build_queries(cat):
    q = ["Cali Off-Road wheels", "Cali Offroad wheels", "Cali Off Road wheel"]
    for no, names in MODELS.items():
        for n in names:
            q += [f"Cali Off-Road {n}", f"Cali Offroad {n} {no}"]
            for d in sorted({c["diameter"] for c in cat if c["model_no"] == no}):
                q.append(f"Cali Off-Road {n} {d}")
    seen, out = set(), []
    for x in q:
        if x.lower() not in seen:
            seen.add(x.lower())
            out.append(x)
    return out


def looks_like_cali(title, handle=""):
    b = f"{title} {handle}".lower().replace("-", " ")
    if "cali" not in b:
        return False
    if any(w in b for w in ["tire and wheel", "wheel and tire", "wheel & tire", "package", "lug nut", "center cap",
                            "tpms", "spacer"]):
        return False
    return True


def parse_wheel(text):
    t = (text or "").lower().replace("-", " - ")
    t = re.sub(r"\s+", " ", t)
    out = {}
    m = re.search(r"(?<!\d)(17|18|20|22|24|26|28)(?:\.0)?\s*[x×]\s*(\d{1,2}(?:\.\d{1,2})?)(?!\d)", t)
    if m:
        out["diameter"], out["width"] = m.group(1), str(float(m.group(2))).rstrip("0").rstrip(".")
    m = re.search(r"(?<![\d.])([5-8])\s*(?:x|-|on|lug)?\s*(1[2-3]\d(?:\.\d)?|1[4-9]\d(?:\.\d)?|2[01]\d|5\.5|4\.5|6\.5)(?![\d])", t)
    if m:
        out["bolt"] = f"{m.group(1)}-{m.group(2)}"
    m = re.search(r"(?:et|offset)?\s*([+-]?\s?\d{1,3})\s*mm", t)
    if m:
        out["offset"] = m.group(1).replace(" ", "").replace("+", "")
    return out


def sku_match(blob, cat_by_sku):
    b = norm(blob)
    hits = [s for s in cat_by_sku if len(s) >= 8 and s in b]
    return max(hits, key=len) if hits else None


def spec_match(title, handle, cat):
    blob = f"{title} {handle.replace('-', ' ')}"
    w = parse_wheel(blob)
    low = blob.lower()
    nos = [no for no in MODELS if re.search(rf"\b{no.lower()}\b", low)]
    names = [no for no, ns in MODELS.items() if any(n.lower() in low for n in ns)]
    model_nos = set(nos or names)
    if "dually" in low:
        model_nos = {n for n in model_nos if n.endswith("D")} or {n + "D" for n in model_nos if n + "D" in MODELS}
    else:
        model_nos = {n for n in model_nos if not n.endswith("D")}
    if not model_nos or "diameter" not in w:
        return None
    cands = [c for c in cat if c["model_no"] in model_nos and c["diameter"] == w["diameter"]
             and str(float(c["width"])).rstrip("0").rstrip(".") == w.get("width")]
    if "bolt" in w:
        cands = [c for c in cands if c["bolt"].replace(".0", "") == w["bolt"].replace(".0", "")]
    if "offset" in w:
        cands = [c for c in cands if c["offset"].lstrip("+") == w["offset"]]
    if len(cands) > 1:
        fin = [c for c in cands if all(k in low for k in FINISH_WORDS.get(c["finish"].lower(), [c["finish"].lower()]))]
        if fin:
            cands = fin
        if len(cands) > 1:  # "gloss black" vs "gloss black milled": prefer the one without extra words
            plain = [c for c in cands if c["finish"].lower() == "gloss black"] if "milled" not in low else []
            cands = plain or cands
    if not cands:
        return None
    # same specs and same finish under two part numbers (e.g. 9110-2181P and 9110-2181PN): count for both
    if len({c["finish"].lower() for c in cands}) == 1:
        return [c["_sku"] for c in cands]
    return None


def per_wheel(price, title):
    t = (title or "").lower()
    if re.search(r"set of 4|4 wheels|\(4\)|4pc|4 pc|qty 4|x4\b", t):
        return round(price / 4, 2), 4
    if re.search(r"set of 2|pair|\(2\)", t):
        return round(price / 2, 2), 2
    return price, 1


def search(cat, headful):
    from playwright.sync_api import sync_playwright
    queries = build_queries(cat)
    hits = base.load_json("shopapp_hits.json", {})
    done = set(base.load_json("shopapp_done_queries.json", []))
    # second pass: part-number searches for anything still unmatched
    with sync_playwright() as p:
        browser = base.launch_browser(p, headful)
        page = browser.new_page(user_agent=base.UA, viewport={"width": 1400, "height": 1000})

        def run(qs, label):
            for i, q in enumerate(qs, 1):
                if q in done:
                    continue
                res = None
                for attempt in (1, 2):
                    try:
                        res = base.scrape_search(page, q, max_rounds=12)
                        break
                    except Exception as e:
                        print(f"  [{label} {i}/{len(qs)}] {q!r}: attempt {attempt} failed ({str(e).splitlines()[0]})")
                        page.wait_for_timeout(5000)
                if res is None:
                    continue
                new = 0
                for h in res:
                    store, title, price, compare = base.parse_card_text(h["texts"])
                    h.update(store=store, title=title, price=price, compare_at=compare, query=q)
                    h.pop("texts", None)
                    if not looks_like_cali(title or "", h["handle"]):
                        continue
                    if h["product_id"] not in hits:
                        new += 1
                    hits.setdefault(h["product_id"], h)
                done.add(q)
                base.save_json("shopapp_hits.json", hits)
                base.save_json("shopapp_done_queries.json", sorted(done))
                print(f"  [{label} {i}/{len(qs)}] {q!r}: {len(res)} results, {new} new (total {len(hits)})")
                base.nap(1.2, 2.5)

        run(queries, "model")
        matched = {m["sku"] for m in match(hits, cat)}
        sku_q = [f"Cali {c['sku']}" for c in cat if c["inv"] >= 4 and c["_sku"] not in matched]
        print(f"  {len(matched)} part numbers matched by model search; searching {len(sku_q)} part numbers directly")
        run(sku_q, "sku")
        browser.close()
    return hits


def match(hits, cat):
    by_sku = {c["_sku"]: c for c in cat}
    ship = base.load_known_shipping() if hasattr(base, "load_known_shipping") else {}
    out = []
    for h in hits.values():
        store = h.get("store") or "UNKNOWN STORE"
        if base.is_own(store) or not h.get("price"):
            continue
        title = h.get("title") or ""
        exact = sku_match(f"{title} {h['handle']}", by_sku)
        skus = [exact] if exact else (spec_match(title, h["handle"], cat) or [])
        if not skus:
            continue
        each, qty = per_wheel(h["price"], title)
        sh = None
        for k, v in ship.items():
            if k == store.lower() or k in store.lower() or store.lower() in k:
                sh = v
                break
        for s in skus:
          c = by_sku[s]
          out.append(dict(sku=c["sku"], model=c["model"], size=f"{c['diameter']}x{c['width']}", bolt=c["bolt"],
                        offset=c["offset"], finish=c["finish"], store=store, title=title, listed_price=h["price"],
                        listing_qty=qty, price_per_wheel=each, compare_at=h.get("compare_at"),
                        ship_per_wheel=sh, delivered_per_wheel=round(each + (sh or 0), 2),
                        how="part #" if exact else "specs", url=h["url"]))
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--headful", action="store_true")
    ap.add_argument("--skip-search", action="store_true")
    a = ap.parse_args()
    cat = load_catalog()
    hits = base.load_json("shopapp_hits.json", {}) if a.skip_search else search(cat, a.headful)
    rows = match(hits, cat)
    print(f"{len(hits)} Cali listings found, {len(rows)} matched to {len({r['sku'] for r in rows})} part numbers")
    if not rows:
        return
    rows.sort(key=lambda r: (r["sku"], r["delivered_per_wheel"]))
    with open("cali_prices.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    best = {}
    for r in rows:
        # skip obvious junk: a price under 40% of the median for that part number
        best.setdefault(r["sku"], []).append(r)
    out = []
    for sku, rs in best.items():
        ps = sorted(x["delivered_per_wheel"] for x in rs)
        med = ps[len(ps) // 2]
        ok = [x for x in rs if x["delivered_per_wheel"] >= 0.4 * med]
        b = min(ok, key=lambda x: x["delivered_per_wheel"])
        out.append(dict(sku=sku, lowest_delivered_per_wheel=b["delivered_per_wheel"], store=b["store"],
                        listings=len(rs), median=med, url=b["url"]))
    with open("cali_best.csv", "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(out[0].keys()))
        w.writeheader()
        w.writerows(out)
    print("saved cali_prices.csv and cali_best.csv")


if __name__ == "__main__":
    main()
