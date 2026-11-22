#!/usr/bin/env python3
"""
Builds a demo book of 20 accounts from the Fieldstone file so Watchtower can be shown
working across a whole team's book, not one client.

- fieldstone.xlsx            the real assessment data (left untouched)
- 18 healthy accounts        Fieldstone's healthy August pattern, rescaled and replayed,
                             with normal day-to-day noise and no problems
- 1 account with a small slip  (to show a MEDIUM alert, not just "all or nothing")

All demo accounts are clearly synthetic and named demo-*.xlsx.
"""
import os, random, shutil, datetime as dt
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
ACC = os.path.join(HERE, "accounts")
SRC = os.path.join(ACC, "fieldstone.xlsx")
CUTOFF = dt.datetime(2026, 9, 7)          # Fieldstone was healthy up to here
NAMES = ["northwind-legal", "brightpath-dental", "keel-logistics", "ember-coffee", "atlas-roofing",
         "harbor-insurance", "pinecrest-clinics", "lumen-analytics", "orbit-payroll", "cedar-realty",
         "summit-cpa", "tidal-swim", "vantage-security", "maple-tutoring", "quarry-construction",
         "nova-fintech", "willow-spa", "granite-saas", "slate-hvac"]


def replay_healthy(ws, value_cols, rnd, scale, slip=None):
    rows = list(ws.iter_rows(min_row=2))
    # snapshot the ORIGINAL values first: scaling cells in place and then copying them
    # forward would scale the replayed weeks twice (that bug made healthy demos look broken)
    orig = [tuple(c.value for c in r) for r in rows]
    healthy = [orig[k] for k, r in enumerate(rows) if r[0].value and r[0].value <= CUTOFF]
    period = len(healthy)
    for k, r in enumerate(rows):
        if not r[0].value:
            continue
        src = [type("V", (), {"value": v}) for v in (healthy[k % period] if r[0].value > CUTOFF else orig[k])]
        for c in value_cols:
            v = src[c].value
            if isinstance(v, (int, float)):
                noise = rnd.uniform(0.9, 1.1)
                pos_col = {"GSC_Daily": 4, "GSC_Queries": 5}.get(ws.title)
                if c == pos_col:                                            # positions: no scaling
                    r[c].value = round(v * rnd.uniform(0.97, 1.03), 1)
                else:
                    r[c].value = max(0, round(v * scale * noise))
            elif r[0].value > CUTOFF and c in (1, 2):
                r[c].value = v
        if slip and ws.title == "GSC_Daily" and r[1].value == slip and r[0].value > dt.datetime(2026, 10, 1):
            r[4].value = round((r[4].value or 0) + 2.6, 1)                 # a ranking slide, impressions intact


def main():
    os.makedirs(ACC, exist_ok=True)
    for i, name in enumerate(NAMES):
        rnd = random.Random(i)
        wb = openpyxl.load_workbook(SRC)
        scale = rnd.uniform(0.3, 1.6)
        replay_healthy(wb["GSC_Daily"], (2, 3, 4), rnd, scale, slip="Blog" if name == "lumen-analytics" else None)
        replay_healthy(wb["GSC_Queries"], (1, 2, 3, 4, 5), rnd, scale)
        replay_healthy(wb["Organic_Leads"], (1, 2), rnd, scale)
        cr = wb["Crawl_Export"]
        for r in cr.iter_rows(min_row=2):
            if r[2].value and "noindex" in str(r[2].value):
                r[2].value = "index, follow"
        cl = wb["Change_Log"]
        cl.delete_rows(2, cl.max_row)
        an = wb["Account_Notes"]
        an.delete_rows(2, an.max_row)
        wb.save(os.path.join(ACC, f"demo-{name}.xlsx"))
    print(f"Wrote {len(NAMES)} demo accounts + fieldstone.xlsx = {len(NAMES)+1} accounts in {ACC}")


if __name__ == "__main__":
    main()
