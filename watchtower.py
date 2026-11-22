#!/usr/bin/env python3
"""
Watchtower: a daily early-warning check for a book of SEO client accounts.

For every account workbook in ./accounts (same 6-tab format as the Fieldstone file),
it checks, as of a given date:

  1. INDEX LOSS     a site section's impressions collapse (Google stopped showing it)
  2. RANK SLIDE     a section keeps impressions but its position slides
  3. URL SWAP       a query's ranking URL switches to a different page (cannibalisation)
  4. NOINDEX MONEY  the latest crawl has noindex on pages that earned traffic
  5. LEAD DROP      organic demo requests fall vs the account's own baseline
  6. CAUSE HINTS    changes in the Change_Log in the 7 days before a signal started

Each finding gets a severity and is weighted by the leads the affected section used to
bring in, so the team works on what costs the client the most first.

Usage:
  python3 watchtower.py                       # check every account as of its latest data
  python3 watchtower.py --as-of 2026-09-12    # what would we have known on that day?
  python3 watchtower.py --backtest            # replay day by day; when did each alert first fire?
  python3 watchtower.py --slack-webhook URL   # also post the summary to Slack

Only free, standard pieces: Python 3 + openpyxl. Schedule with cron / GitHub Actions
(see README). The optional AI step (--ai) turns findings into a plain-English note for
the CSM; the detection itself is deterministic so it never "hallucinates" an alert.
"""
import argparse, collections, datetime as dt, glob, json, os, sys, urllib.request

try:
    import openpyxl
except ImportError:
    sys.exit("Install the one dependency first:  pip install openpyxl")

HERE = os.path.dirname(os.path.abspath(__file__))

# Thresholds (tuned so normal weekday/weekend swings never fire; see README)
BASELINE_DAYS = 28          # what "normal" looks like
RECENT_DAYS = 3             # compare the last 3 days against normal
IMPR_DROP = 0.35            # impressions down 35%+  -> index loss
POS_SLIDE = 2.0             # position worse by 2+ places with impressions holding -> rank slide
LEAD_DROP = 0.30            # leads down 30%+ week over baseline week
MIN_SECTION_IMPR = 500      # ignore tiny sections
SEV = {"CRITICAL": 3, "HIGH": 2, "MEDIUM": 1}


def d(x):
    return x.date() if isinstance(x, dt.datetime) else x


def load(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    rows = lambda name: [r for r in list(wb[name].iter_rows(values_only=True))[1:] if any(r)] if name in wb.sheetnames else []
    acc = {"name": os.path.splitext(os.path.basename(path))[0].replace("_", " ").title()}
    acc["daily"] = collections.defaultdict(dict)            # section -> date -> (clicks, impr, pos)
    for date, sec, c, i, p in rows("GSC_Daily"):
        acc["daily"][sec][d(date)] = (c or 0, i or 0, p or 0)
    acc["queries"] = collections.defaultdict(list)          # query -> [(week, url, clicks, impr, pos)]
    for wk, q, url, c, i, p in rows("GSC_Queries"):
        acc["queries"][q].append((d(wk), url, c or 0, i or 0, p or 0))
    acc["crawl"] = rows("Crawl_Export")
    acc["leads"] = collections.defaultdict(lambda: collections.Counter())  # date -> section -> n
    for date, sec, n in rows("Organic_Leads"):
        acc["leads"][d(date)][sec] += n or 0
    acc["changes"] = [(d(r[0]), r[1], r[2]) for r in rows("Change_Log")]
    acc["dates"] = sorted({x for s in acc["daily"].values() for x in s})
    return acc


def window(series, end, days):
    return [series[end - dt.timedelta(k)] for k in range(days) if end - dt.timedelta(k) in series]


def lead_share(acc, as_of):
    """Each section's share of leads over the baseline period (used to rank findings by money)."""
    tot = collections.Counter()
    start = as_of - dt.timedelta(BASELINE_DAYS + RECENT_DAYS + 7)
    for day, secs in acc["leads"].items():
        if start <= day <= as_of - dt.timedelta(RECENT_DAYS + 7):
            tot.update(secs)
    s = sum(tot.values()) or 1
    return {k: v / s for k, v in tot.items()}


def nearby_changes(acc, start, before=7, after=2):
    return [f"{c[0]} {c[1]}: {c[2]}" for c in acc["changes"]
            if start - dt.timedelta(before) <= c[0] <= start + dt.timedelta(after)
            and "report" not in str(c[2]).lower()]


def check_sections(acc, as_of, share):
    out = []
    base_end = as_of - dt.timedelta(RECENT_DAYS + 7)        # keep a gap so a slow slide still counts
    ref_end = as_of - dt.timedelta(30)                     # older "known good" reference for dating the start
    for sec, series in acc["daily"].items():
        recent = window(series, as_of, RECENT_DAYS)
        # compare like with like: only baseline days that fall on the same weekdays as the recent
        # window (B2B traffic halves at weekends; a Sat+Sun+Mon window must not look like a drop)
        wd = {(as_of - dt.timedelta(k)).weekday() for k in range(RECENT_DAYS)}
        base = [series[base_end - dt.timedelta(k)] for k in range(BASELINE_DAYS)
                if base_end - dt.timedelta(k) in series and (base_end - dt.timedelta(k)).weekday() in wd]
        if len(recent) < RECENT_DAYS or len(base) < 6:
            continue
        bi = sum(x[1] for x in base) / len(base)
        ri = sum(x[1] for x in recent) / len(recent)
        bp = sum(x[2] for x in base) / len(base)
        rp = sum(x[2] for x in recent) / len(recent)
        bc = sum(x[0] for x in base) / len(base)
        rc = sum(x[0] for x in recent) / len(recent)
        if bi < MIN_SECTION_IMPR:
            continue
        money = share.get(sec, 0)
        ref = [series[ref_end - dt.timedelta(k)] for k in range(BASELINE_DAYS)
               if ref_end - dt.timedelta(k) in series and (ref_end - dt.timedelta(k)).weekday() in wd] or base
        ri_ref = max(bi, sum(x[1] for x in ref) / len(ref))
        rp_ref = min(bp, sum(x[2] for x in ref) / len(ref))
        if ri < bi * (1 - IMPR_DROP) or ri < ri_ref * (1 - IMPR_DROP) and rc < bc:
            bi = ri_ref
            start = first_day(series, lambda x: x[1] < bi * 0.9, as_of)
            out.append(finding("INDEX LOSS", sec, money, start,
                f"{sec}: impressions {bi:,.0f}/day -> {ri:,.0f}/day ({ri/bi-1:+.0%}), clicks {bc:,.0f} -> {rc:,.0f}. "
                f"Google is showing these pages far less, not just lower. Check noindex/robots/canonicals/redirects.",
                nearby_changes(acc, start)))
        elif rp - rp_ref >= POS_SLIDE and ri > ri_ref * 0.6:
            bp = rp_ref
            start = first_day(series, lambda x: x[2] - bp >= 0.7, as_of)
            out.append(finding("RANK SLIDE", sec, money, start,
                f"{sec}: average position {bp:.1f} -> {rp:.1f} while impressions held ({ri/bi-1:+.0%}); "
                f"clicks {bc:,.0f} -> {rc:,.0f}/day. Look for another page taking these searches.",
                nearby_changes(acc, start)))
    return out


def first_day(series, bad, as_of):
    """Walk back from as_of to the first day of the current bad streak, judged on a
    3-day smoothed value so one good weekday does not hide the start."""
    def smooth(day):
        pts = [series[day - dt.timedelta(k)] for k in range(3) if day - dt.timedelta(k) in series]
        return tuple(sum(p[i] for p in pts) / len(pts) for i in range(3))
    day = as_of
    while day - dt.timedelta(1) in series and bad(smooth(day - dt.timedelta(1))) and day > min(series):
        day -= dt.timedelta(1)
    return day - dt.timedelta(2)  # smoothing lags by ~2 days


def check_url_swaps(acc, as_of, share):
    swaps = collections.defaultdict(list)
    for q, rows in acc["queries"].items():
        rows = sorted(r for r in rows if r[0] <= as_of)
        if len(rows) < 5:
            continue
        home = collections.Counter(r[1] for r in rows).most_common(1)[0][0]
        last = rows[-1]
        recent_urls = {r[1] for r in rows[-3:]}
        if last[1] == home and len(recent_urls) > 1:
            last = next(r for r in reversed(rows) if r[1] != home)
        if last[1] != home and "/blog/" in last[1] and "/blog/" not in home:
            before = sum(r[2] for r in rows[:-2]) / len(rows[:-2])
            first = next(r[0] for r in rows if r[1] != home)
            swaps[section_of(home)].append((q, home, last[1], before, last[2], first))
    out = []
    for sec, items in swaps.items():
        lines = [f'"{q}": {short(h)} -> {short(n)} (clicks/wk {b:.0f} -> {c})' for q, h, n, b, c, _ in items]
        start = min(i[5] for i in items)
        out.append(finding("URL SWAP", sec, share.get(sec, 0), start,
            f"{len(items)} money queries now rank a blog post instead of the {sec} page. " + "; ".join(lines[:6]),
            nearby_changes(acc, start, before=10)))
    return out


def check_crawl(acc, as_of, share):
    out, money_pages = [], collections.defaultdict(list)
    earning = {r[1] for rows in acc["queries"].values() for r in rows if r[2] >= 50}
    for url, status, robots, canon, title, words, inlinks, lastmod in acc["crawl"]:
        if not url or (lastmod and d(lastmod) > as_of):
            continue
        if robots and "noindex" in str(robots).lower() and (url in earning or "/solutions/" in url or "/compare/" in url):
            money_pages[section_of(url)].append(url)
    for sec, urls in money_pages.items():
        out.append(finding("NOINDEX MONEY", sec, share.get(sec, 0), None,
            f"{len(urls)} {sec} pages that earn traffic carry a noindex tag: " + ", ".join(short(u) for u in urls[:5]) + ("..." if len(urls) > 5 else ""),
            [], force="CRITICAL"))
    return out


def check_leads(acc, as_of):
    days = sorted(x for x in acc["leads"] if x <= as_of)
    if len(days) < 35:
        return []
    last7 = sum(sum(acc["leads"][x].values()) for x in days[-7:])
    lo = max(0, len(days) - 63) if len(days) >= 63 else len(days) - 35
    base = [sum(sum(acc["leads"][x].values()) for x in days[i:i + 7]) for i in range(lo, lo + 28, 7)]
    bw = sum(base) / len(base)
    if bw and last7 < bw * (1 - LEAD_DROP):
        bysec = collections.Counter()
        for x in days[-7:]:
            bysec.update(acc["leads"][x])
        basesec = collections.Counter()
        for x in days[lo:lo + 28]:
            basesec.update(acc["leads"][x])
        parts = [f"{s} {basesec[s]/4:.0f}->{bysec[s]}" for s in basesec if basesec[s] / 4 - bysec[s] >= 3]
        return [finding("LEAD DROP", "All", 1.0, days[-7],
            f"Organic demos last 7 days: {last7} vs {bw:.0f}/week normal ({last7/bw-1:+.0%}). Biggest losses: " + ", ".join(parts),
            [], force="CRITICAL" if last7 < bw * 0.5 else "HIGH")]
    return []


def finding(kind, sec, money, start, detail, causes, force=None):
    sev = force or ("CRITICAL" if money >= 0.25 else "HIGH" if money >= 0.10 else "MEDIUM")
    return {"type": kind, "section": sec, "severity": sev, "lead_share": round(money, 2),
            "started": str(start) if start else None, "detail": detail, "likely_causes": causes}


def section_of(url):
    for k in ("solutions", "compare", "blog", "pricing"):
        if f"/{k}" in url:
            return k.title()
    return "Homepage" if url.rstrip("/").count("/") <= 2 else "Other"


def short(url):
    return "/" + url.split("/", 3)[-1] if url.count("/") >= 3 else url


def run(acc, as_of):
    share = lead_share(acc, as_of)
    f = check_sections(acc, as_of, share) + check_url_swaps(acc, as_of, share) + check_crawl(acc, as_of, share) + check_leads(acc, as_of)
    return sorted(f, key=lambda x: (-SEV[x["severity"]], -x["lead_share"]))


def report(results, as_of_label):
    lines = [f"WATCHTOWER  {as_of_label}  |  {len(results)} accounts checked", ""]
    flagged = [(a, f) for a, f in results if f]
    if not flagged:
        lines.append("All accounts healthy.")
    for name, fs in sorted(flagged, key=lambda x: -max(SEV[f["severity"]] for f in x[1])):
        top = fs[0]["severity"]
        lines.append(f"[{top}] {name}: {len(fs)} issue(s)")
        for f in fs:
            started = f"  started {f['started']}" if f["started"] else ""
            lines.append(f"   - {f['severity']:<8} {f['type']:<13} {f['section']:<9} ({int(f['lead_share']*100)}% of leads){started}")
            lines.append(f"       {f['detail']}")
            for c in f["likely_causes"]:
                lines.append(f"       possible cause -> {c}")
        lines.append("")
    healthy = [a for a, f in results if not f]
    if healthy:
        lines.append(f"Healthy ({len(healthy)}): " + ", ".join(sorted(healthy)))
    return "\n".join(lines)


def ai_note(name, findings):
    """Optional: ask Claude to write the CSM a 5-line plain-English brief. Detection never depends on this."""
    key = os.environ.get("ANTHROPIC_API_KEY")
    if not key:
        return "(set ANTHROPIC_API_KEY to add an AI-written brief)"
    body = json.dumps({"model": "claude-sonnet-4-5", "max_tokens": 400, "messages": [{"role": "user", "content":
        "You brief an SEO account manager. Using ONLY these findings (do not invent numbers or causes), write 5 short lines: "
        "what broke, when, what it costs in leads, the first thing to check, and one sentence to tell the client.\n" + json.dumps(findings)}]}).encode()
    req = urllib.request.Request("https://api.anthropic.com/v1/messages", body, {"x-api-key": key, "anthropic-version": "2023-06-01", "content-type": "application/json"})
    return json.load(urllib.request.urlopen(req, timeout=60))["content"][0]["text"]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--accounts", default=os.path.join(HERE, "accounts"))
    ap.add_argument("--as-of")
    ap.add_argument("--backtest", action="store_true")
    ap.add_argument("--slack-webhook")
    ap.add_argument("--ai", action="store_true")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()

    accounts = [load(p) for p in sorted(glob.glob(os.path.join(a.accounts, "*.xlsx")))]
    if not accounts:
        sys.exit(f"No account workbooks in {a.accounts}")

    if a.backtest:
        for acc in accounts:
            first = {}
            for day in acc["dates"][35:]:
                for f in run(acc, day):
                    first.setdefault((f["type"], f["section"]), (day, f["severity"]))
            if first:
                print(f"\n{acc['name']}: first alert dates")
                for (t, s), (day, sev) in sorted(first.items(), key=lambda x: x[1][0]):
                    print(f"   {day}  {sev:<8} {t:<13} {s}")
        return

    results = []
    for acc in accounts:
        as_of = dt.date.fromisoformat(a.as_of) if a.as_of else acc["dates"][-1]
        results.append((acc["name"], run(acc, as_of)))
    label = f"as of {a.as_of}" if a.as_of else "latest data"
    if a.json:
        print(json.dumps({n: f for n, f in results}, indent=2))
        return
    text = report(results, label)
    if a.ai:
        text += "\n\nAI briefs:\n" + "\n".join(f"{n}:\n{ai_note(n, f)}" for n, f in results if f)
    print(text)
    if a.slack_webhook:
        urllib.request.urlopen(urllib.request.Request(a.slack_webhook, json.dumps({"text": "```" + text[:3500] + "```"}).encode(), {"content-type": "application/json"}))


if __name__ == "__main__":
    main()
