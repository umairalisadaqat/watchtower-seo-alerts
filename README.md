# Watchtower: early-warning SEO checks for a book of client accounts

Runs every morning across every client workbook and posts one ranked list to Slack: who is on fire, what broke, when it started, which change probably caused it, and how many leads it touches.

## Run it
```bash
python3 server.py                  # open http://localhost:4918 : the morning report in the browser, with a date slider to replay any morning
pip install openpyxl
python3 make_demo_book.py          # builds 19 demo accounts next to fieldstone.xlsx (20 total)
python3 watchtower.py              # check all accounts on their latest data
python3 watchtower.py --backtest   # replay day by day: when would each alert first have fired?
python3 watchtower.py --as-of 2026-09-15   # what we would have seen on that morning
python3 watchtower.py --json       # machine-readable, for a dashboard
python3 watchtower.py --slack-webhook https://hooks.slack.com/...   # post to the team channel
ANTHROPIC_API_KEY=... python3 watchtower.py --ai   # optional plain-English brief per flagged account
```

## Schedule it (runs on its own)
cron, every day at 7am:
```
0 7 * * * cd /path/to/watchtower && python3 watchtower.py --slack-webhook "$SLACK_URL"
```
In production the `accounts/` workbooks are refreshed nightly from the GSC API, the crawler and the CRM. The checks don't change.

## Checks
| Check | Rule (defaults in the top of watchtower.py) |
|---|---|
| NOINDEX MONEY | crawl shows noindex on a page that earns traffic or is a Solutions/Compare page |
| INDEX LOSS | section impressions (last 3 days) down 35%+ vs same weekdays in the prior 4 weeks |
| RANK SLIDE | position worse by 2+ places while impressions hold |
| URL SWAP | a money query's ranking URL switches to a blog post |
| LEAD DROP | organic demos last 7 days down 30%+ vs the account's normal week |
| CAUSE HINTS | Change_Log entries in the 7 days before a signal started |

Findings are ranked by severity, then by the share of leads the affected section brings in, so a drop on a page that drives half the demos outranks a bigger traffic drop on a blog that drives none.

## Design choices
- **Detection is plain rules, not AI.** Alerts must be explainable and can't be made up. AI only writes the summary, and only from the findings.
- **Weekday-matched baselines.** B2B traffic halves at weekends, and naive baselines fire every Monday.
- **The "known good" reference sits 30 days back**, so a slow slide can't become the new normal.
- **Lead-weighted.** CSMs work the alerts that cost the client money first.
