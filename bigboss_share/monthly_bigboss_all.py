"""
Big Boss — monthly all-in-one orchestrator.

Fires BOTH Big Boss columns into the same monthly tab, to co-run with the cloud
theme-report on the first Monday (e.g. July 6). Claude runs this manually on the 6th.

What it does (run in July -> targets the 'Jun 2026' tab):
  1. AI Keyword Recommendations  -> monthly_kw_recommendations.py (rebuild tables + assemble
     cells) then sheet_write_recs.py (writes RECS_TAB).
  2. Change History              -> monthly_change_history.py for the PRIOR month's changes
     (tab month MINUS 1; Jun 2026 tab gets May 2026 changes) at the given USD/INR rate.

Month mapping (confirmed): run-date month -> target tab = PRIOR month "%b %Y" (e.g. "Jun 2026").
Change-history month = tab month MINUS 1 (e.g. May 2026 -> --start 2026-05-01 --end 2026-05-31).

PREREQUISITES (NOT autonomous — confirm before running):
  - The target tab already exists AND you have manually inserted BOTH columns
    ('AI Keyword Recommendations' and 'Change History') in it. The sheet writers only UPDATE.
  - The month's Google Ads change-history CSV (with the User column) is sitting in Downloads.
  - --rate = that change-month's average USD/INR (fetch from x-rates.com/average).

Usage:
  python monthly_bigboss_all.py --rate 93.85                 # run-date = today, mode go
  python monthly_bigboss_all.py --rate 93.85 --run-date 2026-07-06
  python monthly_bigboss_all.py --rate 93.85 --tab "Jun 2026"   # explicit tab override
  python monthly_bigboss_all.py --rate 93.85 --skip-recs        # only Change History
  python monthly_bigboss_all.py --skip-changes                  # only AI col (rate not needed)
  python monthly_bigboss_all.py --rate 93.85 --plan             # print the plan, do nothing
Flags:
  --mode go|force|test  (default go) passed through to both sheet writers.
"""
import os, sys, runpy, argparse, calendar
from datetime import date, datetime

import bigboss_config as cfg
sys.stdout.reconfigure(encoding='utf-8')
DESK = cfg.HOME

ap = argparse.ArgumentParser()
ap.add_argument('--run-date', default=None, help='YYYY-MM-DD; default today')
ap.add_argument('--tab', default=None, help='Override target tab name; default = prior month "%%b %%Y"')
ap.add_argument('--rate', default=None, help='Avg USD/INR for the change-month (required unless --skip-changes)')
ap.add_argument('--mode', default='go', choices=['go', 'force', 'test'])
ap.add_argument('--skip-recs', action='store_true')
ap.add_argument('--skip-changes', action='store_true')
ap.add_argument('--plan', action='store_true', help='Print the resolved plan and exit')
a = ap.parse_args()

# --- resolve dates ---------------------------------------------------------
run_dt = datetime.strptime(a.run_date, '%Y-%m-%d').date() if a.run_date else date.today()

# target tab = prior calendar month of the run date
if run_dt.month == 1:
    tab_year, tab_month = run_dt.year - 1, 12
else:
    tab_year, tab_month = run_dt.year, run_dt.month - 1
target_tab = a.tab or date(tab_year, tab_month, 1).strftime('%b %Y')   # e.g. "Jun 2026"

# change-history month = tab month MINUS 1
if tab_month == 1:
    chg_year, chg_month = tab_year - 1, 12
else:
    chg_year, chg_month = tab_year, tab_month - 1
chg_start = date(chg_year, chg_month, 1).strftime('%Y-%m-%d')
chg_end = date(chg_year, chg_month, calendar.monthrange(chg_year, chg_month)[1]).strftime('%Y-%m-%d')

print("=" * 70)
print(f"Big Boss monthly orchestrator")
print(f"  run-date     : {run_dt}")
print(f"  target tab   : {target_tab}")
print(f"  mode         : {a.mode}")
print(f"  AI recs      : {'SKIP' if a.skip_recs else 'RUN'}")
print(f"  Change Hist  : {'SKIP' if a.skip_changes else f'RUN (changes {chg_start}..{chg_end}, rate {a.rate})'}")
print("=" * 70)

if not a.skip_changes and not a.rate:
    sys.exit("ERROR: --rate is required unless --skip-changes. Fetch avg USD/INR for the change-month from x-rates.com/average.")

if a.plan:
    print("[plan] nothing executed.")
    sys.exit(0)

# --- 1) AI Keyword Recommendations ----------------------------------------
if not a.skip_recs:
    print("\n########## [1/2] AI KEYWORD RECOMMENDATIONS ##########")
    print(">>> rebuilding tables + assembling cells (monthly_kw_recommendations.py) ...")
    runpy.run_path(os.path.join(DESK, "monthly_kw_recommendations.py"), run_name="__main__")
    print(f">>> writing 'AI Keyword Recommendations' into tab '{target_tab}' (sheet_write_recs.py {a.mode}) ...")
    os.environ['RECS_TAB'] = target_tab
    sys.argv = ['sheet_write_recs.py', a.mode]
    runpy.run_path(os.path.join(DESK, "sheet_write_recs.py"), run_name="__main__")
    print("[ok] AI Keyword Recommendations done")

# --- 2) Change History -----------------------------------------------------
if not a.skip_changes:
    print("\n########## [2/2] CHANGE HISTORY ##########")
    print(f"    (reads ALL change-history CSVs in Downloads; confirm the {chg_start[:7]} CSV is there)")
    sys.argv = ['monthly_change_history.py', '--start', chg_start, '--end', chg_end,
                '--rate', str(a.rate), '--tab', target_tab, '--mode', a.mode]
    runpy.run_path(os.path.join(DESK, "monthly_change_history.py"), run_name="__main__")
    print("[ok] Change History done")

print("\n=== monthly_bigboss_all complete ===")
