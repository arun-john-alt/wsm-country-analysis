"""
One-command monthly Change-History updater.
Workflow: user drops the new Google Ads change-history CSV (with User column) into Downloads,
then run this. It rebuilds change_events from ALL CSVs, builds that month's cells, and populates
the target tab's 'Change History' column.

Usage:
  python monthly_change_history.py --start 2026-05-01 --end 2026-05-31 --rate 93.85 --tab "June 2026" [--mode go|force]
Notes:
  - --start/--end = the MONTH OF CHANGES to report (goes into the NEXT month's tab).
  - --rate = that month's average USD/INR (from x-rates.com/average). CPC bids (INR) shown as USD.
  - --tab = the sheet tab to write into (must already contain a 'Change History' column).
"""
import os, sys, runpy, argparse
import bigboss_config as cfg
sys.stdout.reconfigure(encoding='utf-8')
DESK=cfg.HOME
ap=argparse.ArgumentParser()
ap.add_argument('--start', required=True); ap.add_argument('--end', required=True)
ap.add_argument('--rate', required=True); ap.add_argument('--tab', required=True)
ap.add_argument('--mode', default='go', choices=['go','force','test'])
a=ap.parse_args()
os.environ['CH_START']=a.start; os.environ['CH_END']=a.end; os.environ['CH_RATE']=a.rate; os.environ['CH_TAB']=a.tab
print(f"=== Change History: {a.start}..{a.end} (rate {a.rate}) -> tab '{a.tab}' [{a.mode}] ===")

print("\n[1/3] Rebuilding change_events from Downloads CSVs ...")
runpy.run_path(os.path.join(DESK,"build_change_events.py"), run_name="__main__")
print("\n[2/3] Building month cells -> kw_change_history ...")
runpy.run_path(os.path.join(DESK,"build_change_history_cells.py"), run_name="__main__")
print("\n[3/3] Populating sheet column ...")
sys.argv=['sheet_write_changes.py', a.mode]
runpy.run_path(os.path.join(DESK,"sheet_write_changes.py"), run_name="__main__")
print("\n=== done ===")
