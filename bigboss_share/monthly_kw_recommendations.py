"""
Monthly orchestrator — rebuilds the keyword-recommendation tables, then assembles the
per-(Product,Country,Theme) AI cells. Run this BEFORE quick_sheet_write.py each month.

Order:
  1. account_keyword_status.sql   (current keyword status + QS, 6 countries)
  2. proven_themes.sql            (themes with leads in last 36mo)
  3. cross_country.sql            (Engine A gap + Engine B reactivation)
  4. build_recommendations.py     (assemble cell_text -> kw_ai_recommendations)

NOTE: Engine C (net-new KP discovery) is refreshed separately by kw_expand_v2.py +
build_kw_suggestions.py (re-run monthly for US/UK; extend to other countries as localized).
After this runs, quick_sheet_write.py picks up kw_ai_recommendations automatically.
"""
import os, sys, runpy
sys.stdout.reconfigure(encoding='utf-8')
import bigboss_config as cfg
import google.auth
from google.cloud import bigquery
bq=cfg.bq_client()
JC=bigquery.QueryJobConfig(default_dataset=cfg.DEFAULT_DATASET)
DESK=cfg.HOME

def run_sql(f):
    sql=open(os.path.join(DESK,f),encoding="utf-8-sig").read()
    bq.query(sql,job_config=JC).result(); print(f"[ok] ran {f}")

for f in ["account_keyword_status.sql","proven_themes.sql"]:
    run_sql(f)

# per-country SV (+MoM surge) for non-universe markets, then unify with universe SV
try:
    runpy.run_path(os.path.join(DESK,"kw_fetch_country_sv.py"), run_name="__main__")
    print("[ok] ran kw_fetch_country_sv.py")
except Exception as e:
    print(f"[warn] kw_fetch_country_sv.py failed ({e}) - continuing with last SV")
run_sql("country_keyword_sv.sql")
run_sql("keyword_intent.sql")
run_sql("cross_country.sql")

# localized discovery (non-English markets): Germany (curated) + FR/ES/MX/BR/IT/NL/PL/TR (dialect-seeded)
for f in ["kw_expand_de.py","build_localized_de.py","kw_expand_locale.py"]:
    try:
        runpy.run_path(os.path.join(DESK,f), run_name="__main__")
        print(f"[ok] ran {f}")
    except Exception as e:
        print(f"[warn] {f} failed ({e}) - continuing")
run_sql("localized_loc.sql")
run_sql("localized_suggestions.sql")

# PAA engine: footprint-check harvested problem queries -> paa_suggestions.
# NOTE: the HARVEST step is a guided monthly step — Claude refreshes paa_queries.csv via
# WebSearch (PAA + related searches per proven theme) BEFORE this orchestrator runs.
try:
    runpy.run_path(os.path.join(DESK,"build_paa_suggestions.py"), run_name="__main__")
    print("[ok] ran build_paa_suggestions.py")
except Exception as e:
    print(f"[warn] build_paa_suggestions.py failed ({e}) - continuing")

# assemble cells (runs build_recommendations.py top-level code)
runpy.run_path(os.path.join(DESK,"build_recommendations.py"), run_name="__main__")
print("=== monthly_kw_recommendations done ===")
