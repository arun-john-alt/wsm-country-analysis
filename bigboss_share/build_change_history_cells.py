"""Aggregate change_events for a month into per-(Product,Country,Theme) cell text for the Big Boss sheet."""
import os, sys, uuid
sys.stdout.reconfigure(encoding='utf-8')
import bigboss_config as cfg
import google.auth
from google.cloud import bigquery
import pandas as pd
bq = cfg.bq_client()
START = os.environ.get('CH_START','2026-04-01')   # month window of changes -> next month's tab
END   = os.environ.get('CH_END','2026-04-30')
USD_INR = float(os.environ.get('CH_RATE','93.40'))  # that month's avg USD/INR (x-rates); CPC bids are INR -> show USD
OUT = f"{cfg.BQ_PROJECT}.{cfg.BQ_DATASET}.kw_change_history"
KW_CAP, CPC_CAP = 5, 5

ev = bq.query(f"""SELECT country,product,theme,category,keyword,old_cpc,new_cpc,count,changed_by
  FROM `it-security-online-marketing.Google_ads_data_ajay.change_events`
  WHERE date BETWEEN '{START}' AND '{END}' AND country IS NOT NULL AND theme IS NOT NULL""").to_dataframe()
print(f"April matched events: {len(ev):,}")

rows=[]
for (prod,ctry,theme), g in ev.groupby(['product','country','theme']):
    parts=[]
    added = sorted(set(g[g.category=='kw_added']['keyword'].dropna()))
    if added:
        shown = added[:KW_CAP]; extra = f" (+{len(added)-KW_CAP} more)" if len(added)>KW_CAP else ""
        parts.append("Keywords added (%d):\n  %s%s" % (len(added), "\n  ".join(shown), extra))
    cpc = g[g.category=='cpc'].dropna(subset=['keyword'])
    if len(cpc):
        cl = [f"{r.keyword}: ${r.old_cpc/USD_INR:.0f} -> ${r.new_cpc/USD_INR:.0f}" for r in cpc.drop_duplicates('keyword').head(CPC_CAP).itertuples()]
        ex = f" (+{cpc['keyword'].nunique()-CPC_CAP} more)" if cpc['keyword'].nunique()>CPC_CAP else ""
        parts.append("CPC changes (%d):\n  %s%s" % (cpc['keyword'].nunique(), "\n  ".join(cl), ex))
    n_ad = int(g[g.category=='adtext']['count'].sum())
    if n_ad: parts.append(f"Ad text changes: {n_ad}")
    n_lp = int(g[g.category=='lp']['count'].sum())
    if n_lp: parts.append(f"LP changes: {n_lp}")
    n_rm = g[g.category=='kw_removed']['keyword'].nunique()
    if n_rm: parts.append(f"Keywords removed: {n_rm}")
    if not parts: continue
    who = ", ".join(sorted(set(g['changed_by'].dropna()))[:4])
    cell = "\n".join(parts) + (f"\n(by: {who})" if who else "")
    rows.append(dict(product=prod, country=ctry, theme=theme,
        n_kw_added=len(added), n_cpc=int(cpc['keyword'].nunique()) if len(cpc) else 0,
        n_adtext=n_ad, n_lp=n_lp, cell_text=cell))
out=pd.DataFrame(rows)
print(f"Cells: {len(out)}  | countries: {out['country'].nunique()}")
bq.load_table_from_dataframe(out, OUT, job_config=bigquery.LoadJobConfig(write_disposition="WRITE_TRUNCATE"),
    job_id_prefix=f"chghist_{uuid.uuid4().hex[:8]}_").result()
print(f"Wrote {len(out)} -> {OUT}\n")
for _,r in out[out.country=='United States'].sort_values('n_kw_added',ascending=False).head(2).iterrows():
    print(f"=== {r.country} / {r.product} / {r.theme} ===\n{r.cell_text}\n")
