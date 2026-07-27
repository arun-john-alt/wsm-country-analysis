"""PAA engine: load harvested problem queries -> footprint-check (net-new) -> attach to
English proven (product,country,theme) cells -> paa_suggestions. build_recommendations.py
reads this and emits a 'PROBLEM QUERY [phrase-match]' line.

Monthly: Claude refreshes paa_queries.csv via WebSearch per proven theme, then runs this.
PAA queries are long-tail (KP reports ~0 individual SV) -> NOT SV-gated; they are a
phrase-match coverage layer, tier T2 (problem/how-to capability)."""
import os, sys, re, csv, uuid
sys.stdout.reconfigure(encoding='utf-8')
import bigboss_config as cfg
import google.auth
from google.cloud import bigquery
import pandas as pd
bq=cfg.bq_client()
P,D=cfg.BQ_PROJECT,cfg.BQ_DATASET
DESK=cfg.HOME
EN_COUNTRIES=('United States','United Kingdom','Canada','Australia','India')

_DIG={'2':'two','3':'three','4':'four','5':'five'}
_STOP={'for','the','a','an','in','of','to','with','and','your','on','is'}
def ckey(kw):
    s=re.sub(r'[^a-z0-9 ]',' ',str(kw).lower())
    return ' '.join(sorted(set(_DIG.get(t,t) for t in s.split() if t and t not in _STOP)))

# 1) load harvested queries
qs=[]
with open(os.path.join(DESK,"paa_queries.csv"),encoding="utf-8-sig",newline="") as f:
    for r in csv.DictReader(f):
        q=(r.get("query") or "").strip().lower()
        if q: qs.append((r["product"].strip(), r["theme"].strip(), q))
print(f"{len(qs)} harvested problem queries")

# 2) footprint (account-wide): drop ones already a keyword/served term anywhere
kw = bq.query(f"SELECT DISTINCT LOWER(TRIM(keyword_text)) t FROM `{P}.{D}.account_keyword_status` WHERE keyword_text IS NOT NULL").to_dataframe()['t'].tolist()
st = bq.query(f"SELECT DISTINCT LOWER(TRIM(keyword_text)) t FROM `{P}.{D}.adg_keyword_universe` WHERE keyword_text IS NOT NULL").to_dataframe()['t'].tolist()
foot=set(kw)|set(st); footc=set(ckey(x) for x in foot)
net=[(p,t,q) for (p,t,q) in qs if q not in foot and ckey(q) not in footc]
print(f"{len(net)} net-new (not a keyword/served term, cluster-new)")

# 3) proven English cells per (product,theme)
pt = bq.query(f"""SELECT DISTINCT product, country, theme FROM `{P}.{D}.proven_themes`
  WHERE country IN {EN_COUNTRIES}""").to_dataframe()
cells={}
for r in pt.itertuples(): cells.setdefault((r.product,r.theme),[]).append(r.country)

# 4) emit per (product,country,theme)
rows=[]
for (p,t,q) in net:
    for c in cells.get((p,t),[]):
        rows.append(dict(product=p, country=c, theme=t, query=q, tier='T2'))
out=pd.DataFrame(rows)
print(f"paa_suggestions rows={len(out)} across {out[['product','country','theme']].drop_duplicates().shape[0] if len(out) else 0} cells")
OUT=f"{P}.{D}.paa_suggestions"
bq.load_table_from_dataframe(out, OUT, job_config=bigquery.LoadJobConfig(write_disposition="WRITE_TRUNCATE"),
    job_id_prefix=f"paa_{uuid.uuid4().hex[:8]}_").result()
print(f"Wrote -> {OUT}")
