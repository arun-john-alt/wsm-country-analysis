"""Curated German localized suggestions (whole-theme gaps: proven in US/UK, German demand, DE underinvested).
Writes localized_suggestions (product,country,theme,keyword,tier,sv). Conquest kept separate (not written)."""
import os, sys, uuid
sys.stdout.reconfigure(encoding='utf-8')
import bigboss_config as cfg
import google.auth
from google.cloud import bigquery
import pandas as pd
bq = cfg.bq_client()
P, D = cfg.BQ_PROJECT, cfg.BQ_DATASET
OUT = f"{P}.{D}.localized_de"

# (product, country, theme): {T1:[...],T2:[...],T3:[...]}  curated German, intent-screened, competitor excluded
CUR = {
 ("RMP","Germany","Microsoft 365 Backup"): dict(
   T1=["backup microsoft 365","backup office 365","backup o365"],
   T2=["ms 365 backup","office 365 backup cloud","backup m365","datensicherung office 365","office 365 datensicherung"],
   T3=["sharepoint datensicherung","outlook 365 backup erstellen"]),
 ("ELA","Germany","Log"): dict(
   T1=["siem soc","siem system"], T2=["siem it security"], T3=["siem software open source","siem open source"]),
}
rows=[]
for (prod,ctry,theme),tiers in CUR.items():
    for tier,kws in tiers.items():
        for kw in kws:
            rows.append(dict(product=prod,country=ctry,theme=theme,keyword=kw,tier=tier))
picks=pd.DataFrame(rows)
# attach SV from the German ideas table
sv=bq.query("SELECT LOWER(TRIM(idea_keyword)) kw, MAX(avg_monthly_searches) sv FROM `%s.%s.kw_exp_ideas_raw_de` GROUP BY 1"%(P,D)).to_dataframe()
svmap=dict(zip(sv.kw, sv.sv))
picks["sv"]=picks.keyword.str.lower().str.strip().map(svmap)
bq.load_table_from_dataframe(picks, OUT, job_config=bigquery.LoadJobConfig(write_disposition="WRITE_TRUNCATE"),
    job_id_prefix=f"loc_{uuid.uuid4().hex[:8]}_").result()
print(f"Wrote {len(picks)} localized suggestions -> {OUT}")
print(picks.to_string(index=False))
