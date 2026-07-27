"""Run a .sql file and print results as CSV. Avoids bq-CLI BOM/backtick mangling.
Usage: python runq.py <file.sql>"""
import os, sys, csv
sys.stdout.reconfigure(encoding='utf-8')
import bigboss_config as cfg
import google.auth
from google.cloud import bigquery
bq=cfg.bq_client()
JC=bigquery.QueryJobConfig(default_dataset=cfg.DEFAULT_DATASET)
sql=open(sys.argv[1],encoding="utf-8-sig").read()
rows=list(bq.query(sql,job_config=JC).result())
if not rows:
    print("(0 rows)"); sys.exit(0)
w=csv.writer(sys.stdout)
w.writerow(rows[0].keys())
for r in rows: w.writerow(list(r.values()))
