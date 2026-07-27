"""
Assemble per (Product, Country, Theme) 'AI Keyword Recommendations' cell text.
- Engine A (GAP) + Engine B (REACTIVATE) from cross_country_suggestions  (6 countries)
- Engine C (NEW) from keyword_expansion_suggestions                      (US/UK discovery)
Cross-country pointers are DEDUPED to the cluster head (one per cluster_keyword + "+N variants").
Reactivation shows last-known QS context. Writes kw_ai_recommendations.
"""
import os, sys, uuid
sys.stdout.reconfigure(encoding='utf-8')
import bigboss_config as cfg
import google.auth
from google.cloud import bigquery
import pandas as pd
bq = cfg.bq_client()
P, D = cfg.BQ_PROJECT, cfg.BQ_DATASET
OUT = f"{P}.{D}.kw_ai_recommendations"
CAP = 8
ABBR = {"United States":"US","United Kingdom":"UK","Canada":"CA","Australia":"AU","India":"IN","Germany":"DE"}
def abbr(s):
    s = str(s) if s is not None else ""
    for k,v in ABBR.items(): s = s.replace(k,v)
    return s
def I(x):
    try: return int(x)
    except: return None
import re
_DIG={'2':'two','3':'three','4':'four','5':'five'}
_STOP={'for','the','a','an','in','of','to','with','and','your','on','is'}
def ckey(kw):
    """normalized cluster key: collapses word-order, 2<->two, match punctuation, stopwords."""
    s=re.sub(r'[^a-z0-9 ]',' ',str(kw).lower())
    toks=[_DIG.get(t,t) for t in s.split() if t and t not in _STOP]
    return ' '.join(sorted(set(toks)))
def dedup_ck(df):
    """group rows by normalized cluster key; yield (rep_row, n_variants) by max sv."""
    d=df.copy(); d['__ck']=d['keyword'].map(ckey); out=[]
    for _,grp in d.groupby('__ck'):
        grp=grp.sort_values('sv',ascending=False)
        out.append((grp.iloc[0], len(grp)-1))
    return out

cc = bq.query("SELECT country,product,theme,keyword,engine,sv,cluster_keyword,source_countries,last_qs,mom_pct,surging,conv_countries,tier,priority FROM `%s.%s.cross_country_suggestions`" % (P,D)).to_dataframe()
disc = bq.query("SELECT country,product,theme,suggested_keyword keyword,tier,sv FROM `%s.%s.keyword_expansion_suggestions` WHERE tier!='CONQUEST'" % (P,D)).to_dataframe()
try:
    loc = bq.query("SELECT country,product,theme,keyword,tier,sv FROM `%s.%s.localized_suggestions`" % (P,D)).to_dataframe()
except Exception:
    loc = pd.DataFrame(columns=["country","product","theme","keyword","tier","sv"])
try:
    ws = bq.query("SELECT country,product,theme,keyword,tier,sv FROM `%s.%s.ws_country_sv`" % (P,D)).to_dataframe()
except Exception:
    ws = pd.DataFrame(columns=["country","product","theme","keyword","tier","sv"])
try:  # PAA engine: long-tail problem queries (phrase-match), no SV (KP reports 0)
    paa = bq.query("SELECT country,product,theme,query,tier FROM `%s.%s.paa_suggestions`" % (P,D)).to_dataframe()
except Exception:
    paa = pd.DataFrame(columns=["country","product","theme","query","tier"])
for df in (cc, disc, loc, ws, paa):
    for c in df.columns: df[c] = df[c].astype(object)
print(f"cc rows={len(cc)} disc rows={len(disc)} loc rows={len(loc)} ws rows={len(ws)} paa rows={len(paa)}")

cc_g = {k: g for k, g in cc.groupby(["product","country","theme"])}
disc_g = {k: g for k, g in disc.groupby(["product","country","theme"])}
loc_g = {k: g for k, g in loc.groupby(["product","country","theme"])}
ws_g = {k: g for k, g in ws.groupby(["product","country","theme"])}
paa_g = {k: g for k, g in paa.groupby(["product","country","theme"])} if len(paa) else {}
keys = set(cc_g) | set(disc_g) | set(loc_g) | set(ws_g) | set(paa_g)
TRANK = {"T1":0,"T2":2,"T3":3}

def cluster_reps(sub):
    """Yield (tier, representative_row, n_variants) per cluster, ordered by tier then SV."""
    out = []
    for clu, grp in sub.groupby("cluster_keyword"):
        grp = grp.copy()
        head = grp[grp.keyword == clu]
        rep = head.iloc[0] if len(head) else grp.sort_values(["sv"], ascending=False).iloc[0]
        out.append((rep, len(grp)-1))
    out.sort(key=lambda x: (TRANK.get(x[0].tier,3), -(I(x[0].sv) or 0)))
    return out

rows = []
for key in keys:
    prod, ctry, theme = key
    pts = []
    g = cc_g.get(key)
    if g is not None:
        for rep, nvar in cluster_reps(g[g.engine=="GAP"]):
            extra = f" +{nvar} variants" if nvar else ""
            surge = f", SV surging +{I(rep.mom_pct)}% MoM" if bool(rep.surging) and pd.notna(rep.mom_pct) and (I(rep.mom_pct) or 0)>0 else (", SV rising" if bool(rep.surging) else "")
            conv = f", search term converted in {abbr(rep.conv_countries)}" if pd.notna(rep.conv_countries) and str(rep.conv_countries).strip() else ""
            pts.append((TRANK.get(rep.tier,3)*10+0,
                f"ADD [{rep.tier}] {rep.keyword} ({I(rep.sv)}){extra} | active keyword in {abbr(rep.source_countries)}, missing in this country{conv}{surge}"))
        for rep, nvar in cluster_reps(g[g.engine=="REACTIVATE"]):
            extra = f" +{nvar} variants" if nvar else ""
            qs = f", was QS {I(rep.last_qs)}" if pd.notna(rep.last_qs) else ", no QS history"
            surge = f", SV surging +{I(rep.mom_pct)}% MoM" if bool(rep.surging) and pd.notna(rep.mom_pct) and (I(rep.mom_pct) or 0)>0 else (", SV rising" if bool(rep.surging) else "")
            conv = f", search term converted in {abbr(rep.conv_countries)}" if pd.notna(rep.conv_countries) and str(rep.conv_countries).strip() else ""
            pts.append((TRANK.get(rep.tier,3)*10+1,
                f"REVIEW PAUSE [{rep.tier}] {rep.keyword} ({I(rep.sv)}){extra} | paused region-wide{qs}{conv}{surge}"))
    d = disc_g.get(key)
    if d is not None:
        for r, nvar in dedup_ck(d):
            extra = f" +{nvar} variants" if nvar else ""
            pts.append((TRANK.get(r.tier,3)*10+2,
                f"NEW [{r.tier}] {r.keyword} ({I(r.sv) if pd.notna(r.sv) else '?'}){extra} | net-new discovery"))
    l = loc_g.get(key)
    if l is not None:
        for _, r in l.sort_values("sv", ascending=False).iterrows():
            pts.append((TRANK.get(r.tier,3)*10+1,
                f"ADD [{r.tier}] {r.keyword} ({I(r.sv) if pd.notna(r.sv) else '?'}) | localized - local demand, US/UK lead-gen theme"))
    # white-space (zero account footprint anywhere) — own guaranteed slots so it isn't squeezed out
    ws_pts=[]
    w = ws_g.get(key)
    if w is not None:
        for r, nvar in dedup_ck(w):
            extra = f" +{nvar} variants" if nvar else ""
            ws_pts.append((TRANK.get(r.tier,3),
                f"NEW (untapped) [{r.tier}] {r.keyword} ({I(r.sv)}){extra} | zero account footprint anywhere"))
        ws_pts.sort(key=lambda x: x[0])
    # PAA problem queries (long-tail, phrase-match) — own guaranteed slots, clearly separated
    paa_pts=[]
    pg = paa_g.get(key)
    if pg is not None:
        for _, r in pg.iterrows():
            paa_pts.append(f"PROBLEM QUERY [{r['tier']}] {r['query']} (phrase-match) | not yet targeted - long-tail")
    main = [p[1] for p in sorted(pts, key=lambda x: x[0])[:CAP]]
    wsel = [p[1] for p in sorted(ws_pts, key=lambda x: x[0])[:3]]
    psel = paa_pts[:3]
    allpts = main + wsel + psel
    if allpts:
        rows.append(dict(product=prod, country=ctry, theme=theme, n_points=len(allpts),
                         cell_text="\n".join("- "+p for p in allpts)))

out = pd.DataFrame(rows)
print(f"cells assembled={len(out)}")
if len(out)==0: sys.exit("empty")
bq.load_table_from_dataframe(out, OUT, job_config=bigquery.LoadJobConfig(write_disposition="WRITE_TRUNCATE"),
    job_id_prefix=f"airec_{uuid.uuid4().hex[:8]}_").result()
print(f"Wrote {len(out)} cells -> {OUT}")
print(out.groupby("country").size().to_string())
