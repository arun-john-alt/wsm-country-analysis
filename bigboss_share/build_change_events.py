"""
Build event-level change_events from Google Ads change-history CSV exports (Downloads).
Native User = changed_by. Categorizes into: kw_added/kw_removed (positive), neg_kw_added/removed,
cpc (old->new), adtext, lp(final URL). Maps Campaign+Ad group -> theme/country. -> change_events.
"""
import os, sys, glob, re, uuid
sys.stdout.reconfigure(encoding='utf-8')
import bigboss_config as cfg
import google.auth
from google.cloud import bigquery
import pandas as pd
bq = cfg.bq_client()

def detect_enc(f):
    with open(f,'rb') as fh: raw=fh.read(3)
    if raw[:2] in (b'\xff\xfe', b'\xfe\xff'): return 'utf-16'
    if raw[:3]==b'\xef\xbb\xbf': return 'utf-8-sig'
    return 'utf-8'
def load(f):
    enc=detect_enc(f)
    with open(f, encoding=enc) as fh: head=[next(fh) for _ in range(3)]
    sep='\t' if '\t' in head[2] else ','
    df=pd.read_csv(f, sep=sep, skiprows=2, engine='python', quotechar='"', encoding=enc, dtype=str)
    df.columns=[c.strip() for c in df.columns]
    return df[['Date & time','Account','User','Campaign','Ad group','Changes']]

frames=[load(f) for f in glob.glob(os.path.join(cfg.DOWNLOADS, "Change history report*.csv"))]
df=pd.concat(frames, ignore_index=True).fillna('')
df=df.drop_duplicates(['Date & time','Account','Campaign','Ad group','Changes'])
print(f"Deduped rows: {len(df):,}")
# datetime
df['dt']=pd.to_datetime(df['Date & time'].str.replace('Sept','Sep',regex=False), format='%d %b %Y, %H:%M:%S', errors='coerce')
print(f"datetime parsed: {df['dt'].notna().mean()*100:.1f}%")
df=df[df['dt'].notna()].copy()
df['date']=df['dt'].dt.date

# theme/country map (two-tier: exact (campaign,adgroup) then name-level fallback)
tmap=bq.query("""SELECT CampaignName, AdGroupName, ANY_VALUE(Product) product, ANY_VALUE(Theme) theme,
  ANY_VALUE(Sub_Theme) sub_theme, ANY_VALUE(CampaignCountry) country, COUNT(*) c
  FROM `it-security-online-marketing.Google_ads_data_ajay.themes_firstlast_semroi`
  WHERE Lead_Type='Mktg(SPL)Leads' AND Theme IS NOT NULL AND CampaignName IS NOT NULL AND AdGroupName IS NOT NULL
  GROUP BY 1,2""").to_dataframe()
tmap['camp_l']=tmap.CampaignName.str.strip().str.lower(); tmap['adg_l']=tmap.AdGroupName.str.strip().str.lower()
M1={(r.camp_l, r.adg_l):(r.product,r.theme,r.sub_theme,r.country) for r in tmap.itertuples()}
_a=tmap.sort_values('c',ascending=False).drop_duplicates('adg_l')   # adgroup-name -> top theme/product
M_adg={r.adg_l:(r.product,r.theme,r.sub_theme) for r in _a.itertuples()}
_c=tmap.sort_values('c',ascending=False).drop_duplicates('camp_l')  # campaign-name -> top country
M_camp={r.camp_l:r.country for r in _c.itertuples()}
def map_row(camp, adg):
    pm=M1.get((camp,adg))
    if pm: return pm
    if adg and adg in M_adg:
        p,t,s=M_adg[adg]; return (p,t,s,M_camp.get(camp))
    return (None,None,None,None)

def parse_changes(text):
    """Return list of event dicts from a Changes cell."""
    events=[]
    lines=text.split('\n')
    cat=None
    for ln in lines:
        if not ln.strip(): continue
        indented = ln[0] in (' ','\t')
        s=ln.strip()
        if not indented:  # header line
            low=s.lower()
            if 'max. cpc' in low and ('increased' in low or 'decreased' in low): cat='cpc'
            elif re.match(r'\d+ negative .*keyword.*added', low): cat='neg_kw_added'
            elif re.match(r'\d+ negative .*keyword.*removed', low): cat='neg_kw_removed'
            elif re.match(r'\d+ (exact|phrase|broad).*keyword.*added', low): cat='kw_added'
            elif re.match(r'\d+ (exact|phrase|broad).*keyword.*removed', low): cat='kw_removed'
            elif 'responsive search ad' in low or 'expanded text ad' in low or re.match(r'\d+ ad\b', low) or re.search(r'\bad (edited|added|created|removed|enabled|paused)', low):
                cat='adtext'
                n=int(re.match(r'(\d+)', s).group(1)) if re.match(r'\d+', s) else 1
                events.append(dict(category='adtext', count=n, keyword=None, old_cpc=None, new_cpc=None, new_url=None, detail=s[:300]))
            elif 'final url' in low or 'landing page' in low:
                cat='lp'
                n=int(re.match(r'(\d+)', s).group(1)) if re.match(r'\d+', s) else 1
                events.append(dict(category='lp', count=n, keyword=None, old_cpc=None, new_cpc=None, new_url=None, detail=s[:300]))
            else: cat='other'
        else:  # detail line under current header
            if cat=='cpc':
                m=re.search(r'\[(.+?)\]:\s*[^\d]*([\d.,]+)\s*to\s*[^\d]*([\d.,]+)', s)
                if m:
                    events.append(dict(category='cpc', count=1, keyword=m.group(1),
                        old_cpc=float(m.group(2).replace(',','')), new_cpc=float(m.group(3).replace(',','')),
                        new_url=None, detail=s[:300]))
            elif cat in ('kw_added','kw_removed'):
                m=re.search(r'\[(.+?)\]|"(.+?)"', s)
                kw=(m.group(1) or m.group(2)) if m else s.lstrip('- ')
                events.append(dict(category=cat, count=1, keyword=kw, old_cpc=None, new_cpc=None, new_url=None, detail=s[:200]))
            elif cat in ('neg_kw_added','neg_kw_removed'):
                events.append(dict(category=cat, count=1, keyword=None, old_cpc=None, new_cpc=None, new_url=None, detail=s[:200]))
    return events

rows=[]
unmatched=0
for r in df.itertuples():
    product,theme,sub_theme,country = map_row(str(r.Campaign).strip().lower(), str(r._5).strip().lower())  # _5='Ad group'
    if country is None: unmatched+=1
    for ev in parse_changes(str(r.Changes)):
        rows.append(dict(date=r.date, changed_by=r.User, account=r.Account, campaign=r.Campaign,
            ad_group=r._5, product=product, theme=theme, sub_theme=sub_theme, country=country, **ev))
ev=pd.DataFrame(rows)
print(f"\nEvents: {len(ev):,} | rows unmatched to theme/country: {unmatched:,} ({unmatched/len(df)*100:.0f}%)")
print("Category counts:\n"+ev['category'].value_counts().to_string())
print("\nSample CPC:"); print(ev[ev.category=='cpc'][['date','keyword','old_cpc','new_cpc','changed_by','country','theme']].head(4).to_string(index=False))
print("\nSample kw_added:"); print(ev[ev.category=='kw_added'][['date','keyword','changed_by','country','theme']].head(4).to_string(index=False))

ev['date']=pd.to_datetime(ev['date'])
bq.load_table_from_dataframe(ev, "it-security-online-marketing.Google_ads_data_ajay.change_events",
    job_config=bigquery.LoadJobConfig(write_disposition="WRITE_TRUNCATE"),
    job_id_prefix=f"chgev_{uuid.uuid4().hex[:8]}_").result()
print(f"\nWrote {len(ev):,} -> change_events")
