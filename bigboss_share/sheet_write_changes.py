"""
Write kw_change_history into the 'Change History' column of the 'May 2026' tab,
matching rows by (Product, Country, Theme). Updates ONLY that column.
Usage:  python sheet_write_changes.py test | go | force
(Requires the 'Change History' column to already exist in the tab.)
"""
import os, sys, json, requests, time
sys.stdout.reconfigure(encoding='utf-8')
import bigboss_config as cfg
import google.auth
from google.cloud import bigquery
MODE = sys.argv[1] if len(sys.argv)>1 else "test"

SHEET_ID=cfg.ZOHO_SHEET_ID
CID=cfg.ZOHO_CLIENT_ID; CS=cfg.ZOHO_CLIENT_SECRET
RT=cfg.ZOHO_REFRESH_TOKEN
TAB=os.environ.get('CH_TAB','May 2026'); COL='Change History'

bq=cfg.bq_client()
recs={(r['product'],r['country'],r['theme']):r['cell_text'] for r in bq.query("SELECT product,country,theme,cell_text FROM `it-security-online-marketing.Google_ads_data_ajay.kw_change_history`").result()}
print(f"{len(recs)} change-history cells loaded")

def refresh_token():
    global h
    tok=requests.post('https://accounts.zoho.in/oauth/v2/token',data={'grant_type':'refresh_token','client_id':CID,'client_secret':CS,'refresh_token':RT}).json()['access_token']
    h={'Authorization':f'Zoho-oauthtoken {tok}'}
refresh_token()
fr=requests.post(f'https://sheet.zoho.in/api/v2/{SHEET_ID}',headers=h,data={'method':'worksheet.records.fetch','worksheet_name':TAB,'header_row':'1'}).json()
sheet_rows=fr.get('records',[])
print(f"{len(sheet_rows)} sheet rows fetched")
if sheet_rows and COL not in sheet_rows[0]:
    print(f"!! Column '{COL}' not found in tab '{TAB}'. Insert it first. Headers: {list(sheet_rows[0].keys())}")
    sys.exit(1)

def esc(v): return str(v).replace('"','\\"')
matches=[]; skipped=0; stale=0
for sr in sheet_rows:
    key=(sr.get('Product'),sr.get('Country'),sr.get('Theme'))
    filled=str(sr.get(COL) or '').strip()
    if key in recs:
        if MODE!="force" and filled: skipped+=1; continue
        matches.append((key, recs[key]))
    elif MODE=="force" and filled:
        matches.append((key, "")); stale+=1
print(f"{len(matches)} rows to update ({skipped} skipped, {stale} stale-clears)")

def update_row(key, text):
    p,c,t=key
    criteria=f'"Product"="{esc(p)}" and "Country"="{esc(c)}" and "Theme"="{esc(t)}"'
    payload={'method':'worksheet.records.update','worksheet_name':TAB,'criteria':criteria,'data':json.dumps({COL:text})}
    for net in range(5):
        try:
            r=requests.post(f'https://sheet.zoho.in/api/v2/{SHEET_ID}',headers=h,data=payload,timeout=60)
            return r.json()
        except requests.exceptions.RequestException as e:
            print(f"  net retry {net+1} ({type(e).__name__})"); time.sleep(10*(net+1))
    return {'status':'neterror'}

if MODE=="test":
    key,text=matches[0]; print("TEST updating",key)
    print(json.dumps(update_row(key,text),ensure_ascii=False)[:600])
elif MODE in ("go","force"):
    ok=0
    for i,(key,text) in enumerate(matches):
        if i>0 and i%150==0: refresh_token()   # proactively refresh before token expiry
        for attempt in range(4):
            resp=update_row(key,text)
            if resp.get('status')=='success': ok+=1; break
            if resp.get('error_code')==2401: refresh_token(); continue   # token expired -> refresh + retry
            if resp.get('error_code')==2950: time.sleep(20*(attempt+1)); continue
            print("FAIL",key,json.dumps(resp,ensure_ascii=False)[:200]); break
        time.sleep(0.8)
        if (i+1)%25==0: print(f"  {i+1}/{len(matches)} done, ok={ok}")
    print(f"Updated {ok}/{len(matches)} rows")
