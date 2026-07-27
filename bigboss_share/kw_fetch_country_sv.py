"""
Reactivation-everywhere (READ-ONLY): KP GenerateKeywordHistoricalMetrics per non-universe
country for region-wide-paused keywords in proven themes. Returns SV + MoM% + surge flag.
-> country_keyword_sv_fetched. (Universe 6 countries get SV from adg_keyword_universe instead.)
"""
import os, sys, uuid, time
sys.stdout.reconfigure(encoding='utf-8')
import bigboss_config as cfg
import google.auth
from google.cloud import bigquery
from google.ads.googleads.client import GoogleAdsClient
import pandas as pd
from datetime import date
bq = cfg.bq_client()
ADS_CONFIG = cfg.GOOGLE_ADS
CID=cfg.GOOGLE_ADS_CUSTOMER_ID
GEO = {"Netherlands":2528,"Saudi Arabia":2682,"United Arab Emirates":2784,"South Africa":2710,
       "Brazil":2076,"Mexico":2484,"Israel":2376,"France":2250,"Italy":2380,"Belgium":2056,
       "Switzerland":2756,"New Zealand":2554,"Singapore":2702,"Spain":2724,"Poland":2616,
       "Turkey":2792,"Malaysia":2458,"Indonesia":2360,"Thailand":2764}
OUT=f"{cfg.BQ_PROJECT}.{cfg.BQ_DATASET}.country_keyword_sv_fetched"
BATCH=2000; GEN=date.today().isoformat()
ENUM={1:"JANUARY",2:"FEBRUARY",3:"MARCH",4:"APRIL",5:"MAY",6:"JUNE",7:"JULY",8:"AUGUST",9:"SEPTEMBER",10:"OCTOBER",11:"NOVEMBER",12:"DECEMBER"}
ads=GoogleAdsClient.load_from_dict(ADS_CONFIG); svc=ads.get_service("KeywordPlanIdeaService")

def paused_keywords(country):
    q="""SELECT DISTINCT k.keyword_text FROM `it-security-online-marketing.Google_ads_data_ajay.account_keyword_status` k
         JOIN `it-security-online-marketing.Google_ads_data_ajay.proven_themes` p
           ON k.product=p.product AND k.country=p.country AND k.theme=p.theme
         WHERE k.country=@c AND k.status IN ('PAUSED','REMOVED') AND k.match_type IN ('EXACT','PHRASE')
           AND k.keyword_text NOT IN (
             SELECT keyword_text FROM `it-security-online-marketing.Google_ads_data_ajay.account_keyword_status`
             WHERE country=@c AND status='ENABLED')"""
    job=bq.query(q,job_config=bigquery.QueryJobConfig(query_parameters=[bigquery.ScalarQueryParameter("c","STRING",country)]))
    return [r["keyword_text"] for r in job.result() if r["keyword_text"]]

def metrics(country, geo_id, kws):
    rows=[]
    geo=f"geoTargetConstants/{geo_id}"
    for i in range(0,len(kws),BATCH):
        batch=kws[i:i+BATCH]
        req=ads.get_type("GenerateKeywordHistoricalMetricsRequest")
        req.customer_id=CID; req.language="languageConstants/1000"
        req.geo_target_constants.append(geo)
        req.keyword_plan_network=ads.enums.KeywordPlanNetworkEnum.GOOGLE_SEARCH
        req.keywords.extend(batch)
        resp=None
        for a in range(3):
            try: resp=svc.generate_keyword_historical_metrics(request=req); break
            except Exception as e: print(f"    {country} batch {i} attempt {a+1}: {e}"); time.sleep(10*(a+1))
        if resp is None: continue
        for r in resp.results:
            m=r.keyword_metrics
            if not m: continue
            series=[]
            for msv in m.monthly_search_volumes:
                am=int(msv.month)-1
                if am<1 or am>12: continue
                try: d=date(int(msv.year),am,1)
                except: continue
                series.append((d, int(msv.monthly_searches) if msv.monthly_searches else 0))
            series.sort()
            sv=series[-1][1] if series else 0
            prev=series[-2][1] if len(series)>=2 else None
            mom=round((sv-prev)/prev*100,0) if prev else None
            recent=sum(v for _,v in series[-3:])/3 if len(series)>=3 else None
            prior=sum(v for _,v in series[-6:-3])/3 if len(series)>=6 else None
            surging = bool((mom is not None and mom>=20 and sv>=30) or (recent and prior and prior>0 and recent/prior>=1.3 and sv>=30))
            rows.append(dict(country=country, keyword_text=r.text.lower().strip(), sv=sv,
                mom_pct=mom, surging=surging,
                bid_high=(m.high_top_of_page_bid_micros/1e6) if m.high_top_of_page_bid_micros else None,
                generated_date=GEN))
        time.sleep(0.4)
    return rows

allrows=[]
for country,geo_id in GEO.items():
    kws=paused_keywords(country)
    if not kws: print(f"[skip] {country}: 0"); continue
    r=metrics(country,geo_id,kws)
    print(f"[ok] {country:<22} paused={len(kws):,} sv_rows={len(r):,} surging={sum(1 for x in r if x['surging'])}")
    allrows.extend(r)

df=pd.DataFrame(allrows)
print(f"\nTotal {len(df):,} sv rows")
bq.load_table_from_dataframe(df,OUT,job_config=bigquery.LoadJobConfig(write_disposition="WRITE_TRUNCATE"),
    job_id_prefix=f"ccsv_{uuid.uuid4().hex[:8]}_").result()
print(f"Wrote -> {OUT}")
