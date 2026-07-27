"""
Germany localization (READ-ONLY) — KP GenerateKeywordIdeas with German language (1001)
+ Germany geo, seeded with hybrid German/English tier-1 phrases (tech nouns kept English,
actions translated). Returns real local phrasings + German SV. -> kw_exp_ideas_raw_de
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
CUSTOMER_ID=cfg.GOOGLE_ADS_CUSTOMER_ID
GEO_DE="geoTargetConstants/2276"; LANG_DE="languageConstants/1001"
OUT=f"{cfg.BQ_PROJECT}.{cfg.BQ_DATASET}.kw_exp_ideas_raw_de"
GEN=date.today().isoformat()

# (product, theme): hybrid German/English tier-1 seeds (tech nouns EN, actions DE)
SEEDS = {
 ("ADAP","Active Directory Audit"): ["active directory auditing tool","ad änderungen überwachen","active directory audit software","wer hat was geändert active directory","active directory überwachung tool","ad audit tool"],
 ("ADMP","AD Management"): ["active directory verwaltung tool","active directory management software","ad benutzerverwaltung tool","active directory verwaltungstool","ad verwaltung software"],
 ("ADAP","Account Lockout"): ["account lockout analyzer","active directory konto gesperrt ursache","kontosperrung active directory finden","ad account lockout tool","konto gesperrt active directory tool"],
 ("ADAP","File Server Audit"): ["fileserver auditing software","dateiserver überwachung tool","wer hat datei gelöscht fileserver","dateizugriff überwachen tool","file server audit tool"],
 ("ELA","Syslog"): ["syslog server software","syslog server windows","syslog tool","syslog management software"],
 ("ADAP","Logon/Logoff"): ["anmeldung überwachen active directory","active directory anmeldeüberwachung","anmeldeprotokoll active directory tool","user logon tracking tool"],
 ("ADMP","AD User Reports"): ["active directory benutzer report tool","inaktive benutzer active directory report","active directory reporting tool","ad benutzer report"],
 ("ADMP","AD GPO Management"): ["gpo verwaltung tool","group policy management tool","gpo management software"],
 ("ELA","Event Log"): ["windows event log management software","ereignisprotokoll überwachung tool","windows ereignisanzeige tool","event log management software"],
 ("ADMP","AD Tools"): ["active directory tools","active directory verwaltungstools","ad admin tools"],
 ("ADMP","AD User Management"): ["active directory benutzerverwaltung tool","benutzer anlegen active directory tool","active directory user management tool"],
 ("ADMP","Permissions Reports"): ["ntfs berechtigungen report tool","active directory berechtigungen report","dateiberechtigungen report tool"],
 ("RMP","Microsoft 365 Backup"): ["microsoft 365 backup","office 365 backup lösung","microsoft 365 sicherung","office 365 backup software","microsoft 365 daten sichern"],
 ("RMP","Exchange Backup"): ["exchange backup software","exchange postfach sichern","exchange mailbox backup","exchange online backup tool"],
 ("ADSSP","MFA"): ["active directory mfa","multi faktor authentifizierung active directory","mfa für windows anmeldung","zwei faktor authentifizierung windows login","mfa lösung active directory"],
 ("ADSSP","SSPR"): ["self service passwort zurücksetzen","passwort zurücksetzen self service active directory","passwort self service tool active directory","active directory passwort zurücksetzen tool"],
 ("ELA","Log"): ["siem lösung","log management software","siem tool","log management tool"],
}
ads = GoogleAdsClient.load_from_dict(ADS_CONFIG)
svc = ads.get_service("KeywordPlanIdeaService")

def expand(product, theme, seeds):
    req=ads.get_type("GenerateKeywordIdeasRequest")
    req.customer_id=CUSTOMER_ID; req.language=LANG_DE
    req.geo_target_constants.append(GEO_DE)
    req.keyword_plan_network=ads.enums.KeywordPlanNetworkEnum.GOOGLE_SEARCH
    req.include_adult_keywords=False
    req.keyword_seed.keywords.extend(seeds)
    rows=[]; sset=set(s.lower().strip() for s in seeds)
    for attempt in range(3):
        try:
            resp=svc.generate_keyword_ideas(request=req)
            for r in resp:
                m=r.keyword_idea_metrics
                rows.append(dict(product=product,country="Germany",theme=theme,
                    idea_keyword=r.text.lower().strip(),
                    avg_monthly_searches=int(m.avg_monthly_searches) if m.avg_monthly_searches else 0,
                    competition=m.competition.name if m.competition else "UNSPECIFIED",
                    bid_high=(m.high_top_of_page_bid_micros/1e6) if m.high_top_of_page_bid_micros else None,
                    is_seed=r.text.lower().strip() in sset, generated_date=GEN))
            return rows
        except Exception as e:
            print(f"   attempt {attempt+1} failed: {e}"); time.sleep(15*(attempt+1))
    return rows

allrows=[]
for (product,theme),seeds in SEEDS.items():
    ideas=expand(product,theme,seeds)
    print(f"[ok] {product:<6} {theme:<26} ideas={len(ideas):,}")
    allrows.extend(ideas); time.sleep(1)
df=pd.DataFrame(allrows)
print(f"\nTotal {len(df):,} ideas, unique {df['idea_keyword'].nunique():,}")
bq.load_table_from_dataframe(df,OUT,job_config=bigquery.LoadJobConfig(write_disposition="WRITE_TRUNCATE"),
    job_id_prefix=f"kwexpde_{uuid.uuid4().hex[:8]}_").result()
print(f"Wrote -> {OUT}")
