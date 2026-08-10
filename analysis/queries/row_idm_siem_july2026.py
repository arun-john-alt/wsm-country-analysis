"""
ROW July 2026 — Leads by Product Group (IDM vs SIEM)
=====================================================
IDM  = ADMP (AD Manager Plus) + ADSSP (AD Self Service Plus)
SIEM = ADAP (ADAudit Plus)    + ELA   (Endpoint Log Analyzer / Log360)

Filters: standard 4-filter salesleads_qt (CRM-verified Aug 2026)
Markets: ROW only (excludes US / India / UK / Canada / Australia)
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '../../wsm-monitor'))
from wsm_cfg import PROJ, bq_client

bq = bq_client()
SL = f"{PROJ}.sales_presales_leads_no_pi.salesleads_qt"

# ── Date filter for July 2026 ──────────────────────────────────────────────
SL_DATE = """CONCAT(
  SUBSTR(Created_Time,8,4), '-',
  LPAD(CAST(CASE SUBSTR(Created_Time,4,3)
    WHEN 'Jan' THEN 1  WHEN 'Feb' THEN 2  WHEN 'Mar' THEN 3
    WHEN 'Apr' THEN 4  WHEN 'May' THEN 5  WHEN 'Jun' THEN 6
    WHEN 'Jul' THEN 7  WHEN 'Aug' THEN 8  WHEN 'Sep' THEN 9
    WHEN 'Oct' THEN 10 WHEN 'Nov' THEN 11 WHEN 'Dec' THEN 12
  END AS STRING), 2, '0'))"""

ROW_EXCLUDE = """COMMON_COUNTRY_NAME NOT IN (
  'United States','India','United Kingdom','Canada','Australia'
)"""

# ── Q1: Overall IDM vs SIEM — July 2026 (ROW total) ──────────────────────────
print("\n" + "="*70)
print("  Q1 · IDM vs SIEM — July 2026 ROW Total (salesleads_qt)")
print("="*70)

sql_q1 = f"""
SELECT
  CASE
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADMP|AD MANAGER|ADSSP|AD SELF') THEN 'IDM'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADAP|AD AUDIT|\\bELA\\b|ENDPOINT|LOG360') THEN 'SIEM'
    ELSE 'OTHER'
  END AS product_group,
  CASE
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADMP|AD MANAGER') THEN 'ADMP'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADSSP|AD SELF') THEN 'ADSSP'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADAP|AD AUDIT') THEN 'ADAP'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'\\bELA\\b|ENDPOINT|LOG360') THEN 'ELA/LOG360'
    ELSE 'OTHER'
  END AS product_subgroup,
  COUNT(DISTINCT Email)                                             AS leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))           AS convs
FROM `{SL}`
WHERE {SL_DATE} = '2026-07'
  AND Junk = 'false'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND {ROW_EXCLUDE}
GROUP BY 1, 2
HAVING product_group != 'OTHER'
ORDER BY product_group, product_subgroup
"""

rows = list(bq.query(sql_q1).result())
print(f"  {'Group':<6} {'Product':<12} {'Leads':>6} {'Convs':>6}  {'Conv%':>6}")
print("  " + "-"*42)
grp_totals = {}
for r in rows:
    pg = r['product_group']
    psg = r['product_subgroup']
    leads = r['leads']
    convs = r['convs']
    pct = f"{convs/leads*100:.1f}%" if leads else "—"
    print(f"  {pg:<6} {psg:<12} {leads:>6,} {convs:>6,}  {pct:>6}")
    grp_totals[pg] = grp_totals.get(pg, {'leads':0,'convs':0})
    grp_totals[pg]['leads'] += leads
    grp_totals[pg]['convs'] += convs

print("  " + "-"*42)
for pg in ['IDM','SIEM']:
    if pg in grp_totals:
        t = grp_totals[pg]
        pct = f"{t['convs']/t['leads']*100:.1f}%" if t['leads'] else "—"
        print(f"  {pg:<6} {'TOTAL':<12} {t['leads']:>6,} {t['convs']:>6,}  {pct:>6}")

# ── Q2: IDM vs SIEM by DM Region — July 2026 ─────────────────────────────────
print("\n" + "="*70)
print("  Q2 · IDM vs SIEM by DM Region — July 2026 (salesleads_qt)")
print("="*70)

DM_CASE = """CASE
  WHEN COMMON_COUNTRY_NAME IN ('United States','India','United Kingdom','Canada','Australia') THEN NULL
  WHEN COMMON_COUNTRY_NAME = 'Germany'              THEN 'Germany'
  WHEN COMMON_COUNTRY_NAME = 'Netherlands'          THEN 'Netherlands'
  WHEN COMMON_COUNTRY_NAME = 'Switzerland'          THEN 'Switzerland'
  WHEN COMMON_COUNTRY_NAME = 'Belgium'              THEN 'Belgium'
  WHEN COMMON_COUNTRY_NAME = 'France'               THEN 'France'
  WHEN COMMON_COUNTRY_NAME = 'Italy'                THEN 'Italy'
  WHEN COMMON_COUNTRY_NAME = 'United Arab Emirates' THEN 'United Arab Emirates'
  WHEN COMMON_COUNTRY_NAME = 'Saudi Arabia'         THEN 'Saudi Arabia'
  WHEN COMMON_COUNTRY_NAME = 'Turkey'               THEN 'Turkey'
  WHEN COMMON_COUNTRY_NAME = 'Spain'                THEN 'Spain'
  WHEN COMMON_COUNTRY_NAME = 'Brazil'               THEN 'Brazil'
  WHEN COMMON_COUNTRY_NAME = 'Mexico'               THEN 'Mexico'
  WHEN COMMON_COUNTRY_NAME = 'South Africa'         THEN 'South Africa'
  WHEN COMMON_COUNTRY_NAME = 'Israel'               THEN 'Israel'
  WHEN COMMON_COUNTRY_NAME = 'Poland'               THEN 'Poland'
  WHEN COMMON_COUNTRY_NAME = 'Singapore'            THEN 'Singapore'
  WHEN COMMON_COUNTRY_NAME IN (
    'Sweden','Norway','Denmark','Finland','Austria','Czech Republic','Hungary',
    'Romania','Bulgaria','Greece','Portugal','Croatia','Slovakia','Slovenia',
    'Estonia','Latvia','Lithuania','Luxembourg','Malta','Cyprus','Albania',
    'Serbia','Bosnia and Herzegovina','North Macedonia','Moldova','Ukraine',
    'Belarus','Iceland','Ireland','Russia','Kosovo','Montenegro','Andorra','Liechtenstein')
    THEN 'Rest Of Europe'
  WHEN COMMON_COUNTRY_NAME IN (
    'Colombia','Peru','Argentina','Chile','Venezuela','Ecuador','Bolivia','Paraguay',
    'Uruguay','Costa Rica','Panama','Honduras','Guatemala','El Salvador','Nicaragua',
    'Puerto Rico','Dominican Republic','Cuba','Trinidad and Tobago','Jamaica',
    'Belize','Haiti','Guyana','Suriname','Barbados')
    THEN 'Rest Of LATAM'
  WHEN COMMON_COUNTRY_NAME IN (
    'Egypt','Nigeria','Kenya','Morocco','Ghana','Tanzania','Ethiopia','Uganda',
    'Cameroon','Senegal','Zimbabwe','Zambia','Mozambique','Ivory Coast','Angola',
    'Algeria','Tunisia','Libya','Sudan','Jordan','Lebanon','Kuwait','Qatar',
    'Bahrain','Oman','Iraq','Yemen','Pakistan','Rwanda','Botswana')
    THEN 'Rest Of MEA'
  WHEN COMMON_COUNTRY_NAME IN (
    'China','Japan','South Korea','Vietnam','Thailand','Indonesia','Philippines',
    'Malaysia','Taiwan','Hong Kong','New Zealand','Bangladesh','Sri Lanka',
    'Myanmar','Cambodia','Nepal','Mongolia','Macao','Brunei','Fiji',
    'Papua New Guinea','Laos','Timor-Leste')
    THEN 'Rest Of APAC'
  ELSE NULL
END"""

sql_q2 = f"""
SELECT
  dm_region,
  product_group,
  SUM(leads) AS leads,
  SUM(convs) AS convs
FROM (
  SELECT
    {DM_CASE} AS dm_region,
    CASE
      WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADMP|AD MANAGER|ADSSP|AD SELF') THEN 'IDM'
      WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADAP|AD AUDIT|\\bELA\\b|ENDPOINT|LOG360') THEN 'SIEM'
      ELSE NULL
    END AS product_group,
    COUNT(DISTINCT Email)                                           AS leads,
    COUNT(DISTINCT IF(Conversion='converted', Email, NULL))         AS convs
  FROM `{SL}`
  WHERE {SL_DATE} = '2026-07'
    AND Junk = 'false'
    AND PRODUCT_GROUP = 'AD_GROUP'
    AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
    AND isHaveToBeRemoved = 'Non Junk Email'
  GROUP BY 1, 2
)
WHERE dm_region IS NOT NULL AND product_group IS NOT NULL
GROUP BY 1, 2
ORDER BY dm_region, product_group
"""

DM_ORDER = [
    "Germany","Netherlands","Switzerland","Belgium",
    "France","Italy","United Arab Emirates","Saudi Arabia","Turkey",
    "Spain","Brazil","Mexico","Rest Of LATAM","South Africa","Israel",
    "Rest Of Europe","Poland","Rest Of MEA","Rest Of APAC","Singapore"
]

rows2 = list(bq.query(sql_q2).result())
data = {}
for r in rows2:
    dm = r['dm_region']
    pg = r['product_group']
    if dm not in data:
        data[dm] = {}
    data[dm][pg] = {'leads': r['leads'], 'convs': r['convs']}

print(f"  {'DM Region':<24} {'IDM Leads':>10} {'IDM Convs':>10} {'SIEM Leads':>11} {'SIEM Convs':>11}")
print("  " + "-"*70)
idm_total = siem_total = idm_c_total = siem_c_total = 0
for dm in DM_ORDER:
    if dm in data:
        idm  = data[dm].get('IDM',  {'leads':0,'convs':0})
        siem = data[dm].get('SIEM', {'leads':0,'convs':0})
        il = idm['leads']; ic = idm['convs']
        sl = siem['leads']; sc = siem['convs']
        if il or sl:
            print(f"  {dm:<24} {il:>10,} {ic:>10,} {sl:>11,} {sc:>11,}")
        idm_total += il; idm_c_total += ic
        siem_total += sl; siem_c_total += sc

print("  " + "-"*70)
print(f"  {'ROW TOTAL':<24} {idm_total:>10,} {idm_c_total:>10,} {siem_total:>11,} {siem_c_total:>11,}")

print("\n✅ Done. IDM = ADMP+ADSSP | SIEM = ADAP+ELA/LOG360")
