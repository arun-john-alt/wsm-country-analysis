"""
Saudi Arabia SEM Deep-Dive — Part 2 Queries
Run after saudi_arabia_sem_queries.py (all constants same).
Covers:
  P1b: ADAP monthly CPC / CTR from ads_AdGroupBasicStats
  P2a: Google NB + Bing NB monthly lead/conv trend
  P2c: Google NB monthly by product (ADAP/ADMP/ELA)
  P3:  Top SEO URLs for Saudi Arabia with leads + convs
"""
import sys
sys.path.insert(0, '/Users/arun-8846/Downloads/Monitor/wsm-monitor')
from wsm_cfg import G, PROJ, bq_client, ACCT

bq = bq_client()
ROI = f"{PROJ}.{G.split('.')[-1]}.themes_firstlast_semroi"
SL  = f"{PROJ}.sales_presales_leads_no_pi.salesleads_qt"
ADS = f"{PROJ}.{G.split('.')[-1]}.ads_AdGroupBasicStats_{ACCT}"
AG  = f"{PROJ}.{G.split('.')[-1]}.ads_AdGroup_{ACCT}"
CAM = f"{PROJ}.{G.split('.')[-1]}.ads_Campaign_{ACCT}"

COUNTRY_SL  = "saudi arabia"

def run(label, sql):
    print(f"\n{'═'*70}")
    print(f"  {label}")
    print('═'*70)
    rows = list(bq.query(sql).result())
    if not rows:
        print("  (no rows)")
        return rows
    keys = list(rows[0].keys())
    widths = [max(len(str(k)), max(len(str(r[k])) for r in rows)) for k in keys]
    fmt = "  " + "  ".join(f"{{:<{w}}}" for w in widths)
    print(fmt.format(*keys))
    print("  " + "  ".join("─"*w for w in widths))
    for r in rows:
        print(fmt.format(*[str(r[k]) for k in keys]))
    return rows

# ── P1: ADAP monthly CPC / CTR from ads_AdGroupBasicStats ────────────────────
# Joins AdGroupBasicStats → AdGroup + Campaign → themes tmap (Product=ADAP, Saudi Arabia)
# Returns INR (raw Google Ads billing)
run("P1 · ADAP Saudi Arabia — monthly clicks, spend(INR), avg CPC(INR), CTR (ads table)", f"""
WITH tmap AS (
  SELECT
    LOWER(TRIM(CampaignName))  AS cn,
    LOWER(TRIM(AdGroupName))   AS agn
  FROM `{ROI}`
  WHERE CampaignCountry = 'Saudi Arabia'
    AND Product = 'ADAP'
    AND CampaignName IS NOT NULL
    AND AdGroupName IS NOT NULL
    AND COALESCE(Source_Medium, Source___Medium) = 'google / cpc'
  GROUP BY 1, 2
),
ag AS (
  SELECT ad_group_id, LOWER(TRIM(ad_group_name)) agn
  FROM `{AG}`
  GROUP BY 1, 2
),
cam AS (
  SELECT campaign_id, LOWER(TRIM(campaign_name)) cn
  FROM `{CAM}`
  GROUP BY 1, 2
)
SELECT
  FORMAT_DATE('%Y-%m', segments_date)  AS ym,
  SUM(metrics_clicks)                  AS clicks,
  SUM(metrics_impressions)             AS impressions,
  ROUND(SUM(metrics_cost_micros)/1e6, 0) AS spend_inr,
  ROUND(SAFE_DIVIDE(SUM(metrics_cost_micros)/1e6,
                    SUM(metrics_clicks)), 0)  AS avg_cpc_inr,
  ROUND(SAFE_DIVIDE(SUM(metrics_clicks),
                    SUM(metrics_impressions)) * 100, 2) AS ctr_pct
FROM `{ADS}` a
JOIN ag   ON ag.ad_group_id = a.ad_group_id
JOIN cam  ON cam.campaign_id = a.campaign_id
JOIN tmap ON tmap.cn = ag.agn AND tmap.agn = ag.agn
             -- correct join: match campaign name AND ad group name
WHERE a.segments_date >= '2024-01-01'
  AND a.segments_ad_network_type IN ('SEARCH', 'SEARCH_PARTNERS')
GROUP BY 1
ORDER BY 1
""")

# ── P1b: ADAP join fix — use both campaign + adgroup name ────────────────────
# The tmap join above has a bug (joins agn=agn instead of cn=cn). Correct version:
run("P1b · ADAP Saudi Arabia — monthly (corrected join)", f"""
WITH tmap AS (
  SELECT
    LOWER(TRIM(CampaignName))  AS cam_name,
    LOWER(TRIM(AdGroupName))   AS ag_name
  FROM `{ROI}`
  WHERE CampaignCountry = 'Saudi Arabia'
    AND Product = 'ADAP'
    AND CampaignName IS NOT NULL
    AND AdGroupName IS NOT NULL
    AND COALESCE(Source_Medium, Source___Medium) = 'google / cpc'
  GROUP BY 1, 2
),
perf AS (
  SELECT
    FORMAT_DATE('%Y-%m', a.segments_date) ym,
    SUM(a.metrics_clicks)               clicks,
    SUM(a.metrics_impressions)          impr,
    SUM(a.metrics_cost_micros)/1e6      cost_inr,
    a.ad_group_id,
    a.campaign_id
  FROM `{ADS}` a
  WHERE a.segments_date >= '2024-01-01'
    AND a.segments_ad_network_type IN ('SEARCH', 'SEARCH_PARTNERS')
  GROUP BY ym, a.ad_group_id, a.campaign_id
)
SELECT
  p.ym,
  SUM(p.clicks)                              AS clicks,
  ROUND(SUM(p.cost_inr), 0)                 AS spend_inr,
  ROUND(SAFE_DIVIDE(SUM(p.cost_inr), SUM(p.clicks)), 0)   AS avg_cpc_inr,
  ROUND(SAFE_DIVIDE(SUM(p.clicks), SUM(p.impr)) * 100, 2) AS ctr_pct
FROM perf p
JOIN `{AG}`  ag  ON ag.ad_group_id  = p.ad_group_id
JOIN `{CAM}` cam ON cam.campaign_id = p.campaign_id
JOIN tmap ON tmap.cam_name = LOWER(TRIM(cam.campaign_name))
         AND tmap.ag_name  = LOWER(TRIM(ag.ad_group_name))
GROUP BY 1
ORDER BY 1
""")

# ── P2a: Google NB — monthly leads + convs (salesleads_qt, Saudi Arabia) ─────────────
# Use to identify WHEN Google NB conv rate collapsed
run("P2a · Google NB — monthly leads + convs (salesleads_qt)", f"""
SELECT
  FORMAT_DATE('%Y-%m',
    SAFE.PARSE_DATE('%d %b %Y', SUBSTR(Created_Time,1,11))) AS ym,
  COUNT(DISTINCT ID)                                         AS leads,
  COUNT(DISTINCT IF(Conversion='converted', ID, NULL))       AS convs
FROM `{SL}`
WHERE LOWER(COMMON_COUNTRY_NAME) = '{COUNTRY_SL}'
  AND FIRST_SRC_GRP = 'google / cpc'
  AND FIRST_SRC_THEME NOT IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
  AND (FIRST_SRC_CAMPAIGN_TYPE != 'Display' OR FIRST_SRC_CAMPAIGN_TYPE IS NULL)
  AND Junk = 'false'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND SAFE.PARSE_DATE('%d %b %Y', SUBSTR(Created_Time,1,11))
      BETWEEN '2024-01-01' AND '2026-06-30'
GROUP BY 1
ORDER BY 1
""")

# ── P2b: Bing NB — monthly leads + convs (salesleads_qt, Saudi Arabia) ───────────────
# Use to identify WHEN Bing NB conv rate improved (0%→10.3%)
run("P2b · Bing NB — monthly leads + convs (salesleads_qt)", f"""
SELECT
  FORMAT_DATE('%Y-%m',
    SAFE.PARSE_DATE('%d %b %Y', SUBSTR(Created_Time,1,11))) AS ym,
  COUNT(DISTINCT ID)                                         AS leads,
  COUNT(DISTINCT IF(Conversion='converted', ID, NULL))       AS convs
FROM `{SL}`
WHERE LOWER(COMMON_COUNTRY_NAME) = '{COUNTRY_SL}'
  AND FIRST_SRC_GRP = 'bing / cpc'
  AND FIRST_SRC_THEME NOT IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
  AND (FIRST_SRC_CAMPAIGN_TYPE != 'Display' OR FIRST_SRC_CAMPAIGN_TYPE IS NULL)
  AND Junk = 'false'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND SAFE.PARSE_DATE('%d %b %Y', SUBSTR(Created_Time,1,11))
      BETWEEN '2024-01-01' AND '2026-06-30'
GROUP BY 1
ORDER BY 1
""")

# ── P2c: Google NB monthly by product (themes) — leads + spend ────────────────
# ADAP/ADMP/ELA group month by month to see when ADAP collapsed
run("P2c · Google NB monthly by product — leads + spend (themes)", f"""
SELECT
  SUBSTR(Date,1,7)      AS ym,
  Product,
  ROUND(SUM(Valid_Sales_Leads_First_Source)) AS leads,
  ROUND(SUM(Cost), 0)                        AS spend_usd
FROM `{ROI}`
WHERE CampaignCountry = 'Saudi Arabia'
  AND COALESCE(Source_Medium, Source___Medium) = 'google / cpc'
  AND Theme NOT IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
  AND Lead_Type = 'All Leads'
  AND SUBSTR(Date,1,7) BETWEEN '2024-01' AND '2026-06'
  AND Product IN ('ADAP','ADMP','ADSSP','ELA')
GROUP BY 1, 2
ORDER BY 1, 2
""")

# ── P3: Top SEO URLs for Saudi Arabia (google organic, salesleads_qt) ────────────────
# Identifies which /it/ and global pages are driving leads (and which are declining)
run("P3 · Top SEO URLs Saudi Arabia — leads + convs by URL (salesleads_qt)", f"""
SELECT
  REGEXP_REPLACE(
    REGEXP_EXTRACT(LOWER(FIRST_SRC_URL_CLEANED), r'^(/[^?#]*)'),
    r'/$', '') AS url_path,
  CASE WHEN REGEXP_CONTAINS(LOWER(FIRST_SRC_URL_CLEANED), r'^/it(/|$)')
       THEN 'Local /it/' ELSE 'Global' END AS page_type,
  COUNT(DISTINCT IF(SUBSTR(Created_Time,8,4)='2024', ID, NULL)) AS leads_2024,
  COUNT(DISTINCT IF(SUBSTR(Created_Time,8,4)='2025', ID, NULL)) AS leads_2025,
  COUNT(DISTINCT IF(SAFE.PARSE_DATE('%d %b %Y', SUBSTR(Created_Time,1,11))
                   BETWEEN '2026-01-01' AND '2026-06-30', ID, NULL)) AS leads_h12026,
  COUNT(DISTINCT IF(Conversion='converted'
                    AND SUBSTR(Created_Time,8,4)='2024', ID, NULL)) AS conv_2024,
  COUNT(DISTINCT IF(Conversion='converted'
                    AND SUBSTR(Created_Time,8,4)='2025', ID, NULL)) AS conv_2025,
  COUNT(DISTINCT IF(Conversion='converted'
                    AND SAFE.PARSE_DATE('%d %b %Y', SUBSTR(Created_Time,1,11))
                       BETWEEN '2026-01-01' AND '2026-06-30', ID, NULL)) AS conv_h12026
FROM `{SL}`
WHERE LOWER(COMMON_COUNTRY_NAME) = '{COUNTRY_SL}'
  AND FIRST_SRC_GRP IN ('google / organic', 'bing / organic')
  AND Junk = 'false'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND SUBSTR(Created_Time,8,4) IN ('2024','2025','2026')
  AND FIRST_SRC_URL_CLEANED IS NOT NULL
GROUP BY 1, 2
HAVING leads_2024 + leads_2025 + leads_h12026 >= 5
ORDER BY leads_2025 DESC
LIMIT 40
""")

print("\n\n✅ P2 queries complete. Add output to saudi_arabia_sem_analysis_jul2026.html.")
print("  - P1b: ADAP CPC/CTR trend → add to Google NB section or new 'Ads Performance' section")
print("  - P2a: Google NB monthly → identify when conv rate changed")
print("  - P2b: Bing NB monthly  → identify when 10.3% CR appeared")
print("  - P2c: Product monthly  → confirm ADAP collapse timing")
print("  - P3:  Top SEO URLs     → add to SEO section after local/global table")
