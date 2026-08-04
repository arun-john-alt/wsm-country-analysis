"""
Saudi Arabia SEM Deep-Dive — BigQuery Queries
Run: python3 analysis/queries/saudi_arabia_sem_queries.py
Produces all numbers needed for saudi_arabia_sem_analysis_jul2026.html.
Same patterns as Italy / France / Germany / Brazil (Jul 2026).

Logic verified in analysis_patterns.py:
  - Source: COALESCE(Source_Medium, Source___Medium)
  - Brand: Theme IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding') — verify extras via Q16
  - Leads: Valid_Sales_Leads_First_Source  (Saudi Arabia = sales country)
  - Lead_Type filter: 'All Leads'
  - Conversions: salesleads_qt COUNT(DISTINCT IF(Conversion='converted', ID, NULL))
  - Dedup: COUNT(DISTINCT Email) — CRM-verified filter Aug 2026
  - Spend: Cost (USD, from themes)
  - DRI: Kowsik
"""
import sys, textwrap
sys.path.insert(0, '/Users/arun-8846/Downloads/Monitor/wsm-monitor')
from wsm_cfg import G, PROJ, bq_client, ACCT

bq = bq_client()
ROI = f"{PROJ}.{G.split('.')[-1]}.themes_firstlast_semroi"
SL  = f"{PROJ}.sales_presales_leads_no_pi.salesleads_qt"

COUNTRY_ROI = "Saudi Arabia"
COUNTRY_SL  = "saudi arabia"
LEAD_TYPE   = "All Leads"

# ── helpers ──────────────────────────────────────────────────────────────────

def run(label, sql):
    print(f"\n{'═'*70}")
    print(f"  {label}")
    print('═'*70)
    rows = list(bq.query(sql).result())
    if not rows:
        print("  (no rows)")
        return rows
    # print header
    keys = list(rows[0].keys())
    widths = [max(len(str(k)), max(len(str(r[k])) for r in rows)) for k in keys]
    fmt = "  " + "  ".join(f"{{:<{w}}}" for w in widths)
    print(fmt.format(*keys))
    print("  " + "  ".join("─"*w for w in widths))
    for r in rows:
        print(fmt.format(*[str(r[k]) for k in keys]))
    return rows

# ── period expressions ────────────────────────────────────────────────────────

YR_ROI = """CASE
  WHEN SUBSTR(Date,1,7) BETWEEN '2026-01' AND '2026-06' THEN 'H12026'
  WHEN SUBSTR(Date,1,4)='2025' THEN '2025'
  WHEN SUBSTR(Date,1,4)='2024' THEN '2024'
  ELSE 'other' END"""

YR_SL = """CASE
  WHEN SAFE.PARSE_DATE('%d %b %Y', SUBSTR(Created_Time,1,11))
       BETWEEN '2026-01-01' AND '2026-06-30' THEN 'H12026'
  WHEN SUBSTR(Created_Time,8,4)='2025' THEN '2025'
  WHEN SUBSTR(Created_Time,8,4)='2024' THEN '2024'
  ELSE 'other' END"""

# ── Q1: Google Non-Brand — Overall leads + spend by year ─────────────────────
run("Q1 · Google Non-Brand — Overall leads + spend (themes)", f"""
SELECT
  {YR_ROI} AS yr,
  ROUND(SUM(Valid_Sales_Leads_First_Source))   AS leads,
  ROUND(SUM(Cost), 0)                          AS spend_usd
FROM `{ROI}`
WHERE CampaignCountry = '{COUNTRY_ROI}'
  AND COALESCE(Source_Medium, Source___Medium) = 'google / cpc'
  AND Theme NOT IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
  AND Lead_Type = '{LEAD_TYPE}'
  AND SUBSTR(Date,1,4) IN ('2024','2025','2026')
GROUP BY 1
HAVING yr != 'other'
ORDER BY yr
""")

# ── Q2: Google Non-Brand — leads + spend by year × product ───────────────────
run("Q2 · Google Non-Brand — by product (themes)", f"""
SELECT
  {YR_ROI} AS yr,
  Product,
  ROUND(SUM(Valid_Sales_Leads_First_Source))   AS leads,
  ROUND(SUM(Cost), 0)                          AS spend_usd
FROM `{ROI}`
WHERE CampaignCountry = '{COUNTRY_ROI}'
  AND COALESCE(Source_Medium, Source___Medium) = 'google / cpc'
  AND Theme NOT IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
  AND Lead_Type = '{LEAD_TYPE}'
  AND SUBSTR(Date,1,4) IN ('2024','2025','2026')
GROUP BY 1, 2
HAVING yr != 'other'
ORDER BY Product, yr
""")

# ── Q3: Google Brand — Overall leads + spend by year ─────────────────────────
run("Q3 · Google Brand — Overall leads + spend (themes)", f"""
SELECT
  {YR_ROI} AS yr,
  ROUND(SUM(Valid_Sales_Leads_First_Source))   AS leads,
  ROUND(SUM(Cost), 0)                          AS spend_usd
FROM `{ROI}`
WHERE CampaignCountry = '{COUNTRY_ROI}'
  AND COALESCE(Source_Medium, Source___Medium) = 'google / cpc'
  AND Theme IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
  AND Lead_Type = '{LEAD_TYPE}'
  AND SUBSTR(Date,1,4) IN ('2024','2025','2026')
GROUP BY 1
HAVING yr != 'other'
ORDER BY yr
""")

# ── Q4: Google Brand — by product (Product column) ───────────────────────────
run("Q4 · Google Brand — by product (themes)", f"""
SELECT
  {YR_ROI} AS yr,
  Product,
  ROUND(SUM(Valid_Sales_Leads_First_Source))   AS leads,
  ROUND(SUM(Cost), 0)                          AS spend_usd
FROM `{ROI}`
WHERE CampaignCountry = '{COUNTRY_ROI}'
  AND COALESCE(Source_Medium, Source___Medium) = 'google / cpc'
  AND Theme IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
  AND Lead_Type = '{LEAD_TYPE}'
  AND SUBSTR(Date,1,4) IN ('2024','2025','2026')
GROUP BY 1, 2
HAVING yr != 'other'
ORDER BY Product, yr
""")

# ── Q5: Bing Non-Brand — Overall leads + spend by year ───────────────────────
run("Q5 · Bing Non-Brand — Overall leads + spend (themes)", f"""
SELECT
  {YR_ROI} AS yr,
  ROUND(SUM(Valid_Sales_Leads_First_Source))   AS leads,
  ROUND(SUM(Cost), 0)                          AS spend_usd
FROM `{ROI}`
WHERE CampaignCountry = '{COUNTRY_ROI}'
  AND COALESCE(Source_Medium, Source___Medium) = 'bing / cpc'
  AND Theme NOT IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
  AND Lead_Type = '{LEAD_TYPE}'
  AND SUBSTR(Date,1,4) IN ('2024','2025','2026')
GROUP BY 1
HAVING yr != 'other'
ORDER BY yr
""")

# ── Q6: Bing Non-Brand — by product ──────────────────────────────────────────
run("Q6 · Bing Non-Brand — by product (themes)", f"""
SELECT
  {YR_ROI} AS yr,
  Product,
  ROUND(SUM(Valid_Sales_Leads_First_Source))   AS leads,
  ROUND(SUM(Cost), 0)                          AS spend_usd
FROM `{ROI}`
WHERE CampaignCountry = '{COUNTRY_ROI}'
  AND COALESCE(Source_Medium, Source___Medium) = 'bing / cpc'
  AND Theme NOT IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
  AND Lead_Type = '{LEAD_TYPE}'
  AND SUBSTR(Date,1,4) IN ('2024','2025','2026')
GROUP BY 1, 2
HAVING yr != 'other'
ORDER BY Product, yr
""")

# ── Q7: Bing Brand — Overall leads + spend by year ───────────────────────────
run("Q7 · Bing Brand — Overall leads + spend (themes)", f"""
SELECT
  {YR_ROI} AS yr,
  ROUND(SUM(Valid_Sales_Leads_First_Source))   AS leads,
  ROUND(SUM(Cost), 0)                          AS spend_usd
FROM `{ROI}`
WHERE CampaignCountry = '{COUNTRY_ROI}'
  AND COALESCE(Source_Medium, Source___Medium) = 'bing / cpc'
  AND Theme IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
  AND Lead_Type = '{LEAD_TYPE}'
  AND SUBSTR(Date,1,4) IN ('2024','2025','2026')
GROUP BY 1
HAVING yr != 'other'
ORDER BY yr
""")

# ── Q8: SEM Conversions — by year × engine × brand flag (salesleads_qt) ──────
# Gives conv cross-check to validate themes-based CPL
run("Q8 · SEM Conversions — by year × engine × brand (salesleads_qt)", f"""
SELECT
  {YR_SL} AS yr,
  FIRST_SRC_GRP                                           AS engine,
  CASE WHEN FIRST_SRC_THEME IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
       THEN 'Brand' ELSE 'Non-Brand' END                  AS brand_flag,
  COUNT(DISTINCT Email)                                    AS total_leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))  AS convs
FROM `{SL}`
WHERE LOWER(COMMON_COUNTRY_NAME) = '{COUNTRY_SL}'
  AND FIRST_SRC_GRP IN ('google / cpc','bing / cpc')
  AND FIRST_SRC_CAMPAIGN_TYPE NOT IN ('Display','Performance Max') OR FIRST_SRC_CAMPAIGN_TYPE IS NULL
  AND Junk = 'false'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND SUBSTR(Created_Time,8,4) IN ('2024','2025','2026')
GROUP BY 1, 2, 3
HAVING yr != 'other'
ORDER BY engine, brand_flag, yr
""")

# ── Q9: SEM Conversions — by year × engine × product (salesleads_qt) ─────────
# Non-Brand product-level convs; FIRST_SRC_THEME → product
run("Q9 · Google Non-Brand Convs — by year × product (salesleads_qt)", f"""
SELECT
  {YR_SL} AS yr,
  -- Derive product from FIRST_SRC_THEME (theme contains product prefix in ManageEngine naming)
  CASE
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADAP|AD AUDIT') THEN 'ADAP'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADMP|AD MANAGER') THEN 'ADMP'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADSSP|AD SELF') THEN 'ADSSP'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'\bELA\b|ENDPOINT') THEN 'ELA'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'LOG360') THEN 'LOG360'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADPLUS|AD\+') THEN 'ADPlus'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'SPMP|PASSWORD') THEN 'SPMP'
    ELSE FIRST_SRC_THEME
  END AS product_group,
  COUNT(DISTINCT Email)                                    AS total_leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))  AS convs
FROM `{SL}`
WHERE LOWER(COMMON_COUNTRY_NAME) = '{COUNTRY_SL}'
  AND FIRST_SRC_GRP = 'google / cpc'
  AND FIRST_SRC_THEME NOT IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
  AND FIRST_SRC_CAMPAIGN_TYPE NOT IN ('Display','Performance Max') OR FIRST_SRC_CAMPAIGN_TYPE IS NULL
  AND Junk = 'false'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND SUBSTR(Created_Time,8,4) IN ('2024','2025','2026')
GROUP BY 1, 2
HAVING yr != 'other'
ORDER BY product_group, yr
""")

# ── Q10: Google Brand Convs — by year × product (salesleads_qt) ──────────────
run("Q10 · Google Brand Convs — by year × product (salesleads_qt)", f"""
SELECT
  {YR_SL} AS yr,
  CASE
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADAP|AD AUDIT') THEN 'ADAP'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADMP|AD MANAGER') THEN 'ADMP'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADSSP|AD SELF') THEN 'ADSSP'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'\bELA\b|ENDPOINT') THEN 'ELA'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'LOG360') THEN 'LOG360'
    ELSE FIRST_SRC_THEME
  END AS product_group,
  COUNT(DISTINCT Email)                                    AS total_leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))  AS convs
FROM `{SL}`
WHERE LOWER(COMMON_COUNTRY_NAME) = '{COUNTRY_SL}'
  AND FIRST_SRC_GRP = 'google / cpc'
  AND FIRST_SRC_THEME IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
  AND FIRST_SRC_CAMPAIGN_TYPE NOT IN ('Display','Performance Max') OR FIRST_SRC_CAMPAIGN_TYPE IS NULL
  AND Junk = 'false'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND SUBSTR(Created_Time,8,4) IN ('2024','2025','2026')
GROUP BY 1, 2
HAVING yr != 'other'
ORDER BY product_group, yr
""")

# ── Q11: Bing NB Convs — by year × product ───────────────────────────────────
run("Q11 · Bing Non-Brand Convs — by year × product (salesleads_qt)", f"""
SELECT
  {YR_SL} AS yr,
  CASE
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADAP|AD AUDIT') THEN 'ADAP'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADMP|AD MANAGER') THEN 'ADMP'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'ADSSP|AD SELF') THEN 'ADSSP'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'\bELA\b|ENDPOINT') THEN 'ELA'
    WHEN REGEXP_CONTAINS(UPPER(FIRST_SRC_THEME), r'LOG360') THEN 'LOG360'
    ELSE FIRST_SRC_THEME
  END AS product_group,
  COUNT(DISTINCT Email)                                    AS total_leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))  AS convs
FROM `{SL}`
WHERE LOWER(COMMON_COUNTRY_NAME) = '{COUNTRY_SL}'
  AND FIRST_SRC_GRP = 'bing / cpc'
  AND FIRST_SRC_THEME NOT IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
  AND FIRST_SRC_CAMPAIGN_TYPE NOT IN ('Display','Performance Max') OR FIRST_SRC_CAMPAIGN_TYPE IS NULL
  AND Junk = 'false'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND SUBSTR(Created_Time,8,4) IN ('2024','2025','2026')
GROUP BY 1, 2
HAVING yr != 'other'
ORDER BY product_group, yr
""")

# ── Q12: Bing Brand Convs — by year ──────────────────────────────────────────
run("Q12 · Bing Brand Convs — by year (salesleads_qt)", f"""
SELECT
  {YR_SL} AS yr,
  COUNT(DISTINCT Email)                                    AS total_leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))  AS convs
FROM `{SL}`
WHERE LOWER(COMMON_COUNTRY_NAME) = '{COUNTRY_SL}'
  AND FIRST_SRC_GRP = 'bing / cpc'
  AND FIRST_SRC_THEME IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')
  AND FIRST_SRC_CAMPAIGN_TYPE NOT IN ('Display','Performance Max') OR FIRST_SRC_CAMPAIGN_TYPE IS NULL
  AND Junk = 'false'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND SUBSTR(Created_Time,8,4) IN ('2024','2025','2026')
GROUP BY 1
HAVING yr != 'other'
ORDER BY yr
""")

# ── Q13: All-channel leads + convs (salesleads_qt) ───────────────────────────
# Shows SEO, Direct, Referral, etc. for the "Other Channels" section
run("Q13 · All channels — leads + convs by year × FIRST_SRC_GRP (salesleads_qt)", f"""
SELECT
  {YR_SL} AS yr,
  FIRST_SRC_GRP,
  COUNT(DISTINCT Email)                                    AS total_leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))  AS convs
FROM `{SL}`
WHERE LOWER(COMMON_COUNTRY_NAME) = '{COUNTRY_SL}'
  AND Junk = 'false'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND SUBSTR(Created_Time,8,4) IN ('2024','2025','2026')
GROUP BY 1, 2
HAVING yr != 'other'
ORDER BY FIRST_SRC_GRP, yr
""")

# ── Q14: SEO — local /sa/ vs global (salesleads_qt) ──────────────────────────
# Local = ANY country-code prefix URL — /it/, /de/, /fr/, /br/ etc. all count.
# Saudi Arabia has no /sa/ pages; this query will confirm 0 local rows.
run("Q14 · SEO — local vs global pages by year (salesleads_qt)", f"""
SELECT
  {YR_SL} AS yr,
  CASE WHEN REGEXP_CONTAINS(LOWER(FIRST_SRC_URL_CLEANED),
                             r'^/(br|fr|de|latam|es|au|za|it|nl|jp|in|uk)(/|$)')
       THEN 'Local (any country page)' ELSE 'Global (EN)' END AS page_type,
  COUNT(DISTINCT Email)                                    AS total_leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))  AS convs
FROM `{SL}`
WHERE LOWER(COMMON_COUNTRY_NAME) = '{COUNTRY_SL}'
  AND FIRST_SRC_GRP IN ('google / organic','bing / organic','organic / (not set)','organic')
  AND Junk = 'false'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND SUBSTR(Created_Time,8,4) IN ('2024','2025','2026')
GROUP BY 1, 2
HAVING yr != 'other'
ORDER BY page_type, yr
""")

# ── Q15: DIAGNOSTIC — unique FIRST_SRC_GRP values for Saudi Arabia (salesleads_qt) ──
# Run once to see exact channel group names used for Saudi Arabia
run("Q15 · DIAGNOSTIC — distinct FIRST_SRC_GRP values for Saudi Arabia", f"""
SELECT
  FIRST_SRC_GRP,
  COUNT(DISTINCT Email) AS leads
FROM `{SL}`
WHERE LOWER(COMMON_COUNTRY_NAME) = '{COUNTRY_SL}'
  AND Junk = 'false'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND SUBSTR(Created_Time,8,4) IN ('2024','2025','2026')
GROUP BY 1
ORDER BY leads DESC
""")

# ── Q16: DIAGNOSTIC — distinct FIRST_SRC_THEME values for Saudi Arabia SEM ────────
# Run once to confirm product theme names — validate Q9/Q10/Q11 CASE mapping
run("Q16 · DIAGNOSTIC — distinct FIRST_SRC_THEME for Saudi Arabia SEM (salesleads_qt)", f"""
SELECT
  FIRST_SRC_GRP,
  FIRST_SRC_THEME,
  COUNT(DISTINCT Email) AS leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL)) AS convs
FROM `{SL}`
WHERE LOWER(COMMON_COUNTRY_NAME) = '{COUNTRY_SL}'
  AND FIRST_SRC_GRP IN ('google / cpc','bing / cpc')
  AND Junk = 'false'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND SUBSTR(Created_Time,8,4) IN ('2024','2025','2026')
GROUP BY 1, 2
ORDER BY leads DESC
LIMIT 50
""")

# ── Q17: All-channel leads + convs by NEW_TRAFFIC_SRC_GRP (salesleads_qt) ────
# Reports the newer traffic source grouping separately — cross-reference with Q13 (FIRST_SRC_GRP)
run("Q17 · All channels — leads + convs by year × NEW_TRAFFIC_SRC_GRP (salesleads_qt)", f"""
SELECT
  {YR_SL} AS yr,
  NEW_TRAFFIC_SRC_GRP,
  COUNT(DISTINCT Email)                                    AS total_leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))  AS convs
FROM `{SL}`
WHERE LOWER(COMMON_COUNTRY_NAME) = '{COUNTRY_SL}'
  AND Junk = 'false'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND SUBSTR(Created_Time,8,4) IN ('2024','2025','2026')
GROUP BY 1, 2
HAVING yr != 'other'
ORDER BY NEW_TRAFFIC_SRC_GRP, yr
""")

print("\n\n✅ All queries complete. Paste output into saudi_arabia_sem_analysis_jul2026.html template.")
