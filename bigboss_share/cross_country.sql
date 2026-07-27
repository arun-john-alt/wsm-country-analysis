CREATE OR REPLACE TABLE cross_country_suggestions AS
WITH csv AS (
  SELECT country, keyword_text, sv, mom_pct, surging, bid_high,
         COALESCE(cluster_keyword, keyword_text) AS cluster_keyword
  FROM country_keyword_sv
),
conv AS (  -- countries where the EXACT search term actually drove conversions (real proof)
  SELECT keyword_text, STRING_AGG(DISTINCT country, ', ' ORDER BY country) AS conv_countries
  FROM adg_keyword_universe WHERE conversions > 0 GROUP BY keyword_text
),
present_country AS (SELECT DISTINCT country, keyword_text FROM account_keyword_status),
-- serving_country = countries where the keyword is genuinely Eligible (criterion + ad group + campaign all enabled),
-- NOT merely criterion-enabled under a paused ad group/campaign.
serving_country AS (SELECT DISTINCT country, keyword_text FROM account_keyword_status WHERE serving),
cc_map AS (SELECT DISTINCT campaign_name, country FROM account_keyword_status),
histqs AS (
  SELECT country, kw, last_qs FROM (
    SELECT m.country, LOWER(TRIM(q.keyword_text)) AS kw, q.quality_score AS last_qs,
           ROW_NUMBER() OVER (PARTITION BY m.country, LOWER(TRIM(q.keyword_text)) ORDER BY q.week_start DESC) rn
    FROM qs_weekly_history q JOIN cc_map m ON q.campaign_name=m.campaign_name
    WHERE q.quality_score IS NOT NULL AND q.quality_score > 0
  ) WHERE rn=1
),
status_clean AS (
  SELECT * FROM account_keyword_status
  WHERE match_type IN ('EXACT','PHRASE')
    AND NOT REGEXP_CONTAINS(LOWER(theme), r'brand|display|competitor|remarket|content|generic|others|^azure$|^iam$')
    AND NOT REGEXP_CONTAINS(keyword_text, r'\b(ad manager|ads manager|password manager|authenticator|museum|fortnite|facebook|instagram|idrive)\b')
    -- intent intelligence: drop definitional/DIY, KEEP how-to/where-to problem queries (map to product capability)
    AND NOT REGEXP_CONTAINS(keyword_text, r'^what (is|are|was|were|does|do|did)\b')
    AND NOT REGEXP_CONTAINS(keyword_text, r'\b(meaning|definition|define|tutorial|explained|wikipedia|wiki|reddit|salary|jobs|certification|course|examples?|powershell|cmdlet)\b')
    AND keyword_text NOT IN (SELECT keyword_text FROM keyword_intent WHERE NOT keep)
    AND keyword_text NOT IN ('active directory','audit','azure','powershell','backup','microsoft 365','office 365','mfa','2fa','sso','siem','dns','rbac','active directory management','ad management')
),
golden_raw AS (
  SELECT k.product, k.theme, k.keyword_text,
         STRING_AGG(DISTINCT k.country, ', ') AS src, COUNT(DISTINCT k.country) AS n_src
  FROM status_clean k
  JOIN proven_themes p USING(product, country, theme)
  WHERE k.serving AND k.country IN ('United States','United Kingdom','Canada','Australia','India')
  GROUP BY 1,2,3
),
golden AS (  -- one PRIMARY theme per (product, keyword): where it is most established
  SELECT * FROM golden_raw
  QUALIFY ROW_NUMBER() OVER (PARTITION BY product, keyword_text ORDER BY n_src DESC, theme) = 1
),
gap AS (
  SELECT pt.country, g.product, g.theme, g.keyword_text AS keyword, 'GAP' AS engine,
         csv.sv, csv.bid_high, csv.cluster_keyword, g.src AS source_countries, g.n_src,
         CAST(NULL AS INT64) AS last_qs, csv.mom_pct, csv.surging, cv.conv_countries,
         COALESCE(kt.tier, CASE WHEN csv.sv>=300 THEN 'T2' ELSE 'T3' END) AS tier,
         CAST(csv.sv * (1 + g.n_src) AS FLOAT64) AS priority
  FROM golden g
  JOIN proven_themes pt ON pt.product=g.product AND pt.theme=g.theme
   AND pt.country IN ('United States','United Kingdom','Canada','Australia','India')
  JOIN csv ON csv.country=pt.country AND csv.keyword_text=g.keyword_text AND csv.sv BETWEEN 20 AND 25000
  LEFT JOIN present_country pc ON pc.country=pt.country AND pc.keyword_text=g.keyword_text
  LEFT JOIN conv cv ON cv.keyword_text=g.keyword_text
  LEFT JOIN keyword_tier kt ON kt.product=g.product AND kt.cluster_keyword=csv.cluster_keyword
  WHERE pc.keyword_text IS NULL AND COALESCE(kt.tier,'T3') != 'DROP'
),
paused_cnt AS (  -- prevalence of each NOT-serving keyword per theme (# placements).
  -- NOT serving = criterion paused/removed OR its ad group/campaign paused (so it isn't actually running).
  SELECT country, product, theme, keyword_text, COUNT(*) AS n_ag
  FROM status_clean WHERE NOT serving
  GROUP BY 1,2,3,4
),
react_raw AS (
  SELECT pc.country, pc.product, pc.theme, pc.keyword_text AS keyword,
         csv.sv, csv.bid_high, csv.cluster_keyword, hq.last_qs, csv.mom_pct, csv.surging, cv.conv_countries,
         kt.tier AS kt_tier, pc.n_ag, p.leads_36mo
  FROM paused_cnt pc
  JOIN proven_themes p ON p.product=pc.product AND p.country=pc.country AND p.theme=pc.theme
  JOIN csv ON csv.country=pc.country AND csv.keyword_text=pc.keyword_text AND csv.sv BETWEEN 20 AND 25000
  LEFT JOIN serving_country ec ON ec.country=pc.country AND ec.keyword_text=pc.keyword_text
  LEFT JOIN histqs hq ON hq.country=pc.country AND hq.kw=pc.keyword_text
  LEFT JOIN conv cv ON cv.keyword_text=pc.keyword_text
  LEFT JOIN keyword_tier kt ON kt.product=pc.product AND kt.cluster_keyword=csv.cluster_keyword
  WHERE ec.keyword_text IS NULL AND COALESCE(kt.tier,'T3') != 'DROP'
),
reactivate AS (  -- one PRIMARY theme per (country, product, keyword): where it is most present
  SELECT country, product, theme, keyword, 'REACTIVATE' AS engine,
         sv, bid_high, cluster_keyword, CAST(NULL AS STRING) AS source_countries,
         CAST(NULL AS INT64) AS n_src, last_qs, mom_pct, surging, conv_countries,
         COALESCE(kt_tier, CASE WHEN sv >= 300 THEN 'T2' ELSE 'T3' END) AS tier,
         CAST(sv AS FLOAT64) AS priority
  FROM react_raw
  QUALIFY ROW_NUMBER() OVER (PARTITION BY country, product, keyword ORDER BY n_ag DESC, leads_36mo DESC, theme) = 1
)
SELECT * FROM gap
UNION ALL
SELECT * FROM reactivate
