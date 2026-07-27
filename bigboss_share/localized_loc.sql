-- Localized gap suggestions for FR/ES/MX/BR/IT/NL/PL/TR from dialect-seeded KP ideas.
-- Keep: net-new (not in that country's account) + has SV. Drop: competitor brands, definitional junk.
CREATE OR REPLACE TABLE localized_loc AS
WITH acct AS (SELECT DISTINCT country, LOWER(TRIM(keyword_text)) AS kw FROM account_keyword_status),
cand AS (
  SELECT product, country, theme, idea_keyword AS keyword, MAX(avg_monthly_searches) AS sv
  FROM kw_exp_ideas_raw_loc
  WHERE NOT is_seed AND avg_monthly_searches >= 20
    AND NOT REGEXP_CONTAINS(idea_keyword, r'\b(veeam|acronis|synology|codetwo|datto|rubrik|cohesity|barracuda|avepoint|splunk|qradar|sentinel|graylog|logrhythm|wazuh|kiwi|solarwinds|prtg|nagios|netwrix|quest|varonis|lepide|specops|okta|auth0|duo|cisco|fortinet|sophos)\b')
    AND NOT REGEXP_CONTAINS(idea_keyword, r'\b(meaning|definition|tutorial|wikipedia|wiki|reddit|salary|jobs)\b')
  GROUP BY 1,2,3,4
),
gap AS (
  SELECT c.* FROM cand c
  LEFT JOIN acct a ON a.country=c.country AND a.kw=LOWER(TRIM(c.keyword))
  WHERE a.kw IS NULL
)
SELECT product, country, theme, keyword,
       CASE WHEN sv>=200 THEN 'T1' WHEN sv>=50 THEN 'T2' ELSE 'T3' END AS tier, sv
FROM gap
