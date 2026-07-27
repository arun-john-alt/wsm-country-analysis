CREATE OR REPLACE TABLE country_keyword_sv AS
WITH uni AS (
  SELECT country, keyword_text,
         MAX(IF(period=DATE '2026-05-01', search_volume, NULL)) AS sv,
         MAX(IF(period=DATE '2026-04-01', search_volume, NULL)) AS prev,
         AVG(IF(period>=DATE '2026-03-01', search_volume, NULL)) AS recent3,
         AVG(IF(period BETWEEN DATE '2025-12-01' AND DATE '2026-02-01', search_volume, NULL)) AS prior3,
         MAX(IF(period=DATE '2026-05-01', bid_high, NULL)) AS bid_high,
         ANY_VALUE(cluster_keyword) AS cluster_keyword
  FROM adg_keyword_universe
  WHERE period >= DATE '2025-12-01'
  GROUP BY 1,2
),
uni_f AS (
  SELECT country, keyword_text, sv,
         CASE WHEN prev>0 THEN ROUND((sv-prev)/prev*100,0) END AS mom_pct,
         ((CASE WHEN prev>0 AND sv IS NOT NULL THEN (sv-prev)/prev ELSE 0 END >= 0.2 AND sv>=30)
          OR (prior3>0 AND recent3/prior3>=1.3 AND sv>=30)) AS surging,
         bid_high, cluster_keyword
  FROM uni WHERE sv IS NOT NULL AND sv > 0
),
fetched AS (
  SELECT f.country, f.keyword_text, f.sv, f.mom_pct, f.surging, f.bid_high,
         CAST(NULL AS STRING) AS cluster_keyword
  FROM country_keyword_sv_fetched f
  WHERE f.sv IS NOT NULL AND f.sv > 0
)
SELECT * FROM uni_f
UNION ALL
SELECT * FROM fetched
