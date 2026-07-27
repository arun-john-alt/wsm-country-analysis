-- monitor_leads_monthly — per Product×Country×Theme×month LEADS + REVENUE, split by engine.
-- Country-dependent metric: PRESALES (FS_PS_Leads/FS_PS_Revenue, Lead_Type='Mktg(SPL)Leads') for
--   US/India/UK/Canada/Australia; SALES (Valid_Sales_Leads_First_Source/FS_Sales_Revenue,
--   single Lead_Type='All Leads' partition) elsewhere. Both First-Source. Engine = Source_Medium.
-- Columns: leads_google / leads_bing / leads (=consolidated total); rev_google / rev_bing / revenue.
-- Back to 2024-01 so 2025 columns get a YoY base.
CREATE OR REPLACE TABLE `it-security-online-marketing.Ads_data_WSM.monitor_leads_monthly` AS
WITH presales AS (
  SELECT Product product, CampaignCountry country, Theme theme, SUBSTR(Date,1,7) ym, Source_Medium sm,
         SUM(FS_PS_Leads) leads, SUM(FS_PS_Revenue) revenue
  FROM `it-security-online-marketing.Google_ads_data_ajay.themes_firstlast_semroi`
  WHERE Lead_Type='Mktg(SPL)Leads'
    AND CampaignCountry IN ('United States','India','United Kingdom','Canada','Australia')
    AND Product IS NOT NULL AND Theme IS NOT NULL AND CampaignCountry IS NOT NULL AND Date >= '2024-01'
  GROUP BY 1,2,3,4,5
),
sales AS (
  SELECT Product product, CampaignCountry country, Theme theme, SUBSTR(Date,1,7) ym, Source_Medium sm,
         SUM(Valid_Sales_Leads_First_Source) leads, SUM(FS_Sales_Revenue) revenue
  FROM `it-security-online-marketing.Google_ads_data_ajay.themes_firstlast_semroi`
  WHERE Lead_Type='All Leads'
    AND CampaignCountry NOT IN ('United States','India','United Kingdom','Canada','Australia')
    AND Product IS NOT NULL AND Theme IS NOT NULL AND CampaignCountry IS NOT NULL AND Date >= '2024-01'
  GROUP BY 1,2,3,4,5
),
u AS (SELECT * FROM presales UNION ALL SELECT * FROM sales)
SELECT product, country, theme, ym,
  SUM(IF(sm='google / cpc', leads, 0))   AS leads_google,
  SUM(IF(sm='bing / cpc',   leads, 0))   AS leads_bing,
  SUM(leads)                             AS leads,        -- consolidated (Google + Bing)
  SUM(IF(sm='google / cpc', revenue, 0)) AS rev_google,
  SUM(IF(sm='bing / cpc',   revenue, 0)) AS rev_bing,
  SUM(revenue)                           AS revenue
FROM u
GROUP BY 1,2,3,4
