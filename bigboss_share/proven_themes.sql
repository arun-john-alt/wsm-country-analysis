CREATE OR REPLACE TABLE proven_themes AS
SELECT Product product, CampaignCountry country, Theme theme,
       ROUND(SUM(CASE WHEN CampaignCountry IN ('United States','United Kingdom','India','Australia','Canada')
                      THEN FS_PS_Leads ELSE Valid_Sales_Leads_First_Source END)) leads_36mo,
       ROUND(SUM(Clicks)) clicks_36mo
FROM themes_firstlast_semroi
WHERE CampaignCountry IS NOT NULL AND CampaignCountry NOT LIKE 'Region -%' AND CampaignCountry NOT IN ('-')
  AND Lead_Type='Mktg(SPL)Leads' AND Theme IS NOT NULL
  AND PARSE_DATE('%Y-%m-%d', Date) >= DATE_SUB(CURRENT_DATE(), INTERVAL 36 MONTH)
GROUP BY 1,2,3
HAVING leads_36mo > 0
