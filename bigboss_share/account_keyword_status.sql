CREATE OR REPLACE TABLE account_keyword_status AS
WITH tmap AS (
  SELECT CampaignName, AdGroupName, ANY_VALUE(Product) product, ANY_VALUE(Theme) theme,
         ANY_VALUE(Sub_Theme) sub_theme, ANY_VALUE(CampaignCountry) country
  FROM themes_firstlast_semroi
  WHERE CampaignCountry IS NOT NULL AND CampaignCountry NOT LIKE 'Region -%' AND CampaignCountry NOT IN ('-')
    AND Lead_Type='Mktg(SPL)Leads' AND Theme IS NOT NULL
  GROUP BY CampaignName, AdGroupName
),
camp AS (SELECT campaign_id, ANY_VALUE(campaign_name) campaign_name, ANY_VALUE(campaign_status) campaign_status FROM ads_Campaign_5419501619 GROUP BY campaign_id),
adg AS (SELECT ad_group_id, ANY_VALUE(ad_group_name) ad_group_name, ANY_VALUE(ad_group_status) ad_group_status FROM ads_AdGroup_5419501619 GROUP BY ad_group_id),
kw AS (
  SELECT LOWER(TRIM(ad_group_criterion_keyword_text)) kw, ad_group_criterion_keyword_match_type mt,
         ad_group_criterion_status status, ad_group_criterion_quality_info_quality_score qs,
         campaign_id, ad_group_id
  FROM ads_Keyword_5419501619
  WHERE _DATA_DATE=(SELECT MAX(_DATA_DATE) FROM ads_Keyword_5419501619)
    AND ad_group_criterion_negative=false AND ad_group_criterion_keyword_text IS NOT NULL
)
SELECT t.country, t.product, t.theme, t.sub_theme, c.campaign_name, a.ad_group_name,
       kw.kw AS keyword_text, kw.mt AS match_type, kw.status, kw.qs AS quality_score,
       a.ad_group_status, c.campaign_status,
       -- serving = the Google Ads "Eligible"/"Eligible (limited)" state: criterion AND ad group AND campaign all enabled.
       -- A keyword whose ad group or campaign is paused is NOT actually serving in that country.
       (kw.status='ENABLED' AND a.ad_group_status='ENABLED' AND c.campaign_status='ENABLED') AS serving
FROM kw
JOIN camp c ON kw.campaign_id=c.campaign_id
JOIN adg a ON kw.ad_group_id=a.ad_group_id
JOIN tmap t ON c.campaign_name=t.CampaignName AND a.ad_group_name=t.AdGroupName
