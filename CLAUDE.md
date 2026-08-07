# WSM Monitor — Claude Code Instructions

## Trigger: "analyse [country]" or "analyze [country]"

When asked to analyse any country for SEM:

1. Read `/memories/repo/wsm-monitor-analysis.md` first using the memory tool. Do not proceed without reading it.
2. Follow the SOP in that file exactly — Steps 1–5 are all mandatory.
3. Clone `analysis/queries/italy_sem_queries.py` and `analysis/queries/italy_sem_queries_p2.py`, adapt `COUNTRY_ROI`/`COUNTRY_SL`/SEO path prefix, save as `analysis/queries/<country>_sem_queries.py` and `analysis/queries/<country>_sem_queries_p2.py`.
4. Run Pass 1. Run Q15+Q16 diagnostics first to validate brand theme names.
5. Run Pass 2 after Pass 1 output is confirmed.
6. Build HTML using `analysis/reports/italy_sem_analysis_jul2026.html` as template. All 13 sections required. Save as `analysis/reports/<country>_sem_analysis_jul2026.html`.
7. Run the Step 5 checklist before declaring done.

## Hard rules (always apply)

- `COALESCE(Source_Medium, Source___Medium)` — never Source_Medium alone
- Brand filter (full list — use exactly this in ALL queries, themes + salesleads_qt): `Theme IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')` — all 7 variants confirmed Jul 2026. Never use only 2-item list again.
- Conv rate denominator = themes leads, not salesleads count
- Spend = USD (themes), CPC = INR (ads_AdGroupBasicStats)
- Sales countries: `Valid_Sales_Leads_First_Source`, `Lead_Type='All Leads'`
- **salesleads_qt standard filter (ALWAYS use all 4 — confirmed Aug 2026 against CRM):**
  1. `Junk = 'false'`
  2. `PRODUCT_GROUP = 'AD_GROUP'`
  3. `User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')` — excludes existing customers and existing leads
  4. `isHaveToBeRemoved = 'Non Junk Email'` — Mail type = Non Junk Email only
  - Dedup: `COUNT(DISTINCT Email)` for lead counts, `COUNT(DISTINCT IF(Conversion='converted', Email, NULL))` for convs — matches CRM "Dist.email" view
  - Do NOT use `COUNT(DISTINCT ID)` for lead totals — it overcounts (one person can have multiple lead IDs)
  - The `PRODUCT_GROUP = 'AD_GROUP'` filter already existed; User_Type and isHaveToBeRemoved are the NEW additions verified Aug 2026
- Campaign type exclusion (Q8–Q12, Q16): `FIRST_SRC_CAMPAIGN_TYPE NOT IN ('Display','Performance Max') OR FIRST_SRC_CAMPAIGN_TYPE IS NULL` — excludes both Display and PMax from SEM conv queries. Search only.
- Reallocation language only — never "increase budget" / "scale up" / "expand allocation"
- No cross-country comparison text in callout copy
- ELA + LOG360 + LOG360CLOUD = one merged product row (global rule — all countries, not Italy-only)
- SEO local vs global split (Q14): Local = ANY URL matching `^/(br|fr|de|latam|es|au|za|it|nl|jp|in|uk)(/|$)` — includes /it/, /de/, /fr/ etc. Global = everything else. Not just the country-specific prefix, ALL country pages are "local"
- Q14 organic filter: `FIRST_SRC_GRP IN ('google / organic','bing / organic','organic / (not set)','organic')` — always include `bing / organic` (France has 129, Germany has 103, Italy has 44+ Bing organic leads — omitting it understates SEO totals)
- ELA group H1 2026 spend: EXCLUDE from `SUM(Cost)` — Log360 spend is duplicated in the themes table for H1 2026. Show ELA/LOG360/LOG360CLOUD H1 2026 CPL and Spend as `—` with footnote `†ELA/LOG360/LOG360CLOUD spend excluded in H1 2026 — duplicate attribution in themes data`. 2024 and 2025 spend is clean and reportable.
- Channel attribution uses TWO columns — always report both: `FIRST_SRC_GRP` (salesleads_qt) for full-funnel channel mix (Q13); `NEW_TRAFFIC_SRC_GRP` (salesleads_qt) for the new grouping view (Q17). Both are required in every country report. Q17 section goes immediately before the footer.
- Q17 template (copy verbatim for every new country, change COUNTRY_SL only): `SELECT {YR_SL} AS yr, NEW_TRAFFIC_SRC_GRP, COUNT(DISTINCT Email) AS total_leads, COUNT(DISTINCT IF(Conversion='converted', Email, NULL)) AS convs FROM \`{SL}\` WHERE LOWER(COMMON_COUNTRY_NAME) = '{COUNTRY_SL}' AND Junk = 'false' AND PRODUCT_GROUP = 'AD_GROUP' AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead') AND isHaveToBeRemoved = 'Non Junk Email' AND SUBSTR(Created_Time,8,4) IN ('2024','2025','2026') GROUP BY 1, 2 HAVING yr != 'other' ORDER BY NEW_TRAFFIC_SRC_GRP, yr`

## ROW Monthly Report (trigger: "ROW monthly report" or "trigger the monthly report")

Script: `analysis/queries/row_monthly_report.py`

Run command:
```
python3 analysis/queries/row_monthly_report.py --month YYYY-MM --ytd
```
- `--month`: target month (e.g. `2026-07`). Defaults to `run.month` in `wsm-monitor/config.yaml` — update that first.
- `--ytd`: adds a second tab for Jan–CUR vs prior year. Always include it.

Before running: update `wsm-monitor/config.yaml` → `run.month` to the target month.
Also run freshness check first: `python3 wsm-monitor/check_freshness.py monthly`

Output file: `<mon><year>_yoy.xlsx` in the repo root (gitignored — share via Drive).

Format (canonical — do not change without updating the script):
- Row 1: Big yellow title "Jul 2026 (vs Jul '25)"
- Row 2: 3 section group headers — pink / light-blue / green
- Row 3: Column headers
- Row 4+: Data rows. YoY% is INLINE in each cell (bold number + grey %). Green fill ≥+10%, pink fill ≤-10%.
- ROW markets only (US, India, UK, Canada, Australia are presales — NOT included)

3 data sections:
  [A] Pink  — All leads except Events/Third Party/Others: salesleads_qt with full 4-filter standard
  [B] Blue  — All leads: salesleads_qt with Junk='false' + PRODUCT_GROUP='AD_GROUP' only
  [C] Green — SEM leads: themes table (Non-Brand, Google+Bing SEM spend/leads; salesleads_qt SEM convs)

DM Region display names and order (must match exactly):
Germany, Netherlands, Switzerland, Belgium, France, Italy, United Arab Emirates, Saudi Arabia, Turkey,
Spain, Brazil, Mexico, Rest Of LATAM, South Africa, Israel, Rest Of Europe, Poland, Rest Of MEA, Rest Of APAC, Singapore
