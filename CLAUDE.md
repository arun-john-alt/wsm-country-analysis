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
- Brand filter: `Theme IN ('Branding','Log360 - Branding')` — verify Q16 for extras
- Conv rate denominator = themes leads, not salesleads count
- Spend = USD (themes), CPC = INR (ads_AdGroupBasicStats)
- Sales countries: `Valid_Sales_Leads_First_Source`, `Lead_Type='All Leads'`
- Reallocation language only — never "increase budget" / "scale up" / "expand allocation"
- No cross-country comparison text in callout copy
- Italy only: ELA + LOG360 + LOG360CLOUD = one merged product row
