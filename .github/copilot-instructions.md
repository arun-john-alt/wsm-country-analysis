# Copilot Workspace Instructions — WSM Monitor

## Trigger: "analyse [country]" or "analyze [country]"

When the user asks to analyse any country for SEM (e.g. "analyse Spain", "next country is Netherlands"):

1. **Read the SOP first** — use the memory tool to read `/memories/repo/wsm-monitor-analysis.md` before doing anything else. Do not proceed without reading it.

2. **Follow the SOP exactly** — Steps 1–5 are mandatory. Do not skip sections, do not skip Pass 2 queries.

3. **Adapt the query templates** — copy `/Users/arun-8846/Downloads/Monitor/analysis/queries/italy_sem_queries.py` and `/Users/arun-8846/Downloads/Monitor/analysis/queries/italy_sem_queries_p2.py`, change `COUNTRY_ROI`, `COUNTRY_SL`, and the SEO local path prefix. Save as `analysis/queries/<country_lower>_sem_queries.py` and `analysis/queries/<country_lower>_sem_queries_p2.py`.

4. **Run Pass 1 first** — ask the user to run the script and share output (or run it directly if terminal access is available). Always run Q15+Q16 diagnostic queries first to validate brand theme names before writing the report.

5. **Run Pass 2** — after Pass 1 output is confirmed, run the P2 script for monthly trends, ADAP CPC/CTR, and top SEO URLs.

6. **Build the HTML** — use `/Users/arun-8846/Downloads/Monitor/analysis/reports/italy_sem_analysis_jul2026.html` as the template. All 13 sections from the SOP are required. Save as `analysis/reports/<country_lower>_sem_analysis_jul2026.html`.

7. **Apply all writing rules from the SOP** before finalising — run the Step 5 checklist mentally before declaring done.

## Always true in this workspace

- Currency: themes Cost = USD, ads_AdGroupBasicStats cost_micros/1e6 = INR
- Brand filter: `Theme IN ('Branding','Log360 - Branding')` — check Q16 for additional brand themes
- Source medium: always `COALESCE(Source_Medium, Source___Medium)`
- Conv rate denominator: themes leads (not salesleads count)
- Italy-specific: ELA/LOG360/LOG360CLOUD = one merged product
- Never use budget expansion language — reallocation only
- No cross-country comparison text in callouts
