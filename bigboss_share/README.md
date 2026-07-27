# Big Boss — portable bundle

Generates the two "Big Boss" sheet columns (AI Keyword Recommendations + Change History)
for the Continuous Monitoring sheet, grain Product × Country × Theme.

## Configuration — supply YOUR credentials via environment variables
No credentials are bundled. `bigboss_config.py` reads everything from environment variables and,
if any required one is missing, refuses to run and prints the exact list to set. You don't edit
the file — you set these env vars (paths below auto-resolve and need nothing):

Required env vars:
```
BIGBOSS_BQ_PROJECT     your BigQuery project id
BIGBOSS_BQ_DATASET     your BigQuery dataset name
GADS_DEV_TOKEN         Google Ads developer token
GADS_CLIENT_ID         Google Ads OAuth client id
GADS_CLIENT_SECRET     Google Ads OAuth client secret
GADS_REFRESH_TOKEN     Google Ads OAuth refresh token
GADS_LOGIN_CID         Google Ads login customer id (MCC or account, digits only)
GADS_CUSTOMER_ID       Google Ads customer id to query (digits only)
ZOHO_SHEET_ID          Zoho Sheet resource id
ZOHO_CLIENT_ID         Zoho OAuth client id
ZOHO_CLIENT_SECRET     Zoho OAuth client secret
ZOHO_REFRESH_TOKEN     Zoho OAuth refresh token
```
Optional: `ZOHO_TOKEN_URL` / `ZOHO_API_BASE` (default to the Zoho **.in** data center — override
if your Zoho account is .com/.eu/etc.), `BIGBOSS_QUOTA_PROJECT`, `BIGBOSS_HOME`, `BIGBOSS_DOWNLOADS`.

Auto-resolved (no action): `HOME` = this folder; Google ADC path = your `%APPDATA%/gcloud`;
`DOWNLOADS` = your user Downloads folder.

Fill in `set_env.example.ps1` and run it in your PowerShell session before running Big Boss.

## One-time setup
1. Install Python 3.11+ and the libraries:
   `pip install google-cloud-bigquery google-ads pandas requests`
2. Authenticate to Google Cloud (gives BigQuery access):
   `gcloud auth application-default login`
   — sign in with an account that can read the BigQuery project in `bigboss_config.py`.
3. Confirm access works:  `python runq.py` against a one-line `SELECT 1 AS x` file.

## Prerequisites that are NOT in this bundle
- Access to the BigQuery dataset named in `bigboss_config.py` (`BQ_PROJECT.BQ_DATASET`), with the
  source tables already present (themes_firstlast_semroi, adg_keyword_universe, ads_* Data
  Transfer tables, etc.). This bundle reads/writes those tables; it does not create them.
- The target monthly tab must already exist in the Zoho sheet with BOTH columns
  ("AI Keyword Recommendations", "Change History") inserted. The writers only UPDATE.
- For Change History: drop the month's Google Ads change-history CSV (with the User column)
  into your Downloads folder first.
- For the PAA "PROBLEM QUERY" line: refresh `paa_queries.csv` (one row per product,theme,query)
  before running — this is harvested manually (People-Also-Ask / related searches per proven theme).
- **Different Google Ads account?** `account_keyword_status.sql` references the Data Transfer tables
  `ads_Keyword_5419501619`, `ads_Campaign_5419501619`, `ads_AdGroup_5419501619` (the number is the
  source account id). If you run BigQuery Data Transfer for a DIFFERENT Ads account, replace
  `5419501619` in that file with your own account id (same value as `GADS_CUSTOMER_ID`). No secrets
  are in any file — only this account-id suffix.

## Run it
All-in-one monthly orchestrator (run in July -> targets the "Jun 2026" tab):
```
python monthly_bigboss_all.py --rate <avg USD/INR for the change-month>
```
Useful flags: `--run-date YYYY-MM-DD`, `--tab "Jun 2026"`, `--skip-recs`, `--skip-changes`,
`--mode go|force|test`, `--plan` (dry-run, prints the resolved plan and exits).

Or run a column on its own:
- AI Keyword Recommendations: `python monthly_kw_recommendations.py` then `python sheet_write_recs.py force`
- Change History: `python monthly_change_history.py --start YYYY-MM-01 --end YYYY-MM-DD --rate <r> --tab "<Mon YYYY>"`

## Files
- `bigboss_config.py` — central config (edit this one if needed)
- `monthly_bigboss_all.py` — top-level orchestrator (both columns)
- `monthly_kw_recommendations.py` + the `*.sql` files + `build_*.py` / `kw_*.py` — AI column pipeline
- `build_paa_suggestions.py` + `paa_queries.csv` — PAA "PROBLEM QUERY" engine
- `monthly_change_history.py` + `build_change_events.py` + `build_change_history_cells.py` — Change History
- `sheet_write_recs.py` / `sheet_write_changes.py` — Zoho sheet writers
- `runq.py` — helper to run a .sql file and print CSV
