# WSM Country Analysis — Teammate Setup Guide

> This repo contains all the query logic, hard rules, and institutional knowledge for WSM SEM & channel analysis.
> Once set up, an AI assistant (Sahaa) will automatically understand concepts like **Local SEO**, **Sales leads filters**, **brand theme exclusions**, etc. — because they are all documented in `CLAUDE.md`.

---

## Prerequisites

- Zoho Corp Google account with BigQuery access to `it-security-online-marketing`
- Python 3.10+ (or 3.9 minimum)
- [gcloud CLI](https://cloud.google.com/sdk/docs/install) installed
- [Sahaa](https://sahaa.ai) or another AI coding assistant in VSCode

---

## Step 1 — Clone both repos into the same parent folder

The analysis repo imports `wsm_cfg.py` from `wsm-monitor`, so both must sit side by side.

```bash
mkdir wsm && cd wsm
git clone https://github.com/arun-john-alt/wsm-country-analysis.git Monitor
git clone https://github.com/arun-john-alt/wsm-monitor.git wsm-monitor
```

Your folder structure should look like:
```
wsm/
├── Monitor/              ← this repo (country analysis)
│   ├── CLAUDE.md
│   ├── analysis/
│   └── SETUP.md
└── wsm-monitor/          ← shared config + monitor
    ├── wsm_cfg.py
    ├── config.yaml
    └── ...
```

---

## Step 2 — Authenticate with Google Cloud

```bash
gcloud auth application-default login
```

Sign in with your Zoho Corp Google account (`@zohocorp.com`). This creates Application Default Credentials that all Python scripts use automatically.

Verify it works:
```bash
gcloud config set project it-security-online-marketing
gcloud auth application-default print-access-token
# Should print a long token — if it does, you're authenticated
```

---

## Step 3 — Install Python dependencies

```bash
cd wsm/Monitor
pip install google-cloud-bigquery openpyxl pyyaml google-auth
```

For the monitor/alerting side (optional):
```bash
cd wsm/wsm-monitor
pip install -r requirements.txt
```

---

## Step 4 — Test a quick BQ query

```bash
cd wsm/Monitor
python3 -c "
import sys
sys.path.insert(0, '../wsm-monitor')
from wsm_cfg import PROJ, bq_client
bq = bq_client()
rows = list(bq.query('SELECT COUNT(*) AS n FROM \`it-security-online-marketing.sales_presales_leads_no_pi.salesleads_qt\` LIMIT 1').result())
print('Connected! Row count sample:', rows[0]['n'])
"
```

If you see `Connected!` — you're ready.

---

## Step 5 — Open in VSCode with Sahaa

1. Open the `Monitor` folder in VSCode: `code /path/to/wsm/Monitor`
2. Sahaa reads `CLAUDE.md` automatically as workspace context
3. You can now ask things like:
   - *"Pull Germany SEM leads for July 2026"*
   - *"What is the local SEO filter?"*
   - *"Analyse Italy"* — follows the full SOP from `CLAUDE.md`

---

## What the AI already knows (from CLAUDE.md)

| Concept | What's pre-defined |
|---|---|
| **Sales leads filter** | `Junk='false'`, `AD_GROUP`, `User_Type IN (new,adcs,mecs,inactive customer,inactive lead)`, `isHaveToBeRemoved='Non Junk Email'`, `COUNT(DISTINCT Email)` |
| **Local SEO** | URL regex `^/(br/fr/de/latam/es/au/za/it/nl/jp/in/uk)(/|$)` |
| **Brand filter** | 7 variants: `Branding`, `Log360 - Branding`, `Cloud Branding`, `cloud branding`, `ELA - Branding`, `AD360 - Branding`, `AD360 Branding` |
| **SEM campaign exclusion** | `FIRST_SRC_CAMPAIGN_TYPE NOT IN ('Display','Performance Max')` |
| **Channel columns** | `FIRST_SRC_GRP` for full funnel, `NEW_TRAFFIC_SRC_GRP` for grouped view |
| **ELA spend** | H1 2026 spend excluded from ELA/LOG360 (duplication) |
| **Country analysis SOP** | Steps 1–7 for running `analyse [country]` |

---

## Key BigQuery Tables

| Table | What it is |
|---|---|
| `sales_presales_leads_no_pi.salesleads_qt` | CRM sales leads (all channels) |
| `Google_ads_data_ajay.themes_firstlast_semroi` | SEM leads + spend by theme (Google + Bing) |
| `Google_ads_data_ajay.ads_AdGroupBasicStats_5419501619` | Google Ads DTS (clicks/spend/IS in INR) |
| `microsoft_ads_data.keyword_performance_v` | Bing Ads DTS (spend in INR) |
| `Google_ads_data_ajay.adgroup_themes` | Ad group → theme mapping |
| `Google_ads_data_ajay.matched_search_terms` | Search term performance (quarterly) |

All in project: **`it-security-online-marketing`**

---

## Repo Structure

```
Monitor/
├── CLAUDE.md                          ← All hard rules (read this first)
├── SETUP.md                           ← This file
├── analysis/
│   ├── queries/
│   │   ├── italy_sem_queries.py       ← Template for Pass 1 queries
│   │   ├── italy_sem_queries_p2.py    ← Template for Pass 2 queries
│   │   ├── saudi_arabia_sem_queries.py
│   │   └── saudi_arabia_sem_queries_p2.py
│   └── reports/                       ← Generated HTML reports (gitignored)
└── .gitignore
```

---

## Running a Country Analysis

```bash
# Clone Italy as template for a new country (e.g., Germany)
cp analysis/queries/italy_sem_queries.py analysis/queries/germany_sem_queries.py
cp analysis/queries/italy_sem_queries_p2.py analysis/queries/germany_sem_queries_p2.py

# Edit: change COUNTRY_ROI = "Germany" and COUNTRY_SL = "germany"
# Then run:
python3 analysis/queries/germany_sem_queries.py
```

Or just ask Sahaa: **"analyse Germany"** and it will follow the full SOP from `CLAUDE.md`.

---

## Contact

DRI for this repo: **Arun John** (`arun.john@zohocorp.com`)
wsm-monitor repo owner: check `config.yaml` DRI map
