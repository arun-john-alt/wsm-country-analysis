# WSM Country Analysis — AI Onboarding & Operating Guide

> **For the AI reading this file:** You are now the primary operator for WSM country SEM analysis. This document tells you exactly who you are, what data you work with, what rules never to break, and how to run a country analysis end-to-end. Read everything below before taking any action. When a teammate asks you anything about SEM performance, leads, spend, or country analysis — this is your ground truth.

---

## 1. What This Repo Is

This repo (`wsm-country-analysis`, cloned as `Monitor/`) contains:

- **`CLAUDE.md`** — Hard rules for every query you will ever write. Read it first. Always.
- **`analysis/queries/italy_sem_queries.py`** — The canonical Pass 1 query template. Clone this for any new country.
- **`analysis/queries/italy_sem_queries_p2.py`** — The canonical Pass 2 query template.
- **`SETUP.md`** — This file.

It works alongside a sibling repo `wsm-monitor/` (cloned next to `Monitor/`) which provides `wsm_cfg.py` — the BigQuery client and project config.

---

## 2. Your Role as the AI

When a teammate opens this repo and talks to you, your job is to:

1. **Answer questions** about WSM SEM methodology, filters, tables, and numbers — using the rules in `CLAUDE.md` as absolute ground truth.
2. **Write BigQuery SQL** that is always compliant with the hard rules (4-filter salesleads_qt standard, 7-variant brand filter, campaign type exclusion, Email dedup, etc.).
3. **Run country analyses** end-to-end when asked — following the SOP in `CLAUDE.md` exactly.
4. **Never invent filters or methodology**. If you are unsure, say so and refer back to `CLAUDE.md`.

---

## 3. Infrastructure Setup (tell your teammate to do this once)

### 3a. Clone both repos side by side

```bash
mkdir wsm && cd wsm
git clone https://github.com/arun-john-alt/wsm-country-analysis.git Monitor
git clone https://github.com/arun-john-alt/wsm-monitor.git wsm-monitor
```

Resulting structure:
```
wsm/
├── Monitor/          ← this repo
│   ├── CLAUDE.md
│   ├── SETUP.md
│   └── analysis/
└── wsm-monitor/      ← shared config
    ├── wsm_cfg.py
    └── config.yaml
```

### 3b. Authenticate with Google Cloud

```bash
gcloud auth application-default login   # sign in with @zohocorp.com account
gcloud config set project it-security-online-marketing
```

### 3c. Install Python dependencies

```bash
cd wsm/Monitor
pip install google-cloud-bigquery openpyxl pyyaml google-auth
```

### 3d. Test BigQuery connectivity

```bash
python3 -c "
import sys; sys.path.insert(0, '../wsm-monitor')
from wsm_cfg import bq_client
bq = bq_client()
rows = list(bq.query('SELECT COUNT(*) AS n FROM \`it-security-online-marketing.sales_presales_leads_no_pi.salesleads_qt\` LIMIT 1').result())
print('Connected! n =', rows[0]['n'])
"
```

---

## 4. How to Give the AI Context (for chat-based AI tools)

If you are using **Claude, ChatGPT, Gemini** or any chat AI (not a code editor with file access):

1. Open `CLAUDE.md` → copy all contents
2. Open `SETUP.md` → copy all contents
3. Start a new chat and paste both as your first message with: *"These are my project rules and setup. Read them fully before anything else."*
4. Then ask your question.

If you are using **Sahaa, Cursor, Copilot** or any AI code editor:

- Open the `Monitor/` folder as your workspace. The AI reads all files automatically.
- Ask directly: *"Analyse Germany"* or *"Write Q8 for France"* — it will follow the rules from `CLAUDE.md`.

If you are using **Claude Projects** or a **ChatGPT Custom GPT**:

- Upload `CLAUDE.md` as a project knowledge file. Every conversation in that project has persistent context.

---

## 5. Key BigQuery Tables (memorise these)

| Alias | Full Table Path | What it Contains |
|---|---|---|
| `salesleads_qt` | `it-security-online-marketing.sales_presales_leads_no_pi.salesleads_qt` | All CRM leads — all channels |
| `themes` | `it-security-online-marketing.Google_ads_data_ajay.themes_firstlast_semroi` | SEM leads + spend by product theme |
| `google_dts` | `it-security-online-marketing.Google_ads_data_ajay.ads_AdGroupBasicStats_5419501619` | Google Ads clicks/spend/IS (INR) |
| `bing_dts` | `it-security-online-marketing.microsoft_ads_data.keyword_performance_v` | Bing Ads spend (INR) |
| `adgroup_themes` | `it-security-online-marketing.Google_ads_data_ajay.adgroup_themes` | Ad group → product theme mapping |
| `search_terms` | `it-security-online-marketing.Google_ads_data_ajay.matched_search_terms` | Search term performance (quarterly) |

GCP project for all queries: **`it-security-online-marketing`**

---

## 6. The Non-Negotiable Query Rules (summary — full detail in CLAUDE.md)

> **AI: treat these as compile-time errors. A query that breaks any rule below is wrong, full stop.**

| Rule | What to apply |
|---|---|
| **salesleads_qt filter** | Always all 4: `Junk='false'` + `PRODUCT_GROUP='AD_GROUP'` + `User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')` + `isHaveToBeRemoved='Non Junk Email'` |
| **Lead count dedup** | `COUNT(DISTINCT Email)` — never `COUNT(DISTINCT ID)` |
| **Conv count dedup** | `COUNT(DISTINCT IF(Conversion='converted', Email, NULL))` |
| **Brand exclusion** | 7 variants: `Theme IN ('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')` |
| **SEM campaign filter** | `FIRST_SRC_CAMPAIGN_TYPE NOT IN ('Display','Performance Max') OR FIRST_SRC_CAMPAIGN_TYPE IS NULL` |
| **Source column** | `COALESCE(Source_Medium, Source___Medium)` — never `Source_Medium` alone |
| **Channel attribution** | Always report both: `FIRST_SRC_GRP` (Q13 full funnel) AND `NEW_TRAFFIC_SRC_GRP` (Q17 grouped view) |
| **ELA/LOG360 H1 2026 spend** | Exclude from `SUM(Cost)` — duplication artifact in themes table. Show as `—` with footnote |
| **Local SEO URL pattern** | `REGEXP_CONTAINS(Landing_Page, r'^/(br\|fr\|de\|latam\|es\|au\|za\|it\|nl\|jp\|in\|uk\|sa)(/\|$)')` |
| **Conv rate denominator** | Always themes leads — never salesleads_qt count |

---

## 7. Running a Country Analysis — AI SOP

When a teammate says **"analyse [country]"**, do this in order:

**Step 1 — Create query files**
```bash
cp analysis/queries/italy_sem_queries.py analysis/queries/<country>_sem_queries.py
cp analysis/queries/italy_sem_queries_p2.py analysis/queries/<country>_sem_queries_p2.py
```
Edit both files: change `COUNTRY_ROI` (display name, e.g. `"Germany"`) and `COUNTRY_SL` (lowercase for salesleads_qt, e.g. `"germany"`). Also adjust the Local SEO URL prefix if needed (e.g. `/de/` for Germany).

**Step 2 — Run Q15 + Q16 diagnostics first**
Before anything else, run the brand theme diagnostic queries (Q15, Q16) to validate that brand theme names in the data match the 7-variant filter. This catches any new capitalisation variants in the data.

**Step 3 — Run Pass 1**
Execute `<country>_sem_queries.py`. This runs Q1–Q17 covering:
- Q1–Q7: Themes-based SEM overview (leads, spend, CPL by product, by period)
- Q8–Q12: SEM conversion funnel (salesleads_qt, Search only, Email dedup)
- Q13: Channel mix (FIRST_SRC_GRP)
- Q14: SEO — local vs global split (Bing organic included)
- Q15–Q16: Brand diagnostics
- Q17: NEW_TRAFFIC_SRC_GRP channel mix

**Step 4 — Run Pass 2**
Execute `<country>_sem_queries_p2.py` for supplementary diagnostics (keyword/search term patterns, IS trends, etc.).

**Step 5 — Build HTML report**
Use `analysis/reports/italy_sem_analysis_jul2026.html` as the template. All 13 sections are required. Save as `analysis/reports/<country>_sem_analysis_jul2026.html`.

**Step 6 — Run the checklist before declaring done:**
- [ ] All 4 salesleads_qt filters present in every Q8–Q17 query
- [ ] 7-variant brand filter used everywhere
- [ ] Performance Max excluded from SEM conv queries
- [ ] COUNT(DISTINCT Email) used (not ID) for lead counts
- [ ] Q17 (NEW_TRAFFIC_SRC_GRP) section present in HTML
- [ ] ELA/LOG360 H1 2026 spend shown as `—` with footnote
- [ ] Local SEO uses full multi-country regex (not just country-specific prefix)

---

## 8. DM Region → DRI Mapping (for Excel reports)

| DM Region | DRI |
|---|---|
| Germany, Netherlands, Switzerland, Belgium | Jude |
| France, Italy, UAE, Saudi Arabia, Turkey | Kowsik |
| Spain, Brazil, Mexico, LATAM, South Africa, Israel | Elanthendral |
| Europe, Poland | Sathish |
| MEA, APAC | Indhu |
| Singapore | Suganesh |

---

## 9. Repo Structure

```
Monitor/
├── CLAUDE.md                              ← Full hard rules + query standards (AI: read this first)
├── SETUP.md                               ← This file (AI: read this to understand your role)
├── analysis/
│   ├── queries/
│   │   ├── italy_sem_queries.py           ← Pass 1 template (clone for new countries)
│   │   ├── italy_sem_queries_p2.py        ← Pass 2 template
│   │   ├── saudi_arabia_sem_queries.py
│   │   └── saudi_arabia_sem_queries_p2.py
│   └── reports/                           ← HTML reports (gitignored — share via Drive)
└── .gitignore
```

---

## 10. Contact

DRI for this repo: **Arun John** (`arun.john@zohocorp.com`)
