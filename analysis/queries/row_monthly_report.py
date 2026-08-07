"""
ROW Monthly Report — DM Region × DRI Excel
==========================================
Generates the standard monthly YoY Excel for the ROW team.

Usage:
    python3 row_monthly_report.py                     # uses config.yaml month (current: 2026-07)
    python3 row_monthly_report.py --month 2026-07     # explicit month override
    python3 row_monthly_report.py --ytd               # include YTD tab (Jan–CUR vs prior year)
    python3 row_monthly_report.py --month 2026-07 --ytd  # both overrides

Output:
    Downloads/Monitor/<mon><year>_yoy.xlsx  (e.g. july2026_yoy.xlsx)

Structure:
    Tab 1 "DM Regions"      — CUR month vs same month prior year
    Tab 2 "YTD DM Regions"  — Jan–CUR vs prior year same span  (if --ytd)

Columns per tab:
    DM Region | DRI | [CUR] All Leads excl E/TP/O | All Convs |
               [PYR] All Leads excl E/TP/O | All Convs |
               [CUR] SEM Spend | SEM Leads | SEM Convs |
               [PYR] SEM Leads | SEM Convs

Data sources (BigQuery project: it-security-online-marketing):
    • salesleads_qt  → all-channel leads + convs (4-filter standard, Email dedup)
    • themes_firstlast_semroi → SEM spend + leads

Hard rules applied (CLAUDE.md):
    • salesleads_qt: Junk='false', PRODUCT_GROUP='AD_GROUP',
      User_Type IN ('new','adcs','mecs','inactive customer','inactive lead'),
      isHaveToBeRemoved='Non Junk Email', COUNT(DISTINCT Email)
    • Brand exclusion (7 variants — ALL must be excluded from SEM leads/spend)
    • SEM conv filter: FIRST_SRC_CAMPAIGN_TYPE NOT IN ('Display','Performance Max')
    • ELA/LOG360 H1 2026 spend excluded (duplication artifact)
"""

import sys, os, argparse
from datetime import date

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '../../wsm-monitor'))
import wsm_cfg as cfg
from wsm_cfg import bq_client, shift, label, mon, month_end

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Missing openpyxl. Run: pip install openpyxl")

# ── CLI args ──────────────────────────────────────────────────────────────────
ap = argparse.ArgumentParser()
ap.add_argument('--month', default=None, help='Target month YYYY-MM (default: from config.yaml)')
ap.add_argument('--ytd', action='store_true', help='Include YTD tab')
args = ap.parse_args()

CUR  = args.month or cfg.CUR
PYR  = shift(CUR, -12)
MON  = mon(CUR)
YEAR = CUR[:4]
PYEAR = PYR[:4]

# YTD: Jan through CUR month
ytd_months_cur = [f"{YEAR}-{m:02d}" for m in range(1, int(CUR[5:7]) + 1)]
ytd_months_pyr = [shift(ym, -12) for ym in ytd_months_cur]

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                   f"../../{MON.lower()}{YEAR}_yoy.xlsx")
OUT = os.path.normpath(OUT)

print(f"ROW Monthly Report — {label(CUR)} vs {label(PYR)}")
print(f"Output: {OUT}")
if args.ytd:
    print(f"YTD: Jan {YEAR} – {MON} {YEAR}  vs  Jan {PYEAR} – {mon(PYR)} {PYEAR}")

bq = bq_client()

# ── DM Region → DRI mapping ───────────────────────────────────────────────────
# Order matters — this is the display order in the Excel
DM_REGIONS = [
    ("United States",      "Aashiq"),
    ("India",              "Ajay"),
    ("United Kingdom",     "Ajay"),
    ("Canada",             "Ajay"),
    ("Australia",          "Ajay"),
    ("Germany",            "Jude"),
    ("Netherlands",        "Jude"),
    ("Switzerland",        "Jude"),
    ("Belgium",            "Jude"),
    ("France",             "Kowsik"),
    ("Italy",              "Kowsik"),
    ("United Arab Emirates","Kowsik"),
    ("Saudi Arabia",       "Kowsik"),
    ("Turkey",             "Kowsik"),
    ("Spain",              "Elanthendral"),
    ("Brazil",             "Elanthendral"),
    ("Mexico",             "Elanthendral"),
    ("Region - LATAM",     "Elanthendral"),
    ("South Africa",       "Elanthendral"),
    ("Israel",             "Elanthendral"),
    ("Region - Europe",    "Sathish"),
    ("Poland",             "Sathish"),
    ("Region - MEA",       "Indhu"),
    ("Region - APAC",      "Indhu"),
    ("Singapore",          "Suganesh"),
]

# Presales countries use FS_PS_Leads / Lead_Type='Mktg(SPL)Leads'
PRESALES = set(cfg.PRESALES_COUNTRIES)

# ── BQ helpers ────────────────────────────────────────────────────────────────
G   = cfg.G
SL  = f"{cfg.PROJ}.sales_presales_leads_no_pi.salesleads_qt"
ROI = f"{cfg.PROJ}.{G.split('.')[-1]}.themes_firstlast_semroi"

BRAND = "('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')"
ELA_GROUP = "('ELA','LOG360','LOG360CLOUD')"  # for spend exclusion in H1 2026

def ym_list(months):
    return ", ".join(f"'{m}'" for m in months)

def run(sql):
    return list(bq.query(sql).result())

def index_rows(rows, key_cols):
    """Turn BQ rows into dict keyed by tuple of key_col values."""
    d = {}
    for r in rows:
        k = tuple(r[c] for c in key_cols)
        d[k] = r
    return d

# ── Query 1: All-channel leads + convs from salesleads_qt ─────────────────────
# Excludes Email/Trial/Other lead types via User_Type filter (standard 4-filter)
# "All Leads excl E/TP/O" = standard salesleads_qt with 4 filters

def q_all_leads(months):
    mlist = ym_list(months)
    return f"""
SELECT
  COMMON_COUNTRY_NAME                                               AS country,
  COUNT(DISTINCT Email)                                             AS total_leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))           AS convs
FROM `{SL}`
WHERE SUBSTR(Created_Time,8,4) || '-' ||
      LPAD(CAST(CASE
        WHEN SUBSTR(Created_Time,4,3)='Jan' THEN 1
        WHEN SUBSTR(Created_Time,4,3)='Feb' THEN 2
        WHEN SUBSTR(Created_Time,4,3)='Mar' THEN 3
        WHEN SUBSTR(Created_Time,4,3)='Apr' THEN 4
        WHEN SUBSTR(Created_Time,4,3)='May' THEN 5
        WHEN SUBSTR(Created_Time,4,3)='Jun' THEN 6
        WHEN SUBSTR(Created_Time,4,3)='Jul' THEN 7
        WHEN SUBSTR(Created_Time,4,3)='Aug' THEN 8
        WHEN SUBSTR(Created_Time,4,3)='Sep' THEN 9
        WHEN SUBSTR(Created_Time,4,3)='Oct' THEN 10
        WHEN SUBSTR(Created_Time,4,3)='Nov' THEN 11
        WHEN SUBSTR(Created_Time,4,3)='Dec' THEN 12
      END AS STRING), 2, '0') IN ({mlist})
  AND Junk = 'false'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
GROUP BY 1
"""

# ── Query 2: SEM spend from themes (Non-Brand, excl ELA H1 2026 duplication) ──

def q_sem_spend(months):
    mlist = ym_list(months)
    # H1 2026 ELA/LOG360 spend exclusion: months in ['2026-01'..'2026-06']
    h1_months = [m for m in months if '2026-01' <= m <= '2026-06']
    ela_excl = ""
    if h1_months:
        ela_excl = f"AND NOT (Product IN {ELA_GROUP} AND SUBSTR(Date,1,7) IN ({ym_list(h1_months)}))"
    return f"""
SELECT
  CampaignCountry                                                   AS country,
  ROUND(SUM(Cost), 0)                                               AS spend_usd
FROM `{ROI}`
WHERE SUBSTR(Date,1,7) IN ({mlist})
  AND Theme NOT IN {BRAND}
  AND COALESCE(Source_Medium, Source___Medium) IN ('google / cpc','bing / cpc')
  {ela_excl}
GROUP BY 1
"""

# ── Query 3: SEM leads from themes (Non-Brand, all countries) ─────────────────

def q_sem_leads(months):
    mlist = ym_list(months)
    # Use FS_PS_Leads for presales, Valid_Sales_Leads_First_Source for others
    presales_list = ", ".join(f"'{c}'" for c in PRESALES)
    return f"""
SELECT
  CampaignCountry                                                   AS country,
  ROUND(SUM(
    IF(CampaignCountry IN ({presales_list}), FS_PS_Leads, Valid_Sales_Leads_First_Source)
  ))                                                                AS sem_leads
FROM `{ROI}`
WHERE SUBSTR(Date,1,7) IN ({mlist})
  AND Theme NOT IN {BRAND}
  AND COALESCE(Source_Medium, Source___Medium) IN ('google / cpc','bing / cpc')
  AND Lead_Type = IF(CampaignCountry IN ({presales_list}), 'Mktg(SPL)Leads', 'All Leads')
GROUP BY 1
"""

# ── Query 4: SEM convs from salesleads_qt (Non-Brand, Search only) ───────────

def q_sem_convs(months):
    mlist = ym_list(months)
    return f"""
SELECT
  COMMON_COUNTRY_NAME                                               AS country,
  COUNT(DISTINCT Email)                                             AS sem_leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))           AS sem_convs
FROM `{SL}`
WHERE SUBSTR(Created_Time,8,4) || '-' ||
      LPAD(CAST(CASE
        WHEN SUBSTR(Created_Time,4,3)='Jan' THEN 1
        WHEN SUBSTR(Created_Time,4,3)='Feb' THEN 2
        WHEN SUBSTR(Created_Time,4,3)='Mar' THEN 3
        WHEN SUBSTR(Created_Time,4,3)='Apr' THEN 4
        WHEN SUBSTR(Created_Time,4,3)='May' THEN 5
        WHEN SUBSTR(Created_Time,4,3)='Jun' THEN 6
        WHEN SUBSTR(Created_Time,4,3)='Jul' THEN 7
        WHEN SUBSTR(Created_Time,4,3)='Aug' THEN 8
        WHEN SUBSTR(Created_Time,4,3)='Sep' THEN 9
        WHEN SUBSTR(Created_Time,4,3)='Oct' THEN 10
        WHEN SUBSTR(Created_Time,4,3)='Nov' THEN 11
        WHEN SUBSTR(Created_Time,4,3)='Dec' THEN 12
      END AS STRING), 2, '0') IN ({mlist})
  AND FIRST_SRC_GRP IN ('google / cpc','bing / cpc')
  AND FIRST_SRC_THEME NOT IN {BRAND}
  AND (FIRST_SRC_CAMPAIGN_TYPE NOT IN ('Display','Performance Max') OR FIRST_SRC_CAMPAIGN_TYPE IS NULL)
  AND Junk = 'false'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
GROUP BY 1
"""

# ── Fetch data ─────────────────────────────────────────────────────────────────
print("\nFetching data from BigQuery...")

# Current month
print(f"  [1/8] All-channel leads {label(CUR)}...")
cur_leads  = index_rows(run(q_all_leads([CUR])),  ['country'])
print(f"  [2/8] SEM spend {label(CUR)}...")
cur_spend  = index_rows(run(q_sem_spend([CUR])),  ['country'])
print(f"  [3/8] SEM leads (themes) {label(CUR)}...")
cur_sl     = index_rows(run(q_sem_leads([CUR])),  ['country'])
print(f"  [4/8] SEM convs {label(CUR)}...")
cur_convs  = index_rows(run(q_sem_convs([CUR])),  ['country'])

# Prior year same month
print(f"  [5/8] All-channel leads {label(PYR)}...")
pyr_leads  = index_rows(run(q_all_leads([PYR])),  ['country'])
print(f"  [6/8] SEM leads (themes) {label(PYR)}...")
pyr_sl     = index_rows(run(q_sem_leads([PYR])),  ['country'])
print(f"  [7/8] SEM spend {label(PYR)}...")
pyr_spend  = index_rows(run(q_sem_spend([PYR])),  ['country'])
print(f"  [8/8] SEM convs {label(PYR)}...")
pyr_convs  = index_rows(run(q_sem_convs([PYR])),  ['country'])

# YTD data (optional)
ytd_cur_leads = ytd_cur_spend = ytd_cur_sl = ytd_cur_convs = {}
ytd_pyr_leads = ytd_pyr_spend = ytd_pyr_sl = ytd_pyr_convs = {}
if args.ytd:
    print(f"\nFetching YTD data (Jan–{MON} {YEAR} vs Jan–{mon(PYR)} {PYEAR})...")
    ytd_cur_leads = index_rows(run(q_all_leads(ytd_months_cur)),  ['country'])
    ytd_cur_spend = index_rows(run(q_sem_spend(ytd_months_cur)), ['country'])
    ytd_cur_sl    = index_rows(run(q_sem_leads(ytd_months_cur)), ['country'])
    ytd_cur_convs = index_rows(run(q_sem_convs(ytd_months_cur)), ['country'])
    ytd_pyr_leads = index_rows(run(q_all_leads(ytd_months_pyr)),  ['country'])
    ytd_pyr_spend = index_rows(run(q_sem_spend(ytd_months_pyr)), ['country'])
    ytd_pyr_sl    = index_rows(run(q_sem_leads(ytd_months_pyr)), ['country'])
    ytd_pyr_convs = index_rows(run(q_sem_convs(ytd_months_pyr)), ['country'])

print("Done fetching.\n")

# ── Excel helpers ──────────────────────────────────────────────────────────────
def v(d, country, col, default=0):
    """Safe value lookup."""
    k = (country,)
    if k not in d: return default
    val = d[k][col]
    return int(val) if val is not None else default

def pct_str(cur, pyr):
    """YoY % string: +12% or -5%"""
    if not pyr: return "—"
    p = (cur - pyr) / pyr * 100
    return f"{'+' if p >= 0 else ''}{p:.0f}%"

# Styles
YELLOW  = PatternFill("solid", fgColor="FFD700")
PINK    = PatternFill("solid", fgColor="FFB6C1")
LBLUE   = PatternFill("solid", fgColor="ADD8E6")
GREEN_H = PatternFill("solid", fgColor="92D050")  # header green
GREEN_F = PatternFill("solid", fgColor="C6EFCE")  # data green (≥+10%)
RED_F   = PatternFill("solid", fgColor="FFC7CE")  # data red (≤-10%)
GREY    = PatternFill("solid", fgColor="F2F2F2")

BOLD    = Font(bold=True)
BOLD_SM = Font(bold=True, size=9)
SM      = Font(size=9)

def hdr(ws, row, col, val, fill=None, bold=True, align='center', wrap=False):
    c = ws.cell(row=row, column=col, value=val)
    if fill: c.fill = fill
    c.font = Font(bold=bold)
    c.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    return c

def cell(ws, row, col, val, fill=None, align='right', fmt=None, color_rule=None):
    c = ws.cell(row=row, column=col, value=val)
    c.font = SM
    c.alignment = Alignment(horizontal=align)
    if fill: c.fill = fill
    if fmt:  c.number_format = fmt
    if color_rule is not None and isinstance(val, (int, float)):
        if color_rule >= 10:   c.fill = GREEN_F
        elif color_rule <= -10: c.fill = RED_F
    return c

# ── Build a sheet ──────────────────────────────────────────────────────────────
def build_sheet(ws, period_label_cur, period_label_pyr,
                leads_cur, spend_cur, sl_cur, convs_cur,
                leads_pyr, spend_pyr, sl_pyr, convs_pyr):

    # Row 1 — period header
    ws.merge_cells('A1:B1'); hdr(ws, 1, 1, "DM Region / DRI", fill=YELLOW)
    ws.merge_cells('C1:D1'); hdr(ws, 1, 3, period_label_cur, fill=YELLOW)
    ws.merge_cells('E1:F1'); hdr(ws, 1, 5, period_label_pyr, fill=YELLOW)
    ws.merge_cells('G1:I1'); hdr(ws, 1, 7, f"SEM — {period_label_cur}", fill=GREEN_H)
    ws.merge_cells('J1:L1'); hdr(ws, 1, 10, f"SEM — {period_label_pyr}", fill=GREEN_H)

    # Row 2 — column headers
    headers = [
        (1, "DM Region", GREY),
        (2, "DRI", GREY),
        (3, "All Leads\nexcl E/TP/O", PINK),
        (4, "All Convs", PINK),
        (5, "All Leads\nexcl E/TP/O", LBLUE),
        (6, "All Convs", LBLUE),
        (7, "Spend (USD)", GREEN_H),
        (8, "SEM Leads\n(themes)", GREEN_H),
        (9, "SEM Convs\n(salesleads)", GREEN_H),
        (10, "SEM Leads\n(themes)", GREEN_H),
        (11, "SEM Convs\n(salesleads)", GREEN_H),
        (12, "Leads YoY%", GREY),
    ]
    for col, title, fill in headers:
        c = ws.cell(row=2, column=col, value=title)
        c.fill = fill; c.font = Font(bold=True, size=9)
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    # Data rows
    row = 3
    for country, dri in DM_REGIONS:
        cl  = v(leads_cur, country, 'total_leads')
        cc  = v(leads_cur, country, 'convs')
        pl  = v(leads_pyr, country, 'total_leads')
        pc  = v(leads_pyr, country, 'convs')
        cs  = v(spend_cur, country, 'spend_usd')
        csl = v(sl_cur,    country, 'sem_leads')
        csc = v(convs_cur, country, 'sem_leads')   # SEM leads from salesleads_qt
        ccc = v(convs_cur, country, 'sem_convs')
        psl = v(sl_pyr,    country, 'sem_leads')
        pcc = v(convs_pyr, country, 'sem_leads')
        pccc= v(convs_pyr, country, 'sem_convs')

        # YoY % for color rules
        yoy_pct = ((cl - pl) / pl * 100) if pl else None

        fill_row = GREY if row % 2 == 0 else None
        cell(ws, row, 1, country,  fill=fill_row, align='left')
        cell(ws, row, 2, dri,      fill=fill_row, align='left')
        cell(ws, row, 3, cl or None, fill=PINK, fmt='#,##0', color_rule=yoy_pct)
        cell(ws, row, 4, cc or None, fill=PINK, fmt='#,##0')
        cell(ws, row, 5, pl or None, fill=LBLUE, fmt='#,##0')
        cell(ws, row, 6, pc or None, fill=LBLUE, fmt='#,##0')
        cell(ws, row, 7, cs or None, fmt='$#,##0')
        cell(ws, row, 8, csl or None, fmt='#,##0', color_rule=yoy_pct)
        cell(ws, row, 9, ccc or None, fmt='#,##0')
        cell(ws, row, 10, psl or None, fmt='#,##0')
        cell(ws, row, 11, pccc or None, fmt='#,##0')
        cell(ws, row, 12, pct_str(cl, pl) if pl else "—", align='center')
        row += 1

    # Column widths
    widths = [22, 14, 14, 10, 14, 10, 12, 14, 14, 14, 14, 10]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 30
    ws.freeze_panes = 'C3'

# ── Build workbook ─────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()

# Tab 1 — Monthly
ws1 = wb.active
ws1.title = "DM Regions"
build_sheet(
    ws1,
    f"{MON} {YEAR}", f"{MON} {PYEAR}",
    cur_leads, cur_spend, cur_sl, cur_convs,
    pyr_leads, pyr_spend, pyr_sl, pyr_convs,
)
print(f"✅ Tab 'DM Regions' built ({label(CUR)} vs {label(PYR)})")

# Tab 2 — YTD (optional)
if args.ytd:
    ws2 = wb.create_sheet("YTD DM Regions")
    ytd_label_cur = f"Jan–{MON} {YEAR}"
    ytd_label_pyr = f"Jan–{mon(PYR)} {PYEAR}"
    build_sheet(
        ws2,
        ytd_label_cur, ytd_label_pyr,
        ytd_cur_leads, ytd_cur_spend, ytd_cur_sl, ytd_cur_convs,
        ytd_pyr_leads, ytd_pyr_spend, ytd_pyr_sl, ytd_pyr_convs,
    )
    print(f"✅ Tab 'YTD DM Regions' built ({ytd_label_cur} vs {ytd_label_pyr})")

wb.save(OUT)
print(f"\n📊 Saved: {OUT}")
print("   Green cells = ≥+10% YoY leads | Red cells = ≤-10% YoY leads")
print("\nNote: ELA/LOG360/LOG360CLOUD spend excluded for H1 2026 months (duplication artifact).")
print("      SEM Leads column uses themes table (Valid_Sales_Leads / FS_PS_Leads by market).")
print("      SEM Convs column uses salesleads_qt (Email dedup, Search only, no PMax/Display).")
