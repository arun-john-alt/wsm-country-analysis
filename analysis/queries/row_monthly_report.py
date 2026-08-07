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

Format (matches the canonical july2026_yoy.xlsx):
    Row 1: Big yellow title spanning all columns e.g. "Jul 2026 (vs Jul '25)"
    Row 2: Section group headers (pink / light-blue / green)
    Row 3: Column headers
    Row 4+: Data rows — YoY% is INLINE in each cell (e.g. "760 (+9%)")
             green fill >= +10%, pink fill <= -10%, no fill otherwise
    ROW markets only (US/India/UK/CA/AU excluded — those are presales, not ROW)

    3 section groups:
      [A] All leads except Events, Third Party & Others  — salesleads_qt 4-filter
          → All Leads (Created date) | All Conversions (Conv. Date)
      [B] All leads  — salesleads_qt WITHOUT User_Type / isHaveToBeRemoved filters
          → All Leads (Created date) | All Conversions (Conv. Date)
      [C] SEM leads  — themes table (Non-Brand, Google+Bing)
          → SEM Spending (Google USD) | SEM Leads (Created date) | SEM Conversions (Conv. Date)

Hard rules applied (CLAUDE.md):
    Section A salesleads_qt: Junk='false', PRODUCT_GROUP='AD_GROUP',
      User_Type IN ('new','adcs','mecs','inactive customer','inactive lead'),
      isHaveToBeRemoved='Non Junk Email', COUNT(DISTINCT Email)
    Brand exclusion (7 variants)
    SEM conv filter: FIRST_SRC_CAMPAIGN_TYPE NOT IN ('Display','Performance Max')
    ELA/LOG360 H1 2026 spend excluded (duplication artifact)
"""

import sys, os, argparse, re

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '../../wsm-monitor'))
import wsm_cfg as cfg
from wsm_cfg import bq_client, shift, label, mon, month_end

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.rich_text import InlineFont, CellRichText, TextBlock
except ImportError:
    sys.exit("Missing openpyxl. Run: pip install openpyxl")

# ── CLI args ──────────────────────────────────────────────────────────────────
ap = argparse.ArgumentParser()
ap.add_argument('--month', default=None, help='Target month YYYY-MM (default: from config.yaml)')
ap.add_argument('--ytd',   action='store_true', help='Include YTD tab')
args = ap.parse_args()

CUR   = args.month or cfg.CUR
PYR   = shift(CUR, -12)
MON   = mon(CUR)
YEAR  = CUR[:4]
PYEAR = PYR[:4]
PYR_SHORT = PYEAR[2:]  # '25' from '2025'

# YTD: Jan through CUR month
ytd_months_cur = [f"{YEAR}-{m:02d}" for m in range(1, int(CUR[5:7]) + 1)]
ytd_months_pyr = [shift(ym, -12) for ym in ytd_months_cur]

OUT = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    f"../../{MON.lower()}{YEAR}_yoy.xlsx"
))

print(f"ROW Monthly Report — {label(CUR)} vs {label(PYR)}")
print(f"Output: {OUT}")
if args.ytd:
    print(f"YTD: Jan {YEAR} – {MON} {YEAR}  vs  Jan {PYEAR} – {mon(PYR)} {PYEAR}")

bq = bq_client()

# ── ROW DM Region list (ROW only — no US/IN/UK/CA/AU) ────────────────────────
# Names must match CampaignCountry in themes table AND COMMON_COUNTRY_NAME in salesleads_qt
# Display name → (themes_country, salesleads_country)
DM_REGIONS = [
    ("Germany",          "Jude",         "Germany",             "germany"),
    ("Netherlands",      "Jude",         "Netherlands",         "netherlands"),
    ("Switzerland",      "Jude",         "Switzerland",         "switzerland"),
    ("Belgium",          "Jude",         "Belgium",             "belgium"),
    ("France",           "Kowsik",       "France",              "france"),
    ("Italy",            "Kowsik",       "Italy",               "italy"),
    ("United Arab Emirates","Kowsik",    "United Arab Emirates","united arab emirates"),
    ("Saudi Arabia",     "Kowsik",       "Saudi Arabia",        "saudi arabia"),
    ("Turkey",           "Kowsik",       "Turkey",              "turkey"),
    ("Spain",            "Elanthendral", "Spain",               "spain"),
    ("Brazil",           "Elanthendral", "Brazil",              "brazil"),
    ("Mexico",           "Elanthendral", "Mexico",              "mexico"),
    ("Rest Of LATAM",    "Elanthendral", "Region - LATAM",      "rest of latam"),
    ("South Africa",     "Elanthendral", "South Africa",        "south africa"),
    ("Israel",           "Elanthendral", "Israel",              "israel"),
    ("Rest Of Europe",   "Sathish",      "Region - Europe",     "rest of europe"),
    ("Poland",           "Sathish",      "Poland",              "poland"),
    ("Rest Of MEA",      "Indhu",        "Region - MEA",        "rest of mea"),
    ("Rest Of APAC",     "Indhu",        "Region - APAC",       "rest of apac"),
    ("Singapore",        "Suganesh",     "Singapore",           "singapore"),
]

# ── BQ constants ──────────────────────────────────────────────────────────────
G   = cfg.G
SL  = f"{cfg.PROJ}.sales_presales_leads_no_pi.salesleads_qt"
ROI = f"{cfg.PROJ}.{G.split('.')[-1]}.themes_firstlast_semroi"

BRAND     = "('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')"
ELA_GROUP = "('ELA','LOG360','LOG360CLOUD')"

def ym_list(months):
    return ", ".join(f"'{m}'" for m in months)

def run(sql):
    return list(bq.query(sql).result())

def idx(rows, key):
    """Index rows by a single key column."""
    return {r[key]: r for r in rows}

# ── Date filter helper for salesleads_qt ─────────────────────────────────────
# Created_Time format: "01 Jan 2026 12:00:00" — extract YYYY-MM
SL_DATE = """FORMAT_DATE('%Y-%m', SAFE.PARSE_DATE('%d %b %Y', SUBSTR(Created_Time,1,11)))"""

def sl_date_filter(months):
    return f"{SL_DATE} IN ({ym_list(months)})"

# ── Query A: All leads excl E/TP/O (salesleads_qt 4-filter standard) ─────────
def q_leads_excl(months, salesleads_key='COMMON_COUNTRY_NAME'):
    return f"""
SELECT
  {salesleads_key}                                                  AS country,
  COUNT(DISTINCT Email)                                             AS leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))           AS convs
FROM `{SL}`
WHERE {sl_date_filter(months)}
  AND Junk = 'false'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
GROUP BY 1
"""

# ── Query B: All leads (no User_Type / isHaveToBeRemoved filter) ──────────────
def q_leads_all(months, salesleads_key='COMMON_COUNTRY_NAME'):
    return f"""
SELECT
  {salesleads_key}                                                  AS country,
  COUNT(DISTINCT Email)                                             AS leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))           AS convs
FROM `{SL}`
WHERE {sl_date_filter(months)}
  AND Junk = 'false'
  AND PRODUCT_GROUP = 'AD_GROUP'
GROUP BY 1
"""

# ── Query C: SEM spend (themes, Non-Brand) ────────────────────────────────────
def q_sem_spend(months):
    h1 = [m for m in months if '2026-01' <= m <= '2026-06']
    ela_excl = (f"AND NOT (Product IN {ELA_GROUP} AND SUBSTR(Date,1,7) IN ({ym_list(h1)}))"
                if h1 else "")
    return f"""
SELECT
  CampaignCountry                                                   AS country,
  ROUND(SUM(Cost), 0)                                               AS spend_usd
FROM `{ROI}`
WHERE SUBSTR(Date,1,7) IN ({ym_list(months)})
  AND Theme NOT IN {BRAND}
  AND COALESCE(Source_Medium, Source___Medium) IN ('google / cpc','bing / cpc')
  {ela_excl}
GROUP BY 1
"""

# ── Query C: SEM leads (themes, Non-Brand) ────────────────────────────────────
def q_sem_leads_themes(months):
    # ROW markets all use Valid_Sales_Leads_First_Source / Lead_Type='All Leads'
    return f"""
SELECT
  CampaignCountry                                                   AS country,
  ROUND(SUM(Valid_Sales_Leads_First_Source))                        AS sem_leads
FROM `{ROI}`
WHERE SUBSTR(Date,1,7) IN ({ym_list(months)})
  AND Theme NOT IN {BRAND}
  AND COALESCE(Source_Medium, Source___Medium) IN ('google / cpc','bing / cpc')
  AND Lead_Type = 'All Leads'
GROUP BY 1
"""

# ── Query C: SEM convs (salesleads_qt, Non-Brand, Search only) ───────────────
def q_sem_convs(months):
    return f"""
SELECT
  LOWER(COMMON_COUNTRY_NAME)                                        AS country_lc,
  COUNT(DISTINCT Email)                                             AS sem_leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))           AS sem_convs
FROM `{SL}`
WHERE {sl_date_filter(months)}
  AND FIRST_SRC_GRP IN ('google / cpc','bing / cpc')
  AND FIRST_SRC_THEME NOT IN {BRAND}
  AND (FIRST_SRC_CAMPAIGN_TYPE NOT IN ('Display','Performance Max') OR FIRST_SRC_CAMPAIGN_TYPE IS NULL)
  AND Junk = 'false'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
GROUP BY 1
"""

# ── Fetch all data ─────────────────────────────────────────────────────────────
def fetch(months, label_str):
    print(f"  Fetching {label_str}...")
    a = idx(run(q_leads_excl(months)),      'country')
    b = idx(run(q_leads_all(months)),       'country')
    c_spend = idx(run(q_sem_spend(months)), 'country')
    c_leads = idx(run(q_sem_leads_themes(months)), 'country')
    c_convs = idx(run(q_sem_convs(months)), 'country_lc')
    return a, b, c_spend, c_leads, c_convs

print("\nFetching data from BigQuery...")
cur_a, cur_b, cur_spend, cur_csl, cur_csc = fetch([CUR], label(CUR))
pyr_a, pyr_b, pyr_spend, pyr_csl, pyr_csc = fetch([PYR], label(PYR))

if args.ytd:
    ytd_cur_a, ytd_cur_b, ytd_cur_spend, ytd_cur_csl, ytd_cur_csc = fetch(ytd_months_cur, f"YTD {YEAR}")
    ytd_pyr_a, ytd_pyr_b, ytd_pyr_spend, ytd_pyr_csl, ytd_pyr_csc = fetch(ytd_months_pyr, f"YTD {PYEAR}")

print("Done.\n")

# ── Value helpers ─────────────────────────────────────────────────────────────
def gv(d, key, col, default=0):
    r = d.get(key)
    if r is None: return default
    v = r[col]
    return int(v) if v is not None else default

def yoy_pct(cur, pyr):
    """Returns (float pct or None, str '+12%' or '—')"""
    if not pyr: return None, "—"
    p = (cur - pyr) / pyr * 100
    return p, f"{'+' if p >= 0 else ''}{p:.0f}%"

# ── Excel styles ──────────────────────────────────────────────────────────────
YELLOW   = PatternFill("solid", fgColor="FFD700")
PINK_H   = PatternFill("solid", fgColor="F4CCCC")   # section header pink
LBLUE_H  = PatternFill("solid", fgColor="CFE2F3")   # section header light blue
GREEN_H  = PatternFill("solid", fgColor="B6D7A8")   # section header green
GREEN_D  = PatternFill("solid", fgColor="D9EAD3")   # data cell green (good YoY)
PINK_D   = PatternFill("solid", fgColor="FCE8E6")   # data cell pink (bad YoY)
COL_HDR  = PatternFill("solid", fgColor="EFEFEF")   # column header grey

def af(bold=False, size=10, color="000000", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic)

def ac(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def set_cell(ws, row, col, value=None, fill=None, font=None, align=None):
    c = ws.cell(row=row, column=col, value=value)
    if fill:  c.fill  = fill
    if font:  c.font  = font
    if align: c.alignment = align
    return c

def rich_cell(ws, row, col, number, yoy_str, fill=None):
    """Write 'number' bold + ' (yoy_str)' in grey, inline in same cell."""
    c = ws.cell(row=row, column=col)
    if number == 0:
        c.value = "—"
        c.font  = af(size=9)
        c.alignment = ac('center')
        if fill: c.fill = fill
        return c
    num_str  = f"{number:,}"
    yoy_part = f" ({yoy_str})"
    c.value = CellRichText(
        TextBlock(InlineFont(b=True,  sz=1800), num_str),
        TextBlock(InlineFont(b=False, sz=1600, color="999999"), yoy_part),
    )
    c.alignment = ac('center')
    if fill: c.fill = fill
    return c

def dollar_rich(ws, row, col, number, yoy_str, fill=None):
    """Same as rich_cell but formats number as $xxx,xxx."""
    c = ws.cell(row=row, column=col)
    if number == 0:
        c.value = "—"
        c.font  = af(size=9)
        c.alignment = ac('center')
        if fill: c.fill = fill
        return c
    num_str  = f"${number:,.0f}"
    yoy_part = f" ({yoy_str})"
    c.value = CellRichText(
        TextBlock(InlineFont(b=True,  sz=1800), num_str),
        TextBlock(InlineFont(b=False, sz=1600, color="999999"), yoy_part),
    )
    c.alignment = ac('center')
    if fill: c.fill = fill
    return c

# ── Build a single worksheet ──────────────────────────────────────────────────
def build_sheet(ws, title,
                cur_a, cur_b, cur_spend, cur_csl, cur_csc,
                pyr_a, pyr_b, pyr_spend, pyr_csl, pyr_csc):

    # ── Row 1: big title ──────────────────────────────────────────────────────
    ws.merge_cells('A1:I1')
    t = ws.cell(row=1, column=1, value=title)
    t.fill = YELLOW
    t.font = af(bold=True, size=14)
    t.alignment = ac('center')
    ws.row_dimensions[1].height = 28

    # ── Row 2: section group headers ─────────────────────────────────────────
    ws.merge_cells('A2:B2')  # blank spacer
    ws.cell(row=2, column=1).fill = YELLOW

    ws.merge_cells('C2:D2')
    s1 = ws.cell(row=2, column=3, value="All leads except Events, Third Party & Others")
    s1.fill = PINK_H; s1.font = af(bold=True, size=10); s1.alignment = ac('center')

    ws.merge_cells('E2:F2')
    s2 = ws.cell(row=2, column=5, value="All leads")
    s2.fill = LBLUE_H; s2.font = af(bold=True, size=10); s2.alignment = ac('center')

    ws.merge_cells('G2:I2')
    s3 = ws.cell(row=2, column=7, value="SEM leads")
    s3.fill = GREEN_H; s3.font = af(bold=True, size=10); s3.alignment = ac('center')
    ws.row_dimensions[2].height = 22

    # ── Row 3: column headers ─────────────────────────────────────────────────
    COL_DEFS = [
        (1, "DM Region",                    COL_HDR),
        (2, "DRI",                          COL_HDR),
        (3, "All Leads\n(Created date)",    PINK_H),
        (4, "All Conversions\n(Conv. Date)",PINK_H),
        (5, "All Leads\n(Created date)",    LBLUE_H),
        (6, "All Conversions\n(Conv. Date)",LBLUE_H),
        (7, "SEM Spending\n(Google USD)",   GREEN_H),
        (8, "SEM Leads\n(Created date)",    GREEN_H),
        (9, "SEM Conversions\n(Conv. Date)",GREEN_H),
    ]
    for col, title_h, fill in COL_DEFS:
        c = ws.cell(row=3, column=col, value=title_h)
        c.fill = fill
        c.font = af(bold=True, size=9)
        c.alignment = ac('center', wrap=True)
    ws.row_dimensions[3].height = 32

    # ── Data rows ─────────────────────────────────────────────────────────────
    for i, (display, dri, themes_key, sl_key) in enumerate(DM_REGIONS):
        row = i + 4

        # Section A — excl E/TP/O
        al_c = gv(cur_a, display, 'leads')
        al_p = gv(pyr_a, display, 'leads')
        ac_c = gv(cur_a, display, 'convs')
        ac_p = gv(pyr_a, display, 'convs')

        # Section B — all leads
        bl_c = gv(cur_b, display, 'leads')
        bl_p = gv(pyr_b, display, 'leads')
        bc_c = gv(cur_b, display, 'convs')
        bc_p = gv(pyr_b, display, 'convs')

        # Section C — SEM
        sp_c = gv(cur_spend, themes_key, 'spend_usd')
        sp_p = gv(pyr_spend, themes_key, 'spend_usd')
        sl_c = gv(cur_csl,   themes_key, 'sem_leads')
        sl_p = gv(pyr_csl,   themes_key, 'sem_leads')
        sc_c = gv(cur_csc,   sl_key,     'sem_convs')
        sc_p = gv(pyr_csc,   sl_key,     'sem_convs')

        # YoY pcts
        al_pct, al_str = yoy_pct(al_c, al_p)
        ac_pct, ac_str = yoy_pct(ac_c, ac_p)
        bl_pct, bl_str = yoy_pct(bl_c, bl_p)
        bc_pct, bc_str = yoy_pct(bc_c, bc_p)
        sp_pct, sp_str = yoy_pct(sp_c, sp_p)
        sl_pct, sl_str = yoy_pct(sl_c, sl_p)
        sc_pct, sc_str = yoy_pct(sc_c, sc_p)

        def fill_for(pct, section):
            if pct is None: return section
            if pct >= 10: return GREEN_D
            if pct <= -10: return PINK_D
            return section  # no change = same as section colour

        # Col A — DM Region (bold)
        c = ws.cell(row=row, column=1, value=display)
        c.font = af(bold=True, size=10); c.alignment = ac('left', 'center')

        # Col B — DRI
        c = ws.cell(row=row, column=2, value=dri)
        c.font = af(size=10); c.alignment = ac('left', 'center')

        # Cols C–D — Section A (excl E/TP/O)
        rich_cell(ws, row, 3, al_c, al_str, fill=fill_for(al_pct, PINK_H))
        rich_cell(ws, row, 4, ac_c, ac_str, fill=fill_for(ac_pct, PINK_H))

        # Cols E–F — Section B (all leads)
        rich_cell(ws, row, 5, bl_c, bl_str, fill=fill_for(bl_pct, LBLUE_H))
        rich_cell(ws, row, 6, bc_c, bc_str, fill=fill_for(bc_pct, LBLUE_H))

        # Cols G–I — Section C (SEM)
        dollar_rich(ws, row, 7, sp_c, sp_str, fill=fill_for(sp_pct, GREEN_H))
        rich_cell(ws,  row, 8, sl_c, sl_str, fill=fill_for(sl_pct, GREEN_H))
        rich_cell(ws,  row, 9, sc_c, sc_str, fill=fill_for(sc_pct, GREEN_H))

        ws.row_dimensions[row].height = 20

    # ── Column widths ─────────────────────────────────────────────────────────
    widths = [22, 14, 16, 16, 16, 16, 16, 16, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = 'C4'

# ── Build workbook ─────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()

# Tab 1 — Monthly
ws1 = wb.active
ws1.title = "DM Regions"
month_title = f"{MON} {YEAR}  (vs {MON} '{PYR_SHORT})"
build_sheet(ws1, month_title,
            cur_a, cur_b, cur_spend, cur_csl, cur_csc,
            pyr_a, pyr_b, pyr_spend, pyr_csl, pyr_csc)
print(f"✅ Tab 'DM Regions' — {month_title}")

# Tab 2 — YTD
if args.ytd:
    ws2 = wb.create_sheet("YTD DM Regions")
    ytd_title = f"Jan–{MON} {YEAR}  (vs Jan–{mon(PYR)} '{PYR_SHORT})"
    build_sheet(ws2, ytd_title,
                ytd_cur_a, ytd_cur_b, ytd_cur_spend, ytd_cur_csl, ytd_cur_csc,
                ytd_pyr_a, ytd_pyr_b, ytd_pyr_spend, ytd_pyr_csl, ytd_pyr_csc)
    print(f"✅ Tab 'YTD DM Regions' — {ytd_title}")

wb.save(OUT)
print(f"\n📊 Saved: {OUT}")
print("   Green cell = ≥+10% YoY  |  Pink cell = ≤-10% YoY  |  Inline grey = YoY%")
print("   Section A (pink)  = salesleads_qt 4-filter (excl existing customers/leads/junk email)")
print("   Section B (blue)  = salesleads_qt Junk+AD_GROUP only (all leads)")
print("   Section C (green) = themes table SEM Non-Brand (Google+Bing, excl PMax/Display)")
print("   ELA/LOG360 H1 2026 spend excluded (duplication artifact in themes table)")
