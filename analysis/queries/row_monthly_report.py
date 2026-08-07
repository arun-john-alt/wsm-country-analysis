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
    RICH_TEXT = True
except ImportError:
    try:
        import openpyxl
        from openpyxl.styles import PatternFill, Font, Alignment
        from openpyxl.utils import get_column_letter
        RICH_TEXT = False
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
# Tuple: (display_name, dri, themes_CampaignCountry_key, sl_group_key)
# For individual countries: sl_group_key = COMMON_COUNTRY_NAME value
# For regional buckets: sl_group_key = display_name, mapped via SL_REGION_MAP below
DM_REGIONS = [
    ("Germany",               "Jude",         "Germany",           "Germany"),
    ("Netherlands",           "Jude",         "Netherlands",       "Netherlands"),
    ("Switzerland",           "Jude",         "Switzerland",       "Switzerland"),
    ("Belgium",               "Jude",         "Belgium",           "Belgium"),
    ("France",                "Kowsik",       "France",            "France"),
    ("Italy",                 "Kowsik",       "Italy",             "Italy"),
    ("United Arab Emirates",  "Kowsik",       "United Arab Emirates","United Arab Emirates"),
    ("Saudi Arabia",          "Kowsik",       "Saudi Arabia",      "Saudi Arabia"),
    ("Turkey",                "Kowsik",       "Turkey",            "Turkey"),
    ("Spain",                 "Elanthendral", "Spain",             "Spain"),
    ("Brazil",                "Elanthendral", "Brazil",            "Brazil"),
    ("Mexico",                "Elanthendral", "Mexico",            "Mexico"),
    ("Rest Of LATAM",         "Elanthendral", "Region - LATAM",    "Rest Of LATAM"),
    ("South Africa",          "Elanthendral", "South Africa",      "South Africa"),
    ("Israel",                "Elanthendral", "Israel",            "Israel"),
    ("Rest Of Europe",        "Sathish",      "Region - Europe",   "Rest Of Europe"),
    ("Poland",                "Sathish",      "Poland",            "Poland"),
    ("Rest Of MEA",           "Indhu",        "Region - MEA",      "Rest Of MEA"),
    ("Rest Of APAC",          "Indhu",        "Region - APAC",     "Rest Of APAC"),
    ("Singapore",             "Suganesh",     "Singapore",         "Singapore"),
]

# ── DM Region CASE expression for salesleads_qt ──────────────────────────────
# Maps COMMON_COUNTRY_NAME to a DM Region display label matching DM_REGIONS above.
# Individual countries map 1:1. Regional buckets catch everything else in that geography.
# Presales markets (US/IN/UK/CA/AU) are excluded via WHEN clause returning NULL → filtered.
DM_REGION_CASE = """CASE
  WHEN COMMON_COUNTRY_NAME IN ('United States','India','United Kingdom','Canada','Australia') THEN NULL
  WHEN COMMON_COUNTRY_NAME = 'Germany'              THEN 'Germany'
  WHEN COMMON_COUNTRY_NAME = 'Netherlands'          THEN 'Netherlands'
  WHEN COMMON_COUNTRY_NAME = 'Switzerland'          THEN 'Switzerland'
  WHEN COMMON_COUNTRY_NAME = 'Belgium'              THEN 'Belgium'
  WHEN COMMON_COUNTRY_NAME = 'France'               THEN 'France'
  WHEN COMMON_COUNTRY_NAME = 'Italy'                THEN 'Italy'
  WHEN COMMON_COUNTRY_NAME = 'United Arab Emirates' THEN 'United Arab Emirates'
  WHEN COMMON_COUNTRY_NAME = 'Saudi Arabia'         THEN 'Saudi Arabia'
  WHEN COMMON_COUNTRY_NAME = 'Turkey'               THEN 'Turkey'
  WHEN COMMON_COUNTRY_NAME = 'Spain'                THEN 'Spain'
  WHEN COMMON_COUNTRY_NAME = 'Brazil'               THEN 'Brazil'
  WHEN COMMON_COUNTRY_NAME = 'Mexico'               THEN 'Mexico'
  WHEN COMMON_COUNTRY_NAME = 'South Africa'         THEN 'South Africa'
  WHEN COMMON_COUNTRY_NAME = 'Israel'               THEN 'Israel'
  WHEN COMMON_COUNTRY_NAME = 'Poland'               THEN 'Poland'
  WHEN COMMON_COUNTRY_NAME = 'Singapore'            THEN 'Singapore'
  WHEN COMMON_COUNTRY_NAME IN (
    'Sweden','Norway','Denmark','Finland','Austria','Czech Republic','Hungary',
    'Romania','Bulgaria','Greece','Portugal','Croatia','Slovakia','Slovenia',
    'Estonia','Latvia','Lithuania','Luxembourg','Malta','Cyprus','Albania',
    'Serbia','Bosnia and Herzegovina','North Macedonia','Moldova','Ukraine',
    'Belarus','Iceland','Ireland','Russia','Kosovo','Montenegro','Andorra','Liechtenstein')
                                                    THEN 'Rest Of Europe'
  WHEN COMMON_COUNTRY_NAME IN (
    'Colombia','Peru','Argentina','Chile','Venezuela','Ecuador','Bolivia','Paraguay',
    'Uruguay','Costa Rica','Panama','Honduras','Guatemala','El Salvador','Nicaragua',
    'Puerto Rico','Dominican Republic','Cuba','Trinidad and Tobago','Jamaica',
    'Belize','Haiti','Guyana','Suriname','Barbados')
                                                    THEN 'Rest Of LATAM'
  WHEN COMMON_COUNTRY_NAME IN (
    'Egypt','Nigeria','Kenya','Morocco','Ghana','Tanzania','Ethiopia','Uganda',
    'Cameroon','Senegal','Zimbabwe','Zambia','Mozambique','Ivory Coast','Angola',
    'Algeria','Tunisia','Libya','Sudan','Jordan','Lebanon','Kuwait','Qatar',
    'Bahrain','Oman','Iraq','Yemen','Pakistan','Rwanda','Botswana')
                                                    THEN 'Rest Of MEA'
  WHEN COMMON_COUNTRY_NAME IN (
    'China','Japan','South Korea','Vietnam','Thailand','Indonesia','Philippines',
    'Malaysia','Taiwan','Hong Kong','New Zealand','Bangladesh','Sri Lanka',
    'Myanmar','Cambodia','Nepal','Mongolia','Macao','Brunei','Fiji',
    'Papua New Guinea','Laos','Timor-Leste')
                                                    THEN 'Rest Of APAC'
  ELSE NULL
END"""

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
# Created_Time format: "01 Jan 2026 12:00:00"
# Year = SUBSTR(Created_Time,8,4), Month abbrev = SUBSTR(Created_Time,4,3)
# Reconstruct YYYY-MM using a CASE on the 3-letter month abbreviation.
def _sl_ym_expr():
    return """CONCAT(
      SUBSTR(Created_Time,8,4), '-',
      LPAD(CAST(CASE SUBSTR(Created_Time,4,3)
        WHEN 'Jan' THEN 1  WHEN 'Feb' THEN 2  WHEN 'Mar' THEN 3
        WHEN 'Apr' THEN 4  WHEN 'May' THEN 5  WHEN 'Jun' THEN 6
        WHEN 'Jul' THEN 7  WHEN 'Aug' THEN 8  WHEN 'Sep' THEN 9
        WHEN 'Oct' THEN 10 WHEN 'Nov' THEN 11 WHEN 'Dec' THEN 12
      END AS STRING), 2, '0'))"""

def sl_date_filter(months):
    return f"{_sl_ym_expr()} IN ({ym_list(months)})"

# ── Query A: All leads excl E/TP/O (salesleads_qt 4-filter standard) ─────────
def q_leads_excl(months):
    return f"""
SELECT
  dm_region                                                         AS country,
  COUNT(DISTINCT Email)                                             AS leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))           AS convs
FROM (
  SELECT Email, Conversion,
    {DM_REGION_CASE} AS dm_region
  FROM `{SL}`
  WHERE {sl_date_filter(months)}
    AND Junk = 'false'
    AND PRODUCT_GROUP = 'AD_GROUP'
    AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
    AND isHaveToBeRemoved = 'Non Junk Email'
)
WHERE dm_region IS NOT NULL
GROUP BY 1
"""

# ── Query B: All leads (no User_Type / isHaveToBeRemoved filter) ──────────────
def q_leads_all(months):
    return f"""
SELECT
  dm_region                                                         AS country,
  COUNT(DISTINCT Email)                                             AS leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))           AS convs
FROM (
  SELECT Email, Conversion,
    {DM_REGION_CASE} AS dm_region
  FROM `{SL}`
  WHERE {sl_date_filter(months)}
    AND Junk = 'false'
    AND PRODUCT_GROUP = 'AD_GROUP'
)
WHERE dm_region IS NOT NULL
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
  dm_region                                                         AS country,
  COUNT(DISTINCT Email)                                             AS sem_leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL))           AS sem_convs
FROM (
  SELECT Email, Conversion,
    {DM_REGION_CASE} AS dm_region
  FROM `{SL}`
  WHERE {sl_date_filter(months)}
    AND FIRST_SRC_GRP IN ('google / cpc','bing / cpc')
    AND FIRST_SRC_THEME NOT IN {BRAND}
    AND (FIRST_SRC_CAMPAIGN_TYPE NOT IN ('Display','Performance Max') OR FIRST_SRC_CAMPAIGN_TYPE IS NULL)
    AND Junk = 'false'
    AND PRODUCT_GROUP = 'AD_GROUP'
    AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
    AND isHaveToBeRemoved = 'Non Junk Email'
)
WHERE dm_region IS NOT NULL
GROUP BY 1
"""

# ── Fetch all data ─────────────────────────────────────────────────────────────
def fetch(months, label_str):
    print(f"  Fetching {label_str}...")
    a = idx(run(q_leads_excl(months)),      'country')
    b = idx(run(q_leads_all(months)),       'country')
    c_spend = idx(run(q_sem_spend(months)), 'country')
    c_leads = idx(run(q_sem_leads_themes(months)), 'country')
    c_convs = idx(run(q_sem_convs(months)), 'country')   # now returns dm_region as 'country'
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

# ── Excel styles (verified against canonical july2026_yoy.xlsx) ───────────────
# Fills — exact hex from canonical
YELLOW   = PatternFill("solid", fgColor="FFD700")   # title row
PINK_H   = PatternFill("solid", fgColor="FFB6C1")   # section A header + data bg
LBLUE_H  = PatternFill("solid", fgColor="D9E8F5")   # section B header + data bg
GREEN_H  = PatternFill("solid", fgColor="90EE90")   # section C header + data bg
GREEN_D  = PatternFill("solid", fgColor="C6EFCE")   # ≥+10% YoY (good)
PINK_D   = PatternFill("solid", fgColor="FFC7CE")   # ≤-10% YoY (bad)
NAVY     = PatternFill("solid", fgColor="1F4E79")   # column header row

def af(bold=False, size=10, color="000000"):
    return Font(bold=bold, size=size, color=color)

def ac(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def rich_cell(ws, row, col, number, yoy_str, fill=None):
    """
    Canonical format: bold 11pt dark (#1A1A1A) number + normal 8pt grey (#808080) YoY%.
    No YoY suffix when no prior year data. Dash when number=0.
    """
    c = ws.cell(row=row, column=col)
    c.alignment = ac('center')
    if fill: c.fill = fill

    if number == 0:
        c.value = "—"
        c.font = af(size=10, color="808080")
        return c

    num_str = f"{number:,}"
    if yoy_str == "—":
        # No prior year data — plain bold number only
        c.value = num_str
        c.font = af(bold=True, size=11, color="1A1A1A")
        return c

    if RICH_TEXT:
        # sz in InlineFont = actual pt (not half-points), matching canonical
        c.value = CellRichText(
            TextBlock(InlineFont(b=True,  sz=11, color="1A1A1A"), num_str),
            TextBlock(InlineFont(b=False, sz=8,  color="808080"), f" ({yoy_str})"),
        )
    else:
        c.value = f"{num_str} ({yoy_str})"
        c.font = af(bold=True, size=11, color="1A1A1A")
    return c

def dollar_rich(ws, row, col, number, yoy_str, fill=None):
    """Same as rich_cell but formats number as $xxx,xxx."""
    c = ws.cell(row=row, column=col)
    c.alignment = ac('center')
    if fill: c.fill = fill

    if number == 0:
        c.value = "—"
        c.font = af(size=10, color="808080")
        return c

    num_str = f"${number:,.0f}"
    if yoy_str == "—":
        c.value = num_str
        c.font = af(bold=True, size=11, color="1A1A1A")
        return c

    if RICH_TEXT:
        c.value = CellRichText(
            TextBlock(InlineFont(b=True,  sz=11, color="1A1A1A"), num_str),
            TextBlock(InlineFont(b=False, sz=8,  color="808080"), f" ({yoy_str})"),
        )
    else:
        c.value = f"{num_str} ({yoy_str})"
        c.font = af(bold=True, size=11, color="1A1A1A")
    return c

# ── QA checklist (called before saving) ───────────────────────────────────────
def qa_check(ws, n_data_rows):
    errors = []
    # Title in B2
    if not ws['B2'].value:
        errors.append("B2 missing title")
    # Section headers in row 3
    if not ws['D3'].value:
        errors.append("D3 missing section A header")
    # Data starts at row 5
    if not ws['B5'].value:
        errors.append("B5 missing first DM Region")
    # Check expected row count
    last_row = 4 + n_data_rows
    if not ws.cell(row=last_row, column=2).value:
        errors.append(f"Row {last_row} missing last DM Region")
    if errors:
        print(f"  ⚠️  QA WARNINGS: {errors}")
    else:
        print(f"  ✅ QA passed ({n_data_rows} data rows, title={ws['B2'].value!r})")

# ── Build a single worksheet (canonical format) ───────────────────────────────
def build_sheet(ws, title,
                cur_a, cur_b, cur_spend, cur_csl, cur_csc,
                pyr_a, pyr_b, pyr_spend, pyr_csl, pyr_csc):
    """
    Canonical layout (from july2026_yoy.xlsx analysis):
      Col A: blank spacer (width 13)
      Col B: DM Region (width 22)
      Col C: DRI (width 14)
      Col D: Section A — Leads (width 20)
      Col E: Section A — Convs (width 13)
      Col F: Section B — All Leads (width 13)
      Col G: Section B — All Convs (width 13)
      Col H: Section C — SEM Spend (width 13)
      Col I: Section C — SEM Leads (width 13)
      Col J: Section C — SEM Convs (width 13)

      Row 1: blank
      Row 2: Title merged B2:J2, yellow fill, bold 14pt
      Row 3: Section group headers — D3:E3 pink, F3:G3 blue, H3:J3 green
      Row 4: Column headers — navy fill (FF1F4E79), bold 10pt white
      Row 5+: Data rows
    """

    # ── Row 2: Title ──────────────────────────────────────────────────────────
    ws.merge_cells('B2:J2')
    t = ws.cell(row=2, column=2, value=title)
    t.fill = YELLOW
    t.font = af(bold=True, size=14)
    t.alignment = ac('center')
    ws.row_dimensions[2].height = 28

    # ── Row 3: Section group headers ─────────────────────────────────────────
    ws.merge_cells('D3:E3')
    s1 = ws.cell(row=3, column=4, value="All leads except Events, Third Party & Others")
    s1.fill = PINK_H; s1.font = af(bold=True, size=10); s1.alignment = ac('center')

    ws.merge_cells('F3:G3')
    s2 = ws.cell(row=3, column=6, value="All leads")
    s2.fill = LBLUE_H; s2.font = af(bold=True, size=10); s2.alignment = ac('center')

    ws.merge_cells('H3:J3')
    s3 = ws.cell(row=3, column=8, value="SEM leads")
    s3.fill = GREEN_H; s3.font = af(bold=True, size=10); s3.alignment = ac('center')
    ws.row_dimensions[3].height = 28

    # ── Row 4: Column headers ─────────────────────────────────────────────────
    COL_DEFS = [
        (2, "DM Region",                     NAVY),
        (3, "DRI",                            NAVY),
        (4, "Leads\n(Created date)",          PINK_H),
        (5, "Conversions\n(Conv. Date)",      PINK_H),
        (6, "All Leads\n(Created date)",      LBLUE_H),
        (7, "All Conversions\n(Conv. Date)",  LBLUE_H),
        (8, "SEM Spending\n(Google USD)",     GREEN_H),
        (9, "SEM Leads\n(Created date)",      GREEN_H),
        (10,"SEM Conversions\n(Conv. Date)",  GREEN_H),
    ]
    for col, hdr, fill in COL_DEFS:
        c = ws.cell(row=4, column=col, value=hdr)
        c.fill = fill
        # Navy header cells get white bold text; section-coloured cells get dark bold text
        txt_color = "FFFFFF" if fill == NAVY else "000000"
        c.font = af(bold=True, size=10, color=txt_color)
        c.alignment = ac('center', wrap=True)
    ws.row_dimensions[4].height = 28

    # ── Data rows (start at row 5) ────────────────────────────────────────────
    for i, (display, dri, themes_key, sl_key) in enumerate(DM_REGIONS):
        row = i + 5

        # Section A
        al_c = gv(cur_a, display, 'leads');  al_p = gv(pyr_a, display, 'leads')
        ac_c = gv(cur_a, display, 'convs');  ac_p = gv(pyr_a, display, 'convs')
        # Section B
        bl_c = gv(cur_b, display, 'leads');  bl_p = gv(pyr_b, display, 'leads')
        bc_c = gv(cur_b, display, 'convs');  bc_p = gv(pyr_b, display, 'convs')
        # Section C
        sp_c = gv(cur_spend, themes_key, 'spend_usd');  sp_p = gv(pyr_spend, themes_key, 'spend_usd')
        sl_c = gv(cur_csl,   themes_key, 'sem_leads');  sl_p = gv(pyr_csl,   themes_key, 'sem_leads')
        sc_c = gv(cur_csc,   sl_key,     'sem_convs');  sc_p = gv(pyr_csc,   sl_key,     'sem_convs')

        al_pct, al_str = yoy_pct(al_c, al_p);  ac_pct, ac_str = yoy_pct(ac_c, ac_p)
        bl_pct, bl_str = yoy_pct(bl_c, bl_p);  bc_pct, bc_str = yoy_pct(bc_c, bc_p)
        sp_pct, sp_str = yoy_pct(sp_c, sp_p);  sl_pct, sl_str = yoy_pct(sl_c, sl_p)
        sc_pct, sc_str = yoy_pct(sc_c, sc_p)

        def fill_for(pct, section_fill):
            if pct is None: return section_fill
            if pct >= 10:  return GREEN_D
            if pct <= -10: return PINK_D
            return section_fill

        # Col A — blank spacer
        ws.cell(row=row, column=1).value = None

        # Col B — DM Region
        c = ws.cell(row=row, column=2, value=display)
        c.font = af(bold=True, size=10); c.alignment = ac('left', 'center')

        # Col C — DRI
        c = ws.cell(row=row, column=3, value=dri)
        c.font = af(size=10); c.alignment = ac('left', 'center')

        # Cols D–E — Section A (excl E/TP/O)
        rich_cell(ws, row, 4, al_c, al_str, fill=fill_for(al_pct, PINK_H))
        rich_cell(ws, row, 5, ac_c, ac_str, fill=fill_for(ac_pct, PINK_H))

        # Cols F–G — Section B (all leads)
        rich_cell(ws, row, 6, bl_c, bl_str, fill=fill_for(bl_pct, LBLUE_H))
        rich_cell(ws, row, 7, bc_c, bc_str, fill=fill_for(bc_pct, LBLUE_H))

        # Cols H–J — Section C (SEM)
        dollar_rich(ws, row, 8,  sp_c, sp_str, fill=fill_for(sp_pct, GREEN_H))
        rich_cell(ws,  row, 9,  sl_c, sl_str, fill=fill_for(sl_pct, GREEN_H))
        rich_cell(ws,  row, 10, sc_c, sc_str, fill=fill_for(sc_pct, GREEN_H))

        ws.row_dimensions[row].height = 28

    # ── Column widths (from canonical) ────────────────────────────────────────
    for col, w in zip('ABCDEFGHIJ', [13, 22, 14, 20, 13, 13, 13, 13, 13, 13]):
        ws.column_dimensions[col].width = w

    # ── QA ────────────────────────────────────────────────────────────────────
    qa_check(ws, len(DM_REGIONS))

    ws.freeze_panes = 'D5'

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
