"""
ROW IDM & SIEM — July 2026 Excel Report
=========================================
Adds two new tabs to jul2026_yoy.xlsx:
  - "IDM Jul 2026"  : DM Region × IDM products (ADMP group + ADSSP group)
  - "SIEM Jul 2026" : DM Region × SIEM products (ADAP group + ELA group)

Columns (same format as canonical july2026_yoy.xlsx):
  Section A (pink)  : All leads excl E/TP/O — salesleads_qt 4-filter
  Section B (blue)  : All leads — salesleads_qt Junk+AD_GROUP only
  Section C (green) : SEM leads — themes table spend + leads

Product group mapping (confirmed by user):
  IDM  = ADMP group (admp, rmp, mmp, spmp, ad360) + ADSSP group (adssp, id360)
  SIEM = ADAP group (adap, dsp, erp) + ELA group (ela, log360, log360cloud, log360 mssp, csp)

Usage:
    python3 row_idm_siem_xlsx.py                  # uses 2026-07 (config.yaml)
    python3 row_idm_siem_xlsx.py --month 2026-07  # explicit month
"""
import sys, os, argparse
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '../../wsm-monitor'))
import wsm_cfg as cfg
from wsm_cfg import bq_client, shift, label, mon

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.rich_text import InlineFont, CellRichText, TextBlock
    RICH_TEXT = True
except ImportError:
    sys.exit("Missing openpyxl. Run: pip install openpyxl")

# ── CLI ───────────────────────────────────────────────────────────────────────
ap = argparse.ArgumentParser()
ap.add_argument('--month', default=None)
args = ap.parse_args()

CUR   = args.month or cfg.CUR
PYR   = shift(CUR, -12)
MON   = mon(CUR)
YEAR  = CUR[:4]
PYEAR = PYR[:4]
PYR_SHORT = PYEAR[2:]

XLSX = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    f"../../{MON.lower()}{YEAR}_yoy.xlsx"
))
print(f"ROW IDM/SIEM — {label(CUR)} vs {label(PYR)}")
print(f"File: {XLSX}")

bq = bq_client()
SL  = f"{cfg.PROJ}.sales_presales_leads_no_pi.salesleads_qt"
ROI = f"{cfg.PROJ}.{cfg.G.split('.')[-1]}.themes_firstlast_semroi"

BRAND = "('Branding','Log360 - Branding','Cloud Branding','cloud branding','ELA - Branding','AD360 - Branding','AD360 Branding')"

# ── DM Region order & display mapping ─────────────────────────────────────────
DM_REGIONS = [
    ("Germany",              "Jude",         "Germany",           "Germany"),
    ("Netherlands",          "Jude",         "Netherlands",       "Netherlands"),
    ("Switzerland",          "Jude",         "Switzerland",       "Switzerland"),
    ("Belgium",              "Jude",         "Belgium",           "Belgium"),
    ("France",               "Kowsik",       "France",            "France"),
    ("Italy",                "Kowsik",       "Italy",             "Italy"),
    ("United Arab Emirates", "Kowsik",       "United Arab Emirates","United Arab Emirates"),
    ("Saudi Arabia",         "Kowsik",       "Saudi Arabia",      "Saudi Arabia"),
    ("Turkey",               "Kowsik",       "Turkey",            "Turkey"),
    ("Spain",                "Elanthendral", "Spain",             "Spain"),
    ("Brazil",               "Elanthendral", "Brazil",            "Brazil"),
    ("Mexico",               "Elanthendral", "Mexico",            "Mexico"),
    ("Rest Of LATAM",        "Elanthendral", "Region - LATAM",    "Rest Of LATAM"),
    ("South Africa",         "Elanthendral", "South Africa",      "South Africa"),
    ("Israel",               "Elanthendral", "Israel",            "Israel"),
    ("Rest Of Europe",       "Sathish",      "Region - Europe",   "Rest Of Europe"),
    ("Poland",               "Sathish",      "Poland",            "Poland"),
    ("Rest Of MEA",          "Indhu",        "Region - MEA",      "Rest Of MEA"),
    ("Rest Of APAC",         "Indhu",        "Region - APAC",     "Rest Of APAC"),
    ("Singapore",            "Suganesh",     "Singapore",         "Singapore"),
]

# ── Product group definitions ──────────────────────────────────────────────────
IDM_PRODUCTS  = ("'admp'","'rmp'","'mmp'","'spmp'","'ad360'","'adssp'","'id360'")
SIEM_PRODUCTS = ("'adap'","'dsp'","'erp'","'ela'","'log360'","'log360cloud'","'log360 mssp'","'csp'")

IDM_THEMES  = ("'ADMP'","'ADSSP'")   # themes table Product values for IDM SEM
SIEM_THEMES = ("'ADAP'","'ELA'","'LOG360'","'LOG360CLOUD'")  # for SIEM SEM

# ── Date filter helper ─────────────────────────────────────────────────────────
def sl_ym_expr():
    return """CONCAT(
      SUBSTR(Created_Time,8,4), '-',
      LPAD(CAST(CASE SUBSTR(Created_Time,4,3)
        WHEN 'Jan' THEN 1  WHEN 'Feb' THEN 2  WHEN 'Mar' THEN 3
        WHEN 'Apr' THEN 4  WHEN 'May' THEN 5  WHEN 'Jun' THEN 6
        WHEN 'Jul' THEN 7  WHEN 'Aug' THEN 8  WHEN 'Sep' THEN 9
        WHEN 'Oct' THEN 10 WHEN 'Nov' THEN 11 WHEN 'Dec' THEN 12
      END AS STRING), 2, '0'))"""

DM_CASE = """CASE
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

def run(sql):
    return list(bq.query(sql).result())

def idx(rows, key):
    return {r[key]: r for r in rows}

# ── BQ queries ─────────────────────────────────────────────────────────────────
def q_leads_excl(ym, products):
    prod_list = ", ".join(products)
    return f"""
SELECT dm_region AS country,
  COUNT(DISTINCT Email) AS leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL)) AS convs
FROM (
  SELECT Email, Conversion, {DM_CASE} AS dm_region
  FROM `{SL}`
  WHERE {sl_ym_expr()} = '{ym}'
    AND Junk = 'false' AND PRODUCT_GROUP = 'AD_GROUP'
    AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
    AND isHaveToBeRemoved = 'Non Junk Email'
    AND LOWER(PRODUCT) IN ({prod_list})
) WHERE dm_region IS NOT NULL GROUP BY 1"""

def q_leads_all(ym, products):
    prod_list = ", ".join(products)
    return f"""
SELECT dm_region AS country,
  COUNT(DISTINCT Email) AS leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL)) AS convs
FROM (
  SELECT Email, Conversion, {DM_CASE} AS dm_region
  FROM `{SL}`
  WHERE {sl_ym_expr()} = '{ym}'
    AND Junk = 'false' AND PRODUCT_GROUP = 'AD_GROUP'
    AND LOWER(PRODUCT) IN ({prod_list})
) WHERE dm_region IS NOT NULL GROUP BY 1"""

def q_sem_spend(ym, theme_products):
    prod_list = ", ".join(f"'{p}'" for p in theme_products)
    return f"""
SELECT CampaignCountry AS country,
  ROUND(SUM(Cost), 0) AS spend_usd
FROM `{ROI}`
WHERE SUBSTR(Date,1,7) = '{ym}'
  AND Lead_Type = 'All Leads'
  AND CampaignCountry NOT IN ('United States','India','United Kingdom','Canada','Australia')
  AND Theme NOT IN {BRAND}
  AND Product IN ({prod_list})
GROUP BY 1"""

def q_sem_leads(ym, theme_products):
    prod_list = ", ".join(f"'{p}'" for p in theme_products)
    return f"""
SELECT CampaignCountry AS country,
  ROUND(SUM(Valid_Sales_Leads_First_Source)) AS sem_leads
FROM `{ROI}`
WHERE SUBSTR(Date,1,7) = '{ym}'
  AND Lead_Type = 'All Leads'
  AND CampaignCountry NOT IN ('United States','India','United Kingdom','Canada','Australia')
  AND Theme NOT IN {BRAND}
  AND Product IN ({prod_list})
GROUP BY 1"""

def q_sem_convs(ym, products):
    prod_list = ", ".join(products)
    return f"""
SELECT dm_region AS country,
  COUNT(DISTINCT Email) AS sem_leads,
  COUNT(DISTINCT IF(Conversion='converted', Email, NULL)) AS sem_convs
FROM (
  SELECT Email, Conversion, {DM_CASE} AS dm_region
  FROM `{SL}`
  WHERE {sl_ym_expr()} = '{ym}'
    AND Junk = 'false' AND PRODUCT_GROUP = 'AD_GROUP'
    AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
    AND isHaveToBeRemoved = 'Non Junk Email'
    AND LOWER(PRODUCT) IN ({prod_list})
    AND FIRST_SRC_GRP IN ('google / cpc','bing / cpc')
    AND FIRST_SRC_THEME NOT IN {BRAND}
    AND (FIRST_SRC_CAMPAIGN_TYPE NOT IN ('Display','Performance Max') OR FIRST_SRC_CAMPAIGN_TYPE IS NULL)
) WHERE dm_region IS NOT NULL GROUP BY 1"""

def fetch_group(ym, sl_products, themes_products):
    a = idx(run(q_leads_excl(ym, sl_products)), 'country')
    b = idx(run(q_leads_all(ym, sl_products)),  'country')
    sp = idx(run(q_sem_spend(ym, themes_products)), 'country')
    sl = idx(run(q_sem_leads(ym, themes_products)), 'country')
    sc = idx(run(q_sem_convs(ym, sl_products)), 'country')
    return a, b, sp, sl, sc

print("\nFetching IDM data...")
idm_cur = fetch_group(CUR, IDM_PRODUCTS,  [p.strip("'") for p in IDM_THEMES])
idm_pyr = fetch_group(PYR, IDM_PRODUCTS,  [p.strip("'") for p in IDM_THEMES])

print("Fetching SIEM data...")
siem_cur = fetch_group(CUR, SIEM_PRODUCTS, [p.strip("'") for p in SIEM_THEMES])
siem_pyr = fetch_group(PYR, SIEM_PRODUCTS, [p.strip("'") for p in SIEM_THEMES])

print("Done.\n")

# ── Excel style helpers (canonical format from july2026_yoy.xlsx) ───────────────
YELLOW  = PatternFill("solid", fgColor="FFD700")
PINK_H  = PatternFill("solid", fgColor="FFB6C1")
LBLUE_H = PatternFill("solid", fgColor="D9E8F5")
GREEN_H = PatternFill("solid", fgColor="90EE90")
GREEN_D = PatternFill("solid", fgColor="C6EFCE")
PINK_D  = PatternFill("solid", fgColor="FFC7CE")
NAVY    = PatternFill("solid", fgColor="1F4E79")

def af(bold=False, size=10, color="000000"):
    return Font(bold=bold, size=size, color=color)

def ac(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def gv(d, key, col, default=0):
    r = d.get(key)
    if r is None: return default
    v = r[col]
    return int(v) if v is not None else default

def yoy_pct(cur, pyr):
    if not pyr: return None, "—"
    p = (cur - pyr) / pyr * 100
    return p, f"{'+' if p >= 0 else ''}{p:.0f}%"

def rich_cell(ws, row, col, number, yoy_str, fill=None):
    c = ws.cell(row=row, column=col)
    c.alignment = ac('center')
    if fill: c.fill = fill
    if number == 0:
        c.value = "—"; c.font = af(size=10, color="808080"); return c
    num_str = f"{number:,}"
    if yoy_str == "—":
        c.value = num_str; c.font = af(bold=True, size=11, color="1A1A1A"); return c
    if RICH_TEXT:
        c.value = CellRichText(
            TextBlock(InlineFont(b=True,  sz=11, color="1A1A1A"), num_str),
            TextBlock(InlineFont(b=False, sz=8,  color="808080"), f" ({yoy_str})"),
        )
    else:
        c.value = f"{num_str} ({yoy_str})"; c.font = af(bold=True, size=11, color="1A1A1A")
    return c

def dollar_rich(ws, row, col, number, yoy_str, fill=None):
    c = ws.cell(row=row, column=col)
    c.alignment = ac('center')
    if fill: c.fill = fill
    if number == 0:
        c.value = "—"; c.font = af(size=10, color="808080"); return c
    num_str = f"${number:,.0f}"
    if yoy_str == "—":
        c.value = num_str; c.font = af(bold=True, size=11, color="1A1A1A"); return c
    if RICH_TEXT:
        c.value = CellRichText(
            TextBlock(InlineFont(b=True,  sz=11, color="1A1A1A"), num_str),
            TextBlock(InlineFont(b=False, sz=8,  color="808080"), f" ({yoy_str})"),
        )
    else:
        c.value = f"{num_str} ({yoy_str})"; c.font = af(bold=True, size=11, color="1A1A1A")
    return c

# ── Build a product group tab ──────────────────────────────────────────────────
def build_tab(ws, title, cur_data, pyr_data):
    cur_a, cur_b, cur_sp, cur_sl, cur_sc = cur_data
    pyr_a, pyr_b, pyr_sp, pyr_sl, pyr_sc = pyr_data

    # Row 2: title
    ws.merge_cells('B2:J2')
    t = ws.cell(row=2, column=2, value=title)
    t.fill = YELLOW; t.font = af(bold=True, size=14); t.alignment = ac('center')
    ws.row_dimensions[2].height = 28

    # Row 3: section group headers
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

    # Row 4: column headers
    hdrs = [
        (2, "DM Region",                    NAVY,    "FFFFFF"),
        (3, "DRI",                           NAVY,    "FFFFFF"),
        (4, "Leads\n(Created date)",         PINK_H,  "000000"),
        (5, "Conversions\n(Conv. Date)",     PINK_H,  "000000"),
        (6, "All Leads\n(Created date)",     LBLUE_H, "000000"),
        (7, "All Conversions\n(Conv. Date)", LBLUE_H, "000000"),
        (8, "SEM Spending\n(Google USD)",    GREEN_H, "000000"),
        (9, "SEM Leads\n(Created date)",     GREEN_H, "000000"),
        (10,"SEM Conversions\n(Conv. Date)", GREEN_H, "000000"),
    ]
    for col, hdr, fill, txt in hdrs:
        c = ws.cell(row=4, column=col, value=hdr)
        c.fill = fill; c.font = af(bold=True, size=10, color=txt)
        c.alignment = ac('center', wrap=True)
    ws.row_dimensions[4].height = 28

    # Data rows
    for i, (display, dri, themes_key, sl_key) in enumerate(DM_REGIONS):
        row = i + 5

        al_c = gv(cur_a, display, 'leads');   al_p = gv(pyr_a, display, 'leads')
        ac_c = gv(cur_a, display, 'convs');   ac_p = gv(pyr_a, display, 'convs')
        bl_c = gv(cur_b, display, 'leads');   bl_p = gv(pyr_b, display, 'leads')
        bc_c = gv(cur_b, display, 'convs');   bc_p = gv(pyr_b, display, 'convs')
        sp_c = gv(cur_sp, themes_key, 'spend_usd');  sp_p = gv(pyr_sp, themes_key, 'spend_usd')
        sl_c = gv(cur_sl, themes_key, 'sem_leads');  sl_p = gv(pyr_sl, themes_key, 'sem_leads')
        sc_c = gv(cur_sc, display, 'sem_convs');     sc_p = gv(pyr_sc, display, 'sem_convs')

        al_pct, al_str = yoy_pct(al_c, al_p);  ac_pct, ac_str = yoy_pct(ac_c, ac_p)
        bl_pct, bl_str = yoy_pct(bl_c, bl_p);  bc_pct, bc_str = yoy_pct(bc_c, bc_p)
        sp_pct, sp_str = yoy_pct(sp_c, sp_p);  sl_pct, sl_str = yoy_pct(sl_c, sl_p)
        sc_pct, sc_str = yoy_pct(sc_c, sc_p)

        def fill_for(pct, section_fill):
            if pct is None: return section_fill
            if pct >= 10: return GREEN_D
            if pct <= -10: return PINK_D
            return section_fill

        c = ws.cell(row=row, column=2, value=display)
        c.font = af(bold=True, size=10); c.alignment = ac('left', 'center')
        c = ws.cell(row=row, column=3, value=dri)
        c.font = af(size=10); c.alignment = ac('left', 'center')

        rich_cell(ws, row, 4, al_c, al_str, fill=fill_for(al_pct, PINK_H))
        rich_cell(ws, row, 5, ac_c, ac_str, fill=fill_for(ac_pct, PINK_H))
        rich_cell(ws, row, 6, bl_c, bl_str, fill=fill_for(bl_pct, LBLUE_H))
        rich_cell(ws, row, 7, bc_c, bc_str, fill=fill_for(bc_pct, LBLUE_H))
        dollar_rich(ws, row, 8, sp_c, sp_str, fill=fill_for(sp_pct, GREEN_H))
        rich_cell(ws, row, 9, sl_c, sl_str, fill=fill_for(sl_pct, GREEN_H))
        rich_cell(ws, row, 10, sc_c, sc_str, fill=fill_for(sc_pct, GREEN_H))
        ws.row_dimensions[row].height = 28

    for col, w in zip('ABCDEFGHIJ', [13, 22, 14, 20, 13, 13, 13, 13, 13, 13]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'D5'
    print(f"  ✅ Tab '{ws.title}' — {len(DM_REGIONS)} rows")

# ── Open existing workbook and add tabs ────────────────────────────────────────
if os.path.exists(XLSX):
    wb = openpyxl.load_workbook(XLSX, rich_text=True)
else:
    wb = openpyxl.Workbook()
    wb.active.title = "DM Regions"

# Remove existing IDM/SIEM tabs if present
for tab_name in [f"IDM {MON} {YEAR}", f"SIEM {MON} {YEAR}"]:
    if tab_name in wb.sheetnames:
        del wb[tab_name]

# Add IDM tab
ws_idm = wb.create_sheet(f"IDM {MON} {YEAR}")
idm_title = f"IDM — {MON} {YEAR}  (vs {MON} '{PYR_SHORT})"
build_tab(ws_idm, idm_title, idm_cur, idm_pyr)

# Add SIEM tab
ws_siem = wb.create_sheet(f"SIEM {MON} {YEAR}")
siem_title = f"SIEM — {MON} {YEAR}  (vs {MON} '{PYR_SHORT})"
build_tab(ws_siem, siem_title, siem_cur, siem_pyr)

wb.save(XLSX)
print(f"\n📊 Saved: {XLSX}")
print(f"   Tabs added: 'IDM {MON} {YEAR}' | 'SIEM {MON} {YEAR}'")
print(f"   IDM  = admp, rmp, mmp, spmp, ad360, adssp, id360")
print(f"   SIEM = adap, dsp, erp, ela, log360, log360cloud, log360 mssp, csp")
