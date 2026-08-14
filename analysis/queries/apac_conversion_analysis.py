"""
APAC 30/60/90-Day Conversion Analysis
======================================
Countries: Vietnam, Singapore, Philippines, Indonesia, Malaysia, Thailand
Period: Jan–Jul for each year (2024, 2025, 2026)
Metric: Leads + Convs (30d / 60d / 90d) + Conv%
Output: apac_conversion_jul2026.xlsx
"""
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '../../wsm-monitor'))
from wsm_cfg import PROJ, bq_client

try:
    import openpyxl
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter
    from openpyxl.cell.rich_text import InlineFont, CellRichText, TextBlock
    RICH_TEXT = True
except ImportError:
    sys.exit("pip install openpyxl")

bq = bq_client()
SL = f"{PROJ}.sales_presales_leads_no_pi.salesleads_qt"

COUNTRIES = ['Vietnam', 'Singapore', 'Philippines', 'Indonesia', 'Malaysia', 'Thailand']
YEARS     = ['2024', '2025', '2026']
APAC_TOTAL_KEY = '__APAC_TOTAL__'

OUT = os.path.normpath(os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "../../apac_conversion_jul2026.xlsx"
))

print("Fetching data...")

sql = f"""
SELECT
  COMMON_COUNTRY_NAME AS country,
  SUBSTR(Created_Time,8,4) AS yr,
  COUNT(DISTINCT NON_JUNK_EMAIL) AS leads,
  COUNT(DISTINCT IF(
    Conversion = 'converted'
    AND ConversionDate IS NOT NULL AND ConversionDate != ''
    AND SAFE.PARSE_DATE('%d %b %Y', SUBSTR(ConversionDate,1,11)) IS NOT NULL
    AND DATE_DIFF(
      SAFE.PARSE_DATE('%d %b %Y', SUBSTR(ConversionDate,1,11)),
      SAFE.PARSE_DATE('%d %b %Y', SUBSTR(Created_Time,1,11)),
      DAY) BETWEEN 0 AND 30,
    NON_JUNK_EMAIL, NULL)) AS convs_30d,
  COUNT(DISTINCT IF(
    Conversion = 'converted'
    AND ConversionDate IS NOT NULL AND ConversionDate != ''
    AND SAFE.PARSE_DATE('%d %b %Y', SUBSTR(ConversionDate,1,11)) IS NOT NULL
    AND DATE_DIFF(
      SAFE.PARSE_DATE('%d %b %Y', SUBSTR(ConversionDate,1,11)),
      SAFE.PARSE_DATE('%d %b %Y', SUBSTR(Created_Time,1,11)),
      DAY) BETWEEN 0 AND 60,
    NON_JUNK_EMAIL, NULL)) AS convs_60d,
  COUNT(DISTINCT IF(
    Conversion = 'converted'
    AND ConversionDate IS NOT NULL AND ConversionDate != ''
    AND SAFE.PARSE_DATE('%d %b %Y', SUBSTR(ConversionDate,1,11)) IS NOT NULL
    AND DATE_DIFF(
      SAFE.PARSE_DATE('%d %b %Y', SUBSTR(ConversionDate,1,11)),
      SAFE.PARSE_DATE('%d %b %Y', SUBSTR(Created_Time,1,11)),
      DAY) BETWEEN 0 AND 90,
    NON_JUNK_EMAIL, NULL)) AS convs_90d
FROM `{SL}`
WHERE Junk = 'false'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND Region = 'APAC'
  AND COMMON_COUNTRY_NAME IN ('Vietnam','Singapore','Philippines','Indonesia','Malaysia','Thailand')
  AND SUBSTR(Created_Time,8,4) IN ('2024','2025','2026')
  AND CONCAT(SUBSTR(Created_Time,8,4), '-',
      LPAD(CAST(CASE SUBSTR(Created_Time,4,3)
        WHEN 'Jan' THEN 1  WHEN 'Feb' THEN 2  WHEN 'Mar' THEN 3
        WHEN 'Apr' THEN 4  WHEN 'May' THEN 5  WHEN 'Jun' THEN 6
        WHEN 'Jul' THEN 7  WHEN 'Aug' THEN 8  WHEN 'Sep' THEN 9
        WHEN 'Oct' THEN 10 WHEN 'Nov' THEN 11 WHEN 'Dec' THEN 12
      END AS STRING), 2, '0'))
      BETWEEN CONCAT(SUBSTR(Created_Time,8,4), '-01')
          AND CONCAT(SUBSTR(Created_Time,8,4), '-07')
GROUP BY 1, 2
ORDER BY COMMON_COUNTRY_NAME, yr
"""

rows = list(bq.query(sql).result())

# Index by (country, year)
data = {}
for r in rows:
    data[(r['country'], r['yr'])] = {
        'leads': int(r['leads'] or 0),
        'c30':   int(r['convs_30d'] or 0),
        'c60':   int(r['convs_60d'] or 0),
        'c90':   int(r['convs_90d'] or 0),
    }

# Also fetch APAC Total (all Region='APAC' countries)
sql_total = f"""
SELECT
  'APAC Total' AS country,
  SUBSTR(Created_Time,8,4) AS yr,
  COUNT(DISTINCT NON_JUNK_EMAIL) AS leads,
  COUNT(DISTINCT IF(
    Conversion = 'converted'
    AND ConversionDate IS NOT NULL AND ConversionDate != ''
    AND SAFE.PARSE_DATE('%d %b %Y', SUBSTR(ConversionDate,1,11)) IS NOT NULL
    AND DATE_DIFF(
      SAFE.PARSE_DATE('%d %b %Y', SUBSTR(ConversionDate,1,11)),
      SAFE.PARSE_DATE('%d %b %Y', SUBSTR(Created_Time,1,11)),
      DAY) BETWEEN 0 AND 30,
    NON_JUNK_EMAIL, NULL)) AS convs_30d,
  COUNT(DISTINCT IF(
    Conversion = 'converted'
    AND ConversionDate IS NOT NULL AND ConversionDate != ''
    AND SAFE.PARSE_DATE('%d %b %Y', SUBSTR(ConversionDate,1,11)) IS NOT NULL
    AND DATE_DIFF(
      SAFE.PARSE_DATE('%d %b %Y', SUBSTR(ConversionDate,1,11)),
      SAFE.PARSE_DATE('%d %b %Y', SUBSTR(Created_Time,1,11)),
      DAY) BETWEEN 0 AND 60,
    NON_JUNK_EMAIL, NULL)) AS convs_60d,
  COUNT(DISTINCT IF(
    Conversion = 'converted'
    AND ConversionDate IS NOT NULL AND ConversionDate != ''
    AND SAFE.PARSE_DATE('%d %b %Y', SUBSTR(ConversionDate,1,11)) IS NOT NULL
    AND DATE_DIFF(
      SAFE.PARSE_DATE('%d %b %Y', SUBSTR(ConversionDate,1,11)),
      SAFE.PARSE_DATE('%d %b %Y', SUBSTR(Created_Time,1,11)),
      DAY) BETWEEN 0 AND 90,
    NON_JUNK_EMAIL, NULL)) AS convs_90d
FROM `{SL}`
WHERE Junk = 'false'
  AND PRODUCT_GROUP = 'AD_GROUP'
  AND User_Type IN ('new','adcs','mecs','inactive customer','inactive lead')
  AND isHaveToBeRemoved = 'Non Junk Email'
  AND Region = 'APAC'
  AND SUBSTR(Created_Time,8,4) IN ('2024','2025','2026')
  AND CONCAT(SUBSTR(Created_Time,8,4), '-',
      LPAD(CAST(CASE SUBSTR(Created_Time,4,3)
        WHEN 'Jan' THEN 1  WHEN 'Feb' THEN 2  WHEN 'Mar' THEN 3
        WHEN 'Apr' THEN 4  WHEN 'May' THEN 5  WHEN 'Jun' THEN 6
        WHEN 'Jul' THEN 7  WHEN 'Aug' THEN 8  WHEN 'Sep' THEN 9
        WHEN 'Oct' THEN 10 WHEN 'Nov' THEN 11 WHEN 'Dec' THEN 12
      END AS STRING), 2, '0'))
      BETWEEN CONCAT(SUBSTR(Created_Time,8,4), '-01')
          AND CONCAT(SUBSTR(Created_Time,8,4), '-07')
GROUP BY 1, 2
ORDER BY yr
"""
for r in bq.query(sql_total).result():
    data[(APAC_TOTAL_KEY, r['yr'])] = {
        'leads': int(r['leads'] or 0),
        'c30':   int(r['convs_30d'] or 0),
        'c60':   int(r['convs_60d'] or 0),
        'c90':   int(r['convs_90d'] or 0),
    }

print("Done. Building Excel...")

# ── Styles ────────────────────────────────────────────────────────────────────
YELLOW  = PatternFill("solid", fgColor="FFD700")
NAVY    = PatternFill("solid", fgColor="1F4E79")
PINK    = PatternFill("solid", fgColor="FFB6C1")
BLUE    = PatternFill("solid", fgColor="D9E8F5")
GREEN   = PatternFill("solid", fgColor="90EE90")
GREEN_D = PatternFill("solid", fgColor="C6EFCE")
PINK_D  = PatternFill("solid", fgColor="FFC7CE")
LGREY   = PatternFill("solid", fgColor="F2F2F2")

def af(bold=False, size=10, color="000000"):
    return Font(bold=bold, size=size, color=color)

def ac(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def cell(ws, row, col, value, fill=None, bold=False, size=10, color="000000", h='center', wrap=False):
    c = ws.cell(row=row, column=col, value=value)
    if fill: c.fill = fill
    c.font = Font(bold=bold, size=size, color=color)
    c.alignment = Alignment(horizontal=h, vertical='center', wrap_text=wrap)
    return c

def pct(n, d):
    if not d: return "—"
    return f"{n/d*100:.1f}%"

def conv_cell(ws, row, col, convs, leads, fill):
    """Write 'N (X.X%)' — bold number + grey pct"""
    c = ws.cell(row=row, column=col)
    c.alignment = ac('center')
    if fill: c.fill = fill
    if convs == 0 and leads == 0:
        c.value = "—"; c.font = af(size=10, color="808080"); return
    p = f"{convs/leads*100:.1f}%" if leads else "—"
    txt = f"{convs}"
    pct_txt = f" ({p})"
    if RICH_TEXT:
        c.value = CellRichText(
            TextBlock(InlineFont(b=True,  sz=11, color="1A1A1A"), txt),
            TextBlock(InlineFont(b=False, sz=8,  color="808080"), pct_txt),
        )
    else:
        c.value = f"{txt}{pct_txt}"
        c.font = af(bold=True, size=11)

# ── Build workbook ─────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "APAC Conv Analysis"

# Row 1: Title
ws.merge_cells('A1:M1')
cell(ws, 1, 1, "APAC — 30/60/90-Day Conversion Analysis  (Jan–Jul 2024 / 2025 / 2026)",
     fill=YELLOW, bold=True, size=13)
ws.row_dimensions[1].height = 28

# Row 2: blank
ws.row_dimensions[2].height = 8

# Row 3: Section group headers
# Col A = Country (navy)
# Cols B-E = 2024 (pink): Leads | 30d | 60d | 90d
# Cols F-I = 2025 (blue)
# Cols J-M = 2026 (green)
ws.merge_cells('A3:A4'); cell(ws, 3, 1, "Country", fill=NAVY, bold=True, size=10, color="FFFFFF")
ws.merge_cells('B3:E3'); cell(ws, 3, 2, "2024  (Jan–Jul)", fill=PINK, bold=True, size=10)
ws.merge_cells('F3:I3'); cell(ws, 3, 6, "2025  (Jan–Jul)", fill=BLUE, bold=True, size=10)
ws.merge_cells('J3:M3'); cell(ws, 3, 10, "2026  (Jan–Jul)*", fill=GREEN, bold=True, size=10)
ws.row_dimensions[3].height = 22

# Row 4: Column sub-headers (skip col 1 — it's merged from A3:A4)
sub_hdrs = ["Leads", "30d Convs", "60d Convs", "90d Convs"]
fills = [PINK]*4 + [BLUE]*4 + [GREEN]*4
for i, (hdr, f) in enumerate(zip(sub_hdrs*3, fills)):
    col = i + 2
    cell(ws, 4, col, hdr, fill=f, bold=True, size=9, wrap=True)
ws.row_dimensions[4].height = 30

# Data rows
row = 5
for i, country in enumerate(COUNTRIES):
    row_fill = LGREY if i % 2 == 0 else None

    # Country name
    cell(ws, row, 1, country, fill=row_fill, bold=True, size=10, h='left')

    col = 2
    for yr in YEARS:
        d = data.get((country, yr), {'leads':0,'c30':0,'c60':0,'c90':0})
        l = d['leads']; c30 = d['c30']; c60 = d['c60']; c90 = d['c90']

        sec_fill = [PINK, BLUE, GREEN][YEARS.index(yr)]

        # Leads cell
        c = ws.cell(row=row, column=col, value=l if l else 0)
        c.font = af(bold=True, size=11, color="1A1A1A")
        c.alignment = ac('center')
        c.fill = sec_fill

        # 30d, 60d, 90d
        conv_cell(ws, row, col+1, c30, l, sec_fill)
        conv_cell(ws, row, col+2, c60, l, sec_fill)
        conv_cell(ws, row, col+3, c90, l, sec_fill)
        col += 4

    ws.row_dimensions[row].height = 26
    row += 1

# Separator row
ws.row_dimensions[row].height = 6
row += 1

# APAC Total row — navy background, white text, bold
cell(ws, row, 1, "APAC Total (all countries)", fill=NAVY, bold=True, size=10, color="FFFFFF", h='left')
col = 2
for yr in YEARS:
    d = data.get((APAC_TOTAL_KEY, yr), {'leads':0,'c30':0,'c60':0,'c90':0})
    l = d['leads']; c30 = d['c30']; c60 = d['c60']; c90 = d['c90']
    sec_fill = [PINK, BLUE, GREEN][YEARS.index(yr)]

    c = ws.cell(row=row, column=col, value=l if l else 0)
    c.font = af(bold=True, size=11, color="1A1A1A"); c.alignment = ac('center'); c.fill = sec_fill

    conv_cell(ws, row, col+1, c30, l, sec_fill)
    conv_cell(ws, row, col+2, c60, l, sec_fill)
    conv_cell(ws, row, col+3, c90, l, sec_fill)
    col += 4
ws.row_dimensions[row].height = 26
row += 1

# Footnote
ws.merge_cells(f'A{row+1}:M{row+1}')
cell(ws, row+1, 1, "* 2026 note: 90-day window is complete only for leads created before ~16 May 2026 (Jan–Apr cohort). Jun–Jul leads still have time remaining.",
     bold=False, size=8, color="808080", h='left')

# Column widths
ws.column_dimensions['A'].width = 16
for col in range(2, 14):
    ws.column_dimensions[get_column_letter(col)].width = 13

ws.freeze_panes = 'B5'

wb.save(OUT)
print(f"\n📊 Saved: {OUT}")
print(f"   Countries: {', '.join(COUNTRIES)}")
print(f"   Sheets: 'APAC Conv Analysis'")
print(f"   Columns: Leads | 30d Convs (%) | 60d Convs (%) | 90d Convs (%) × 3 years")
