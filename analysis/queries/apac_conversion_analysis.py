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
PINK    = PatternFill("solid", fgColor="FFB6C1")   # 2024
BLUE    = PatternFill("solid", fgColor="D9E8F5")   # 2025
GREEN   = PatternFill("solid", fgColor="90EE90")   # 2026
LGREY   = PatternFill("solid", fgColor="F2F2F2")

YEAR_FILLS = {'2024': PINK, '2025': BLUE, '2026': GREEN}
YEAR_LABELS = {'2024': '2024  (Jan–Jul)', '2025': '2025  (Jan–Jul)', '2026': '2026  (Jan–Jul)*'}

# Columns: COUNTRIES + APAC Total
ALL_COLS = COUNTRIES + ['APAC Total']
# Row key for APAC Total
def get_d(country, yr):
    key = APAC_TOTAL_KEY if country == 'APAC Total' else country
    return data.get((key, yr), {'leads':0,'c30':0,'c60':0,'c90':0})

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

def conv_cell(ws, row, col, convs, leads, fill):
    """Write bold N + grey (X.X%) in same cell."""
    c = ws.cell(row=row, column=col)
    c.alignment = ac('center')
    if fill: c.fill = fill
    if convs == 0 and leads == 0:
        c.value = "—"; c.font = af(size=10, color="808080"); return
    p = f"{convs/leads*100:.1f}%" if leads else "—"
    if RICH_TEXT:
        c.value = CellRichText(
            TextBlock(InlineFont(b=True,  sz=11, color="1A1A1A"), str(convs)),
            TextBlock(InlineFont(b=False, sz=8,  color="808080"), f" ({p})"),
        )
    else:
        c.value = f"{convs} ({p})"; c.font = af(bold=True, size=11)

# ── Build workbook ─────────────────────────────────────────────────────────────
# Layout (transposed):
#   Rows = Year × Metric (Leads, 30d, 60d, 90d)
#   Cols = Country (Vietnam, Singapore, Philippines, Indonesia, Malaysia, Thailand, APAC Total)
#
# Col 1 = Year label (merged across 4 metric rows)
# Col 2 = Metric label (Leads / 30d Convs / 60d Convs / 90d Convs)
# Col 3+ = one col per country/region

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "APAC Conv Analysis"

N_COUNTRIES = len(ALL_COLS)   # 7
LAST_COL    = 2 + N_COUNTRIES  # col 9

# Row 1: Title
ws.merge_cells(f'A1:{get_column_letter(LAST_COL)}1')
cell(ws, 1, 1, "APAC — 30/60/90-Day Conversion Analysis  (Jan–Jul 2024 / 2025 / 2026)",
     fill=YELLOW, bold=True, size=13)
ws.row_dimensions[1].height = 28

# Row 2: blank spacer
ws.row_dimensions[2].height = 8

# Row 3: Column headers
cell(ws, 3, 1, "Year",   fill=NAVY, bold=True, size=10, color="FFFFFF")
cell(ws, 3, 2, "Metric", fill=NAVY, bold=True, size=10, color="FFFFFF")
for ci, country in enumerate(ALL_COLS):
    is_total = country == 'APAC Total'
    f = NAVY if is_total else LGREY
    txt_color = "FFFFFF" if is_total else "000000"
    cell(ws, 3, 3 + ci, country, fill=f, bold=True, size=10, color=txt_color, wrap=True)
ws.row_dimensions[3].height = 30

# Data rows: 3 years × 4 metrics = 12 data rows + 2 separator rows
METRICS = [
    ('Leads',     'leads'),
    ('30d Convs', 'c30'),
    ('60d Convs', 'c60'),
    ('90d Convs', 'c90'),
]

data_row = 4
for yi, yr in enumerate(YEARS):
    yr_fill = YEAR_FILLS[yr]
    yr_label = YEAR_LABELS[yr]

    # Merge year label across 4 metric rows
    yr_start = data_row
    yr_end   = data_row + len(METRICS) - 1
    ws.merge_cells(f'A{yr_start}:A{yr_end}')
    cell(ws, yr_start, 1, yr_label, fill=yr_fill, bold=True, size=11)

    for mi, (metric_label, metric_key) in enumerate(METRICS):
        row = data_row + mi
        # Metric label
        cell(ws, row, 2, metric_label, fill=yr_fill, bold=False, size=10)

        # Data cells per country
        for ci, country in enumerate(ALL_COLS):
            d = get_d(country, yr)
            l = d['leads']
            is_total = country == 'APAC Total'
            col_fill = NAVY if is_total else yr_fill

            if metric_key == 'leads':
                c = ws.cell(row=row, column=3+ci, value=l if l else 0)
                c.font = af(bold=True, size=11, color="FFFFFF" if is_total else "1A1A1A")
                c.alignment = ac('center'); c.fill = col_fill
            else:
                v = d[metric_key]
                conv_cell(ws, row, 3+ci, v, l, col_fill)

        ws.row_dimensions[row].height = 24

    data_row += len(METRICS)

    # Separator between years
    if yi < len(YEARS) - 1:
        ws.row_dimensions[data_row].height = 6
        data_row += 1

# Footnote
fn_row = data_row + 1
ws.merge_cells(f'A{fn_row}:{get_column_letter(LAST_COL)}{fn_row}')
cell(ws, fn_row, 1,
     "* 2026 note: 90-day window complete only for leads created before ~16 May 2026 (Jan–Apr). Jun–Jul leads still have time remaining.",
     bold=False, size=8, color="808080", h='left')

# Column widths
ws.column_dimensions['A'].width = 20  # Year
ws.column_dimensions['B'].width = 13  # Metric
for ci in range(N_COUNTRIES):
    col_letter = get_column_letter(3 + ci)
    ws.column_dimensions[col_letter].width = 16 if ALL_COLS[ci] == 'APAC Total' else 14

ws.freeze_panes = 'C4'

wb.save(OUT)
print(f"\n📊 Saved: {OUT}")
print(f"   Layout: Years as rows, Countries as columns")
print(f"   Countries: {', '.join(ALL_COLS)}")
print(f"   Metrics per year: Leads | 30d Convs (%) | 60d Convs (%) | 90d Convs (%)")
