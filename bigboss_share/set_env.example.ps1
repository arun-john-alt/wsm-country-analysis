# Big Boss — fill in YOUR values, then run this in PowerShell before running the pipeline:
#   . .\set_env.example.ps1      (note the leading dot — sets vars in the current session)
# These persist only for the current PowerShell window. For permanent, use [Environment]::SetEnvironmentVariable.

# --- BigQuery ---
$env:BIGBOSS_BQ_PROJECT  = "your-bq-project-id"
$env:BIGBOSS_BQ_DATASET  = "your_bq_dataset"

# --- Google Ads API ---
$env:GADS_DEV_TOKEN      = ""
$env:GADS_CLIENT_ID      = ""
$env:GADS_CLIENT_SECRET  = ""
$env:GADS_REFRESH_TOKEN  = ""
$env:GADS_LOGIN_CID      = ""   # digits only, no dashes
$env:GADS_CUSTOMER_ID    = ""   # digits only, no dashes

# --- Zoho Sheet ---
$env:ZOHO_SHEET_ID       = ""
$env:ZOHO_CLIENT_ID      = ""
$env:ZOHO_CLIENT_SECRET  = ""
$env:ZOHO_REFRESH_TOKEN  = ""

# --- optional: Zoho data center (default .in). Uncomment + adjust if .com / .eu / etc. ---
# $env:ZOHO_TOKEN_URL = "https://accounts.zoho.com/oauth/v2/token"
# $env:ZOHO_API_BASE  = "https://sheet.zoho.com/api/v2"

Write-Output "Big Boss env vars set for this session."
