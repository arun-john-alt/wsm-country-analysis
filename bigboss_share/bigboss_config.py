"""
Big Boss — single config file. THIS IS THE ONLY FILE YOU CONFIGURE.

Credentials and the BigQuery project/dataset are NOT bundled — supply your own via environment
variables (see README.md). If any required variable is missing, importing this module fails with
a list of exactly what to set, so nothing runs against the wrong account by accident.

Paths auto-resolve and rarely need touching:
  - HOME      = the folder these scripts live in.
  - ADC path  = your %APPDATA%/gcloud location (run: gcloud auth application-default login).
  - DOWNLOADS = your user Downloads folder (where change-history CSVs are dropped).
"""
import os

_MISSING = []
def _req(name):
    """Required env var; record (don't crash yet) so we can report all missing at once."""
    v = os.environ.get(name)
    if not v:
        _MISSING.append(name)
    return v or ""

def _appdata_adc():
    base = os.environ.get("APPDATA") or os.path.expanduser(r"~/AppData/Roaming")
    return os.path.join(base, "gcloud", "application_default_credentials.json")

# --- paths (auto-resolve; rarely need editing) ---
HOME      = os.environ.get("BIGBOSS_HOME") or os.path.dirname(os.path.abspath(__file__))
DOWNLOADS = os.environ.get("BIGBOSS_DOWNLOADS") or os.path.join(os.path.expanduser("~"), "Downloads")
ADC_PATH  = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS") or _appdata_adc()
os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = ADC_PATH

# --- BigQuery (YOUR project / dataset) ---
BQ_PROJECT      = _req("BIGBOSS_BQ_PROJECT")
BQ_DATASET      = _req("BIGBOSS_BQ_DATASET")
QUOTA_PROJECT   = os.environ.get("BIGBOSS_QUOTA_PROJECT") or BQ_PROJECT
DEFAULT_DATASET = f"{BQ_PROJECT}.{BQ_DATASET}"

# --- Google Ads API (YOUR credentials; read-only KP usage) ---
GOOGLE_ADS = {
    "developer_token":   _req("GADS_DEV_TOKEN"),
    "client_id":         _req("GADS_CLIENT_ID"),
    "client_secret":     _req("GADS_CLIENT_SECRET"),
    "refresh_token":     _req("GADS_REFRESH_TOKEN"),
    "use_proto_plus":    True,
    "login_customer_id": _req("GADS_LOGIN_CID"),
}
GOOGLE_ADS_CUSTOMER_ID = _req("GADS_CUSTOMER_ID")

# --- Zoho Sheet (YOUR sheet + OAuth app) ---
ZOHO_SHEET_ID      = _req("ZOHO_SHEET_ID")
ZOHO_CLIENT_ID     = _req("ZOHO_CLIENT_ID")
ZOHO_CLIENT_SECRET = _req("ZOHO_CLIENT_SECRET")
ZOHO_REFRESH_TOKEN = _req("ZOHO_REFRESH_TOKEN")
# data-center URLs: default to .in; override if your Zoho account is .com / .eu / .com.au etc.
ZOHO_TOKEN_URL     = os.environ.get("ZOHO_TOKEN_URL", "https://accounts.zoho.in/oauth/v2/token")
ZOHO_API_BASE      = os.environ.get("ZOHO_API_BASE", "https://sheet.zoho.in/api/v2")

if _MISSING:
    raise RuntimeError(
        "Big Boss config: set these environment variables before running (see README.md):\n  "
        + "\n  ".join(_MISSING)
    )

# --- helpers ---
def bq_client():
    """Configured BigQuery client (project + quota project)."""
    import google.auth
    from google.cloud import bigquery
    cred, _ = google.auth.default()
    if hasattr(cred, "with_quota_project"):
        cred = cred.with_quota_project(QUOTA_PROJECT)
    return bigquery.Client(project=BQ_PROJECT, credentials=cred)

def ads_client():
    """Configured Google Ads API client."""
    from google.ads.googleads.client import GoogleAdsClient
    return GoogleAdsClient.load_from_dict(GOOGLE_ADS)
