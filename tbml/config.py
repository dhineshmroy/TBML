"""Central configuration for the TBML Reporting Portal.

The 18 indicators below are transcribed from General Instruction Circular
No. 1996/2026 (Annex I). If the circular is ever amended, this is the ONLY
place that needs to change.

YTD rules (field "ytd"):
    "sum"    -> YTD = sum of the monthly bank-wide figures in the same calendar year
    "latest" -> YTD = the present-month figure (used for point-in-time balances
                such as "unresolved alerts at period end", where adding months
                together would be meaningless)
These two rules are an ASSUMPTION - confirm with Compliance (see README).
"""
from __future__ import annotations

import json
import os
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent.parent


def _data_dir() -> Path:
    env = os.environ.get("TBML_DATA_DIR")
    if not env:
        return BASE_DIR / "data"
    p = Path(env)
    return p if p.is_absolute() else BASE_DIR / p


DATA_DIR = _data_dir()
BACKUP_DIR = DATA_DIR / "backups"
ASSETS_DIR = BASE_DIR / "assets"
TEMPLATES_DIR = BASE_DIR / "templates"
RECORDS_FILE = DATA_DIR / "TBML_Records.xlsx"
MASTER_FILE = DATA_DIR / "Regions_Master.xlsx"
SETTINGS_FILE = BASE_DIR / "settings.json"
MAX_BACKUPS = 30

FIRST_REPORT_MONTH = "2026-10"  # October 2026 report is the first (circular s.8)

DEFAULT_SETTINGS = {
    "bank_name": "People's Bank",
    "department": "Channel Management Department",
    "unit_label": "Channel Management Department - All Regions (Amalgamated)",
    "branch_due_day": 5,      # Branch -> Regional Manager (circular s.4)
    "regional_due_day": 10,   # Regional Manager -> Channel Management (circular s.4)
    "password": "",           # empty = no login screen
    "prepared_by": "",
    "prepared_designation": "",
    "approved_by": "",
    "approved_designation": "",
}


def load_settings() -> dict:
    s = dict(DEFAULT_SETTINGS)
    if SETTINGS_FILE.exists():
        try:
            s.update(json.loads(SETTINGS_FILE.read_text(encoding="utf-8")))
        except Exception:
            pass
    return s


def _ind(no, short, full, unit, kind="count", ytd="sum", group=""):
    return {
        "no": no,
        "code": f"KRI-{no:02d}",
        "short": short,
        "full": full,
        "unit": unit,
        "kind": kind,   # "count" (whole numbers) or "value" (Rs. Million, 2 decimals)
        "ytd": ytd,
        "group": group,
    }


G1 = "Alerts & STRs"
G2 = "High-risk jurisdictions"
G3 = "Customer & transaction behaviour"
G4 = "Relationships & high-risk customers"

INDICATORS = [
    _ind(1, "Alerts generated",
         "Total number of trade related alerts generated during the reporting period.",
         "No. of alerts", group=G1),
    _ind(2, "Alerts unresolved (period end)",
         "Number of trade related alerts remaining unresolved at the end of the reporting period.",
         "No. of unresolved alerts", ytd="latest", group=G1),
    _ind(3, "Alerts unresolved > 30 days",
         "Number of trade-related alerts remaining unresolved for more than 30 calendar days "
         "as at the end of the reporting period.",
         "No. of unresolved alerts", ytd="latest", group=G1),
    _ind(4, "Suspected TBML alerts",
         "Number of trade related alerts identified as suspected TBML alerts during the reporting period.",
         "No. of alerts", group=G1),
    _ind(5, "Value of suspected TBML alerts (Rs. Mn)",
         "Aggregate value of trade related alerts identified as suspected TBML incidents during the "
         "reporting period. (To get the Rupee value, do the Foreign Currency conversion using mid rate "
         "of the last day of reporting period)",
         "Rs. Million", kind="value", group=G1),
    _ind(6, "STRs submitted (TBML)",
         "Number of Suspicious Transaction Reports (STRs) submitted to Compliance Department where TBML "
         "is identified as the underlying or related suspicious activity during the reporting period.",
         "No. of STRs", group=G1),
    _ind(7, "High-risk jurisdiction transactions",
         "Number of trade transactions involving sanctioned, FATF black-list/grey-list or other "
         "Bank-defined high-risk jurisdictions during the reporting period.",
         "No. of transactions", group=G2),
    _ind(8, "Value of high-risk jurisdiction transactions (Rs. Mn)",
         "Aggregate value of trade transactions involving sanctioned, FATF black-list/grey-list or other "
         "Bank-defined high-risk jurisdictions during the reporting period. (To get the Rupee value, do "
         "the Foreign Currency conversion using mid rate of the last day of reporting period)",
         "Rs. Million", kind="value", group=G2),
    _ind(9, "Transactions - customers without credit relationship",
         "Number of trade transactions (excluding LCs) processed during the reporting period for "
         "customers who do not maintain a credit relationship with the Bank.",
         "No. of transactions", group=G3),
    _ind(10, "Third-party payment/settlement transactions",
         "Number of trade transactions involving payment or settlement by/to third parties other than "
         "the importer/exporter or documented contractual counterparty or likewise during the reporting period.",
         "No. of transactions", group=G3),
    _ind(11, "Customers without credit: > 10 transactions",
         "Number of customers without a credit relationship who conducted more than 10 trade transactions "
         "(excluding LCs) during the reporting period.",
         "No. of customers", group=G3),
    _ind(12, "Transactions - no valid Business Registration",
         "Number of trade transactions processed (excluding LCs) for customers for business purpose "
         "without a valid Business Registration during the reporting period.",
         "No. of transactions", group=G3),
    _ind(13, "Transactions exceeding KYC income",
         "Number of trade transactions processed exceeding the disclosed income value in the KYC form, "
         "during the reporting period.",
         "No. of transactions", group=G3),
    _ind(14, "Cash transactions > LKR 25 Mn",
         "Number of trade transactions exceeding per transaction value of more than LKR 25 Mn for a single "
         "customer for a single day, same branch, settled through physical cash during the reporting period. "
         "(To get the Rupee value, do the Foreign Currency conversion using mid rate of the last day of "
         "reporting period)",
         "No. of transactions", group=G3),
    _ind(15, "Relationships discontinued (TBML)",
         "Number of customer relationships discontinued during the reporting period following identification "
         "of suspected or confirmed TBML-related activity.",
         "No. of relationships", group=G4),
    _ind(16, "Transactions - High Risk customers",
         "Number of trade transactions conducted by customers categorized as High Risk during the reporting period.",
         "No. of transactions", group=G4),
    _ind(17, "Customers without credit: > 3 transactions/day",
         "Number of customers without a credit relationship who conducted more than 3 trade transactions "
         "using the same account in a single day during the reporting period.",
         "No. of customers", group=G4),
    _ind(18, "Customers: > 5 transactions/day, single country",
         "Number of customers conducting more than 5 trade transactions using the same account involving a "
         "single country in a single day during the reporting period.",
         "No. of customers", group=G4),
]

KRI_CODES = [i["code"] for i in INDICATORS]
IND_BY_CODE = {i["code"]: i for i in INDICATORS}
VALUE_CODES = [i["code"] for i in INDICATORS if i["kind"] == "value"]
COUNT_CODES = [i["code"] for i in INDICATORS if i["kind"] == "count"]

def num_col(code: str, period: str = "Present") -> str:
    """Stored column name of an indicator figure: 'KRI-01' (present), 'KRI-01 Prev', 'KRI-01 YTD'."""
    return {"Present": code, "Previous": f"{code} Prev", "YTD": f"{code} YTD"}[period]


# Every branch reports THREE figures per indicator (as in Annex I): present month, previous month, year to date.
NUM_COLS = [num_col(c, p) for c in KRI_CODES for p in ("Present", "Previous", "YTD")]
PERIODS = ("Present", "Previous", "YTD")


def base_code(col: str) -> str:
    return col[:6]


REGION_COLS = ["Region Code", "Region Name", "Officer Email", "Alternate Email"]
BRANCH_COLS = ["Region Code", "Region Name", "Branch Code", "Branch Name", "Branch Email"]

STATUS_SUBMITTED = "Submitted"
STATUS_NIL = "Nil Report"
STATUS_NONE = "Not Submitted"
