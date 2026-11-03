"""Creates SAMPLE response files (fake figures) for the REAL regions and branches in demo_data/Regions_Master.xlsx.

    python make_samples.py

Writes sample_files/Sample_Region_<code>_<name>_Oct2026.xlsx - one file per region, in the SAME layout as the
Microsoft Forms export of the real form (Reporting Month / Reporting Year / Region / Region Code / Branch Name /
Branch Code, then Previous Month / Present Month / Year To Date for each of the 18 questions).
The first five regions contain deliberate mistakes so you can see how the portal reacts; the rest are clean.
The figures are random - they are NOT real data.
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

import numpy as np
import pandas as pd

from tbml import config as cfg

BASE_DIR = Path(__file__).resolve().parent

RS = " (To get the Rupee value, do the Foreign Currency conversion using mid rate of the last day of reporting period. If the aggregate value is Rs. 2,500,000, convert it to millions:2,500,000 / 1,000,000 = 2.50 ) Rs. Million"
BASE = [
 "Total number of trade related alerts generated during the reporting period.",
 "Number of trade related alerts remaining unresolved at the end of the reporting period.",
 "Number of trade-related alerts remaining unresolved for more than 30 calendar days as at the end of the reporting period.",
 "Number of trade related alerts identified as suspected TBML alerts during the reporting period.",
 "Aggregate value of trade related alerts identified as suspected TBML incidents during the reporting period." + RS,
 "Number of Suspicious Transaction Reports (STRs) submitted to Compliance Department where TBML is identified as the underlying or related suspicious activity during the reporting period. ",
 "Number of trade transactions involving sanctioned, FATF black-list/grey-list or other Bank-defined high-risk jurisdictions during the reporting period.",
 "Aggregate value of trade transactions involving sanctioned, FATF black-list/grey-list or other Bank-defined high-risk jurisdictions during the reporting period." + RS + ".",
 "Number of trade transactions (excluding LCs) processed during the reporting period for customers who do not maintain a credit relationship with the Bank.",
 "Number of trade transactions involving payment or settlement by/to third parties other than the importer/exporter or documented contractual counterparty or likewise during the reporting period.",
 "Number of customers without a credit relationship who conducted more than 10 trade transactions.",
 "Number of trade transactions processed (excluding LCs) for customers for business purpose without a valid Business Registration during the reporting period.",
 "Number of trade transactions processed exceeding the disclosed income value in the KYC form, during the reporting period.",
 "Number of trade transactions exceeding per transaction value of more than LKR 25 Mn for a single customer for a single day, same branch, settled through physical cash during the reporting period.",
 "Number of customer relationships discontinued during the reporting period following identification of suspected or confirmed TBML-related activity.",
 "Number of trade transactions conducted by customers categorized as High Risk during the reporting period.",
 "Number of customers without a credit relationship who conducted more than 3 trade transactions using the same account in a single day during the reporting period.",
 "Number of customers conducting more than 5 trade transactions using the same account involving a single country in a single day during the reporting period.",
]
SUFFIX = {"Previous": "(Previous Month)", "Present": "(Present Month)", "YTD": "(Year To Date)"}
def title(i, period):
    suf = SUFFIX[period]
    if i == 1 and period == "Previous": suf = "(Previos Month)"   # typo in the real form
    if i == 5: return f"{BASE[i]} {suf}" if False else BASE[i] + f" {suf}"
    return BASE[i] + suf


DECL = "I confirm the information is accurate, complete, validated and supported by relevant records (Circular 1996/2026, section 4)."


def main() -> None:
    rng = np.random.default_rng(2026)
    master = BASE_DIR / "demo_data" / "Regions_Master.xlsx"
    regions = pd.read_excel(master, sheet_name="Regions", dtype=str).fillna("")
    branches = pd.read_excel(master, sheet_name="Branches", dtype=str).fillna("")
    (BASE_DIR / "sample_files").mkdir(exist_ok=True)
    for old in (BASE_DIR / "sample_files").glob("*.xlsx"):
        old.unlink()
    ind = cfg.INDICATORS

    def make_row(rid, reg, br, done, month="October", numbered=False):
        row = {"ID": rid, "Start time": done - dt.timedelta(minutes=9), "Completion time": done,
               "Email": f"officer.{br['Branch Code']}@example.com", "Name": f"Officer {br['Branch Code']}",
               "Reporting Month": month, "Reporting Year": 2026,
               "Select Your Region": reg["Region Name"], "Select Your Region Code": int(reg["Region Code"]),
               "Select Your Branch Name": br["Branch Name"], "Select Your Branch Code": int(br["Branch Code"])}
        q = 7
        for i, it in enumerate(ind):
            prev = round(float(rng.uniform(0, 12)), 2) if it["kind"] == "value" else int(rng.integers(0, 12))
            pres = round(float(rng.uniform(0, 12)), 2) if it["kind"] == "value" else int(rng.integers(0, 12))
            ytd = prev + pres + (round(float(rng.uniform(0, 30)), 2) if it["kind"] == "value" else int(rng.integers(0, 30)))
            if it["ytd"] == "latest":
                ytd = pres + int(rng.integers(0, 6))
            for period, v in (("Previous", prev), ("Present", pres), ("YTD", ytd)):
                t = title(i, period)
                row[f"{q}.{t}" if numbered else t] = float(v)
                q += 1
        d = done.strftime("%d/%m/%Y")
        row.update({"Prepared By - Name": f"Officer {br['Branch Code']}", "Prepared By - Designation": "Assistant Manager",
                    "Prepared By - Date": d, "Approved By-  Name": f"BM {br['Branch Code']}",
                    "Approved By-  Designation": "Branch Manager", "Approved By-  Date": d, DECL: "Yes, I Confirm."})
        return row

    for ridx, reg in regions.iterrows():
        rc = reg["Region Code"]
        brs = branches[branches["Region Code"] == rc].reset_index(drop=True)
        data = []
        for k, br in brs.iterrows():
            if ridx == 3 and k == len(brs) - 1:
                continue  # one branch has not submitted yet
            done = dt.datetime(2026, 11, int(rng.integers(2, 5)), int(rng.integers(8, 17)), 20)
            data.append(make_row(len(data) + 1, reg, br, done, numbered=(ridx == 9)))
        df = pd.DataFrame(data)
        first_present = [c for c in df.columns if "Present Month" in c][0]
        first_prev = [c for c in df.columns if "Previous Month" in c or "Previos Month" in c][0]
        if ridx == 0:       # a response for another month (ignored) + a decimal in a count
            old = pd.DataFrame([make_row(99, reg, brs.iloc[0], dt.datetime(2026, 10, 3, 10, 0), month="September")])
            df = pd.concat([old, df], ignore_index=True)
            df.loc[df["Select Your Branch Code"] == int(brs.iloc[1]["Branch Code"]), first_present] = 12.5
        if ridx == 1:       # a branch submitted twice (earlier one has a wrong figure)
            early = df.iloc[[0]].copy()
            early["Completion time"] = early["Completion time"] - dt.timedelta(days=1)
            early[first_present] = 999.0
            df = pd.concat([early, df], ignore_index=True)
        if ridx == 2:       # unknown branch code, late submission, name that does not match the code, other sender
            df.loc[0, "Select Your Branch Code"] = 999
            df.loc[0, "Select Your Branch Name"] = "Not a real branch"
            df.loc[1, "Completion time"] = dt.datetime(2026, 11, 8, 11, 0)
            df.loc[2, "Select Your Branch Name"] = brs.iloc[0]["Branch Name"]
            df.loc[3, "Email"] = "someone.else@example.com"
        if ridx == 4:       # YTD smaller than the present month
            ytd_col = [c for c in df.columns if "Year To Date" in c][0]
            df.loc[0, ytd_col] = 0.0
            df.loc[0, first_present] = 7.0
        df["ID"] = range(1, len(df) + 1)
        safe = "".join(ch for ch in reg["Region Name"] if ch.isalnum())
        df.to_excel(BASE_DIR / "sample_files" / f"Sample_Region_{rc}_{safe}_Oct2026.xlsx", index=False)
    print("Demo files written.")


if __name__ == "__main__":
    main()
