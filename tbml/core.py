"""Core data logic for the TBML Reporting Portal (no Streamlit code in here).

Data model (branch level):
  * Master data: Regions (RHO) and Branches (each branch belongs to one region).
  * Each Microsoft Form belongs to ONE region; branches pick their own branch.
  * Stored records: TBML_Records.xlsx, one tab per month, ONE ROW PER BRANCH.
  * Region and bank-wide figures are always calculated from the branch rows.
"""
from __future__ import annotations

import datetime as dt
import difflib
import os
import re
import shutil
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from . import config as cfg
from .config import INDICATORS, KRI_CODES, IND_BY_CODE, NUM_COLS, num_col, base_code

# --------------------------------------------------------------------------- #
# Constants / small helpers
# --------------------------------------------------------------------------- #
FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F3A5F")
HDR_FONT = Font(name=FONT, bold=True, color="FFFFFF", size=10)
BODY_FONT = Font(name=FONT, size=10)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

TEXT_COLS = ["Region Code", "Region Name", "Branch Code", "Branch Name", "Status", "Responder Email",
             "Submitted At", "Prepared By", "Prepared Designation", "Prepared Date",
             "Approved By", "Approved Designation", "Approved Date", "Edited",
             "RHO Confirmed By", "Source File"]
MONTH_COLS = TEXT_COLS + NUM_COLS
RESERVED_SHEETS = {"Legend", "Log"}
LOG_COLS = ["Timestamp", "Reporting Month", "Source File", "Mode", "Rows in File",
            "Branches Saved", "Cells Edited Manually", "Notes"]


class FileLockedError(Exception):
    """Raised when an Excel file is open in Excel and cannot be overwritten."""


def month_label(key: str) -> str:
    return dt.datetime.strptime(key, "%Y-%m").strftime("%B %Y")


def key_to_tab(key: str) -> str:
    return dt.datetime.strptime(key, "%Y-%m").strftime("%b-%Y")


def tab_to_key(tab: str) -> str | None:
    try:
        return dt.datetime.strptime(tab.strip(), "%b-%Y").strftime("%Y-%m")
    except ValueError:
        return None


def prev_key(key: str) -> str:
    y, m = int(key[:4]), int(key[5:7])
    y, m = (y - 1, 12) if m == 1 else (y, m - 1)
    return f"{y:04d}-{m:02d}"


def next_month_start(key: str) -> dt.datetime:
    y, m = int(key[:4]), int(key[5:7])
    y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return dt.datetime(y, m, 1)


def due_date(key: str, day: int) -> dt.datetime:
    """End of the given day of the month FOLLOWING the reporting month."""
    return next_month_start(key).replace(day=day, hour=23, minute=59, second=59)


def ordinal(n: int) -> str:
    suffix = "th" if 10 <= n % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def month_options(n_after_first: int = 60) -> list[str]:
    y, m = int(cfg.FIRST_REPORT_MONTH[:4]), int(cfg.FIRST_REPORT_MONTH[5:7])
    out = []
    for _ in range(n_after_first):
        out.append(f"{y:04d}-{m:02d}")
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def fmt_val(x, kind: str = "count") -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "N/A"
    if kind == "value":
        return f"{x:,.2f}"
    return f"{x:,.0f}"


def branch_label(code: str, name: str) -> str:
    return f"{code} | {name}"


def find_logo() -> Path | None:
    for name in ("logo.png", "logo.jpg", "logo.jpeg"):
        p = cfg.ASSETS_DIR / name
        if p.exists():
            return p
    return None


# --------------------------------------------------------------------------- #
# File safety: backups + atomic save
# --------------------------------------------------------------------------- #
def _backup(path: Path) -> None:
    if not path.exists():
        return
    cfg.BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    shutil.copy2(path, cfg.BACKUP_DIR / f"{path.stem}_{stamp}{path.suffix}")
    files = sorted(cfg.BACKUP_DIR.glob(f"{path.stem}_*{path.suffix}"))
    for old in files[: -cfg.MAX_BACKUPS]:
        try:
            old.unlink()
        except OSError:
            pass


def _replace(tmp: Path, target: Path) -> None:
    try:
        os.replace(tmp, target)
    except PermissionError as exc:
        try:
            tmp.unlink()
        except OSError:
            pass
        raise FileLockedError(f"'{target.name}' is open in Excel. Close it and try again.") from exc


def _style_sheet(ws, header_row: int = 1, widths: dict | None = None, freeze: str | None = "C2") -> None:
    for cell in ws[header_row]:
        cell.fill, cell.font, cell.border = HDR_FILL, HDR_FONT, BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[header_row].height = 32
    for row in ws.iter_rows(min_row=header_row + 1):
        for cell in row:
            cell.font = BODY_FONT
            cell.border = BORDER
    for i, col in enumerate(ws.iter_cols(min_row=1, max_row=1), start=1):
        header = str(col[0].value or "")
        ws.column_dimensions[get_column_letter(i)].width = (widths or {}).get(header, max(12, min(34, len(header) + 4)))
    if freeze:
        ws.freeze_panes = freeze


# --------------------------------------------------------------------------- #
# Master data (regions / branches)
# --------------------------------------------------------------------------- #
def _clean(df: pd.DataFrame, cols: list[str], key_cols: list[str] | None = None) -> pd.DataFrame:
    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols].fillna("").astype(str)
    for c in cols:
        df[c] = df[c].str.strip()
    key_cols = key_cols or cols[:2]
    blank = pd.Series(True, index=df.index)
    for k in key_cols:
        blank &= df[k].eq("")
    return df[~blank].reset_index(drop=True)


def clean_regions(df):
    return _clean(df, cfg.REGION_COLS, ["Region Code", "Region Name"])


def clean_branches(df):
    return _clean(df, cfg.BRANCH_COLS, ["Region Code", "Branch Code", "Branch Name"])


def load_master() -> tuple[pd.DataFrame, pd.DataFrame]:
    regions = pd.DataFrame(columns=cfg.REGION_COLS)
    branches = pd.DataFrame(columns=cfg.BRANCH_COLS)
    if cfg.MASTER_FILE.exists():
        xl = pd.ExcelFile(cfg.MASTER_FILE)
        if "Regions" in xl.sheet_names:
            regions = pd.read_excel(xl, sheet_name="Regions", dtype=str)
        if "Branches" in xl.sheet_names:
            branches = pd.read_excel(xl, sheet_name="Branches", dtype=str)
    regions, branches = clean_regions(regions), clean_branches(branches)
    if not branches.empty:  # Region Name on a branch row always follows the Regions list
        names = dict(zip(regions["Region Code"], regions["Region Name"]))
        branches["Region Name"] = branches["Region Code"].map(names).fillna(branches["Region Name"])
    return regions, branches


_EMAIL_RE = r"^[^@\s]+@[^@\s]+\.[^@\s]+$"


def validate_master(regions: pd.DataFrame, branches: pd.DataFrame) -> list[str]:
    errs: list[str] = []
    r, b = clean_regions(regions), clean_branches(branches)
    if r.empty:
        errs.append("The Regions list is empty.")
    if (r["Region Code"] == "").any() or (r["Region Name"] == "").any():
        errs.append("Every region needs both a Region Code and a Region Name.")
    dup = r["Region Code"][r["Region Code"].duplicated() & (r["Region Code"] != "")]
    if len(dup):
        errs.append("Duplicate Region Code(s): " + ", ".join(sorted(set(dup))))
    for col in ("Officer Email", "Alternate Email"):
        bad = r[(r[col] != "") & ~r[col].str.contains(_EMAIL_RE, regex=True)]
        if len(bad):
            errs.append(f"Invalid {col} for: " + ", ".join(bad["Region Code"]))
    if not b.empty:
        if (b["Branch Code"] == "").any() or (b["Branch Name"] == "").any():
            errs.append("Every branch needs a Branch Code and a Branch Name.")
        unknown = set(b["Region Code"]) - set(r["Region Code"])
        if unknown:
            errs.append("Branches refer to unknown Region Code(s): " + ", ".join(sorted(unknown)))
        dupb = b["Branch Code"][b["Branch Code"].duplicated() & (b["Branch Code"] != "")]
        if len(dupb):
            errs.append("Duplicate Branch Code(s): " + ", ".join(sorted(set(dupb))))
        bad = b[(b["Branch Email"] != "") & ~b["Branch Email"].str.contains(_EMAIL_RE, regex=True)]
        if len(bad):
            errs.append("Invalid Branch Email for: " + ", ".join(bad["Branch Code"]))
    return errs


def save_master(regions: pd.DataFrame, branches: pd.DataFrame) -> None:
    cfg.DATA_DIR.mkdir(parents=True, exist_ok=True)
    r, b = clean_regions(regions), clean_branches(branches)
    names = dict(zip(r["Region Code"], r["Region Name"]))
    if not b.empty:
        b["Region Name"] = b["Region Code"].map(names).fillna(b["Region Name"])
    _backup(cfg.MASTER_FILE)
    tmp = cfg.MASTER_FILE.with_name(cfg.MASTER_FILE.stem + ".tmp.xlsx")
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        r.to_excel(xw, sheet_name="Regions", index=False)
        b.to_excel(xw, sheet_name="Branches", index=False)
    wb = load_workbook(tmp)
    for ws in wb.worksheets:
        _style_sheet(ws, widths={"Officer Email": 34, "Alternate Email": 34, "Region Name": 30,
                                 "Branch Name": 30, "Branch Email": 34})
    wb.save(tmp)
    _replace(tmp, cfg.MASTER_FILE)


def branch_lists_workbook(regions: pd.DataFrame, branches: pd.DataFrame) -> bytes:
    """One sheet per region listing 'code | name' - paste into that region's Form dropdown."""
    import io
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        idx = []
        for _, r in regions.iterrows():
            sub = branches[branches["Region Code"] == r["Region Code"]]
            idx.append({"Region Code": r["Region Code"], "Region Name": r["Region Name"],
                        "RHO Email": r["Officer Email"], "Branches": len(sub),
                        "Form name": f"TBML KRI - {r['Region Code']} {r['Region Name']}"})
        pd.DataFrame(idx).to_excel(xw, sheet_name="Index", index=False)
        for _, r in regions.iterrows():
            sub = branches[branches["Region Code"] == r["Region Code"]]
            pd.DataFrame({"Branch Name": list(sub["Branch Name"]), "Branch Code": list(sub["Branch Code"])}
                         ).to_excel(xw, sheet_name=r["Region Code"][:31], index=False)
    buf.seek(0)
    wb = load_workbook(buf)
    for ws in wb.worksheets:
        _style_sheet(ws, widths={"Branch Name": 44, "Branch Code": 16, "Form name": 44, "RHO Email": 34, "Region Name": 30},
                     freeze="A2")
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# --------------------------------------------------------------------------- #
# Parsing + validating Microsoft Forms exports (one export per region Form)
# --------------------------------------------------------------------------- #
KRI_RE = re.compile(r"^\s*KRI[\s_-]*0*(\d{1,2})\b", re.I)

STAGE_META = ["Include", "Source File", "Source ID", "Region Code", "Region Name", "Branch Code", "Branch Name",
              "Status", "Responder Email", "Submitted At", "Prepared By", "Prepared Designation", "Prepared Date",
              "Approved By", "Approved Designation", "Approved Date"]
STAGE_COLS = STAGE_META + NUM_COLS + ["Check", "Issues"]

_MONTHS = {m.lower(): i for i, m in enumerate(
    ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
     "November", "December"], start=1)}
_MONTHS.update({k[:3]: v for k, v in list(_MONTHS.items())})


def parse_month(v, year=None) -> str | None:
    """'October 2026' / 'Oct-2026' / 2026-10  - or a month name/number plus a separate year."""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None
    if isinstance(v, (pd.Timestamp, dt.datetime, dt.date)):
        return f"{v.year:04d}-{v.month:02d}"
    s = str(v).strip()
    y = None
    if year is not None and not (isinstance(year, float) and np.isnan(year)) and str(year).strip() != "":
        try:
            y = int(float(str(year).strip()))
        except ValueError:
            y = None
    if y is not None:
        key = s.lower().rstrip(".")
        if key in _MONTHS:
            return f"{y:04d}-{_MONTHS[key]:02d}"
        if key.isdigit() and 1 <= int(key) <= 12:
            return f"{y:04d}-{int(key):02d}"
    for fmt in ("%B %Y", "%b %Y", "%b-%Y", "%B-%Y", "%Y-%m", "%m/%Y"):
        try:
            return dt.datetime.strptime(s, fmt).strftime("%Y-%m")
        except ValueError:
            continue
    try:
        d = pd.to_datetime(s)
        return f"{d.year:04d}-{d.month:02d}"
    except Exception:
        return None


def _num(v):
    """-> (value or None, flag) where flag is None | 'blank' | 'na' | 'bad'"""
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return None, "blank"
    s = str(v).strip().replace(",", "")
    if s == "":
        return None, "blank"
    if s.lower() in ("n/a", "na", "nil", "-"):
        return None, "na"
    try:
        return float(s), None
    except ValueError:
        return None, "bad"


def _ckey(v) -> str:
    """Code comparison key: trimmed, lower-case, leading zeros ignored ('006' == '6')."""
    s = str(v if v is not None else "").strip().lower()
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return (s.lstrip("0") or "0") if s.isdigit() else s


def _tkey(v) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(v or "").lower())


def _branch_lookup(branches: pd.DataFrame):
    by_code = {_ckey(r["Branch Code"]): r for _, r in branches.iterrows()}
    nm = branches["Branch Name"].map(_tkey)
    uniq = nm[~nm.duplicated(keep=False)]
    by_name = {nm[i]: branches.loc[i] for i in uniq.index}
    return by_code, by_name


def _resolve_branch(code, name, lookup):
    """-> (branch row or None, how) where how in 'code' | 'name' | None."""
    by_code, by_name = lookup
    c, n = str(code or "").strip(), str(name or "").strip()
    m = re.match(r"^\s*([^|]+?)\s*\|", n)           # old style 'code | name' in one answer
    if m and _ckey(m.group(1)) in by_code:
        return by_code[_ckey(m.group(1))], "code"
    if c and _ckey(c) in by_code:
        return by_code[_ckey(c)], "code"
    if n and _tkey(n) in by_name:
        return by_name[_tkey(n)], "name"
    return None, None


# --- recognising the 54 indicator columns (3 per indicator) --------------------------------------
_PERIOD_PAREN = re.compile(r"\([^)]*(previo|present|current|year to date|ytd)[^)]*\)", re.I)


def _period_of(text: str) -> str | None:
    t = str(text).lower()
    if "year to date" in t or re.search(r"\bytd\b", t):
        return "YTD"
    if re.search(r"previo", t):          # also catches the typo 'Previos'
        return "Previous"
    if re.search(r"\b(present|current)\b", t):
        return "Present"
    return None


def _norm_text(s) -> str:
    s = str(s).lower()
    s = re.sub(r"^\s*\d+\s*[.)]\s*", "", s)         # a leading question number such as '7.'
    s = _PERIOD_PAREN.sub(" ", s)
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


_IND_NORM = {i["code"]: _norm_text(i["full"]) for i in INDICATORS}


def _match_indicator(col) -> str | None:
    t = _norm_text(col)
    if not t:
        return None
    scored = sorted(((difflib.SequenceMatcher(None, t, v, autojunk=False).ratio(), code)
                     for code, v in _IND_NORM.items()), reverse=True)
    if scored[0][0] >= 0.6 and scored[0][0] - scored[1][0] >= 0.04:
        return scored[0][1]
    return None


def _map_value_columns(cols: list[str]) -> dict:
    """{(code, period): column}. Uses 'KRI-01 | ...' titles if present, otherwise the question wording."""
    mapping: dict = {}
    coded = [c for c in cols if KRI_RE.match(c)]
    if coded:
        for c in coded:
            code = f"KRI-{int(KRI_RE.match(c).group(1)):02d}"
            if code in IND_BY_CODE:
                mapping.setdefault((code, _period_of(c) or "Present"), c)
        return mapping
    dupes = []
    for c in cols:
        per = _period_of(c)
        if per is None:
            continue
        code = _match_indicator(c)
        if code is None:
            continue
        if (code, per) in mapping:
            dupes.append(f"{code} {per}")
        else:
            mapping[(code, per)] = c
    if dupes:
        raise ValueError("Two columns look like the same question: " + ", ".join(dupes[:6]))
    return mapping


def _find_columns(columns) -> dict:
    cols = [str(c) for c in columns]
    values = _map_value_columns(cols)
    used = set(values.values())
    others = [c for c in cols if c not in used]

    def pick(pred):
        for c in others:
            if pred(c.lower().strip()):
                return c
        return None

    return {
        "values": values,
        "id": pick(lambda l: l == "id"),
        "completed": pick(lambda l: l.startswith("completion")),
        "email": pick(lambda l: l.startswith("email")),
        "month": pick(lambda l: "reporting month" in l or l == "month"),
        "year": pick(lambda l: "reporting year" in l or l == "year"),
        "region": pick(lambda l: "region" in l and "code" not in l and "regional" not in l),
        "region_code": pick(lambda l: "region" in l and "code" in l),
        "branch_name": pick(lambda l: "branch" in l and "code" not in l and "approved" not in l and "prepared" not in l),
        "branch_code": pick(lambda l: "branch" in l and "code" in l),
        "has": pick(lambda l: "indicators to report" in l or "any tbml" in l),
        "prep_name": pick(lambda l: "prepared" in l and "designation" not in l and "date" not in l),
        "prep_des": pick(lambda l: "prepared" in l and "designation" in l),
        "prep_date": pick(lambda l: "prepared" in l and "date" in l),
        "appr_name": pick(lambda l: "approved" in l and "designation" not in l and "date" not in l),
        "appr_des": pick(lambda l: "approved" in l and "designation" in l),
        "appr_date": pick(lambda l: "approved" in l and "date" in l),
    }


def _get(row, col):
    if col is None:
        return ""
    v = row.get(col, "")
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return ""
    if isinstance(v, (pd.Timestamp, dt.datetime)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def validate_kri_values(df: pd.DataFrame) -> list[str]:
    """Re-check numeric rules on (possibly hand-edited) rows. Returns messages."""
    msgs = []
    for _, r in df.iterrows():
        who = r.get("Branch Code", "?") or r.get("Region Code", "?")
        for col in NUM_COLS:
            if col not in r.index:
                continue
            v, flag = _num(r.get(col))
            if flag == "bad":
                msgs.append(f"{who} {col}: '{r.get(col)}' is not a number")
            elif v is not None:
                if v < 0:
                    msgs.append(f"{who} {col}: negative value")
                elif IND_BY_CODE[base_code(col)]["kind"] == "count" and abs(v - round(v)) > 1e-9:
                    msgs.append(f"{who} {col}: must be a whole number")
    return msgs


def _parse_one(raw: pd.DataFrame, regions, branches, month_key, source, bdue, rdue, prior=None):
    found = _find_columns(raw.columns)
    vals_map = found["values"]
    missing = [f"{c} (Present)" for c in KRI_CODES if (c, "Present") not in vals_map]
    # Previous / YTD are worked out by the portal; any such columns in an old-style file are ignored.
    vals_map = {k: v for k, v in vals_map.items() if k[1] == "Present"}
    has_extra = False
    if missing:
        raise ValueError(f"{source}: could not find these indicator columns: " + ", ".join(missing[:8])
                         + (" ..." if len(missing) > 8 else "")
                         + ". Was a question title edited or removed in the Form?")
    if found["branch_code"] is None and found["branch_name"] is None:
        raise ValueError(f"{source}: no 'Branch Name' / 'Branch Code' column was found.")
    if found["month"] is None:
        raise ValueError(f"{source}: no 'Reporting Month' column was found.")

    lookup = _branch_lookup(branches)
    reg_names = dict(zip(regions["Region Code"], regions["Region Name"]))
    d_branch, d_region = due_date(month_key, bdue), due_date(month_key, rdue)
    prior_idx = None
    if prior is not None and not prior.empty:
        prior_idx = prior[prior["Status"] != cfg.STATUS_NONE].drop_duplicates("Branch Code").set_index("Branch Code")

    rows, skipped = [], 0
    for idx, rr in raw.iterrows():
        row = rr.to_dict()
        mk = parse_month(row.get(found["month"]), row.get(found["year"]) if found["year"] else None)
        if mk is not None and mk != month_key:
            skipped += 1          # a response for another month: not an error, simply not part of this upload
            continue
        issues: list[tuple[str, str]] = []
        if mk is None:
            issues.append(("Error", "Reporting month/year is missing or unreadable"))

        in_code = _get(row, found["branch_code"])
        in_name = _get(row, found["branch_name"])
        br, how = _resolve_branch(in_code, in_name, lookup)
        if br is None:
            issues.append(("Error", f"Branch '{in_name or '?'}' (code '{in_code or '?'}') not found in the Branches master"))
        else:
            if how == "name":
                issues.append(("Warning", f"Branch code '{in_code}' not recognised - matched by name"))
            if in_name and in_code and _tkey(in_name) != _tkey(br["Branch Name"]) and "|" not in in_name:
                issues.append(("Warning", f"Branch name '{in_name}' does not match code {in_code} "
                                          f"('{br['Branch Name']}' in the master)"))
            rc_in = _get(row, found["region_code"])
            if rc_in and _ckey(rc_in) != _ckey(br["Region Code"]):
                issues.append(("Warning", f"Region code '{rc_in}' differs from the branch's region "
                                          f"({br['Region Code']}) in the master"))
        email = _get(row, found["email"])
        if br is not None and br["Branch Email"] and email.lower() != br["Branch Email"].lower():
            issues.append(("Warning", f"Sender '{email or 'blank'}' differs from the registered branch e-mail"))

        submitted = pd.to_datetime(row.get(found["completed"]), errors="coerce") if found["completed"] else pd.NaT
        if pd.notna(submitted):
            if submitted > d_region:
                issues.append(("Warning", f"Submitted after the regional deadline ({ordinal(rdue)})"))
            elif submitted > d_branch:
                issues.append(("Warning", f"Submitted after the branch due date ({ordinal(bdue)})"))

        nil = _get(row, found["has"]).lower().startswith("no") if found["has"] else False
        status = cfg.STATUS_NIL if nil else cfg.STATUS_SUBMITTED

        vals: dict = {}
        blanks = []
        for (code, period), col in vals_map.items():
            name = num_col(code, period)
            v, flag = _num(row.get(col))
            kind = IND_BY_CODE[code]["kind"]
            if flag == "bad":
                issues.append(("Error", f"{name} is not a number ('{_get(row, col)}')"))
            elif v is not None and v < 0:
                issues.append(("Error", f"{name} is negative"))
                v = None
            elif v is not None and kind == "count" and abs(v - round(v)) > 1e-9:
                issues.append(("Error", f"{name} must be a whole number (got {v})"))
            elif v is not None and kind == "value":
                v = round(v, 2)
            elif v is not None:
                v = float(round(v))
            if not nil and flag in ("blank", "na"):
                blanks.append(name)
            vals[name] = v
        if blanks:
            issues.append(("Warning", f"{len(blanks)} figure(s) left blank/N/A - counted as 0: "
                                      + ", ".join(blanks[:5]) + (" ..." if len(blanks) > 5 else "")))

        # --- consistency checks between the three figures --------------------------------------
        if has_extra and not nil:
            low = [c for c in KRI_CODES if IND_BY_CODE[c]["ytd"] == "sum"
                   and vals.get(c) is not None and vals.get(num_col(c, "YTD")) is not None
                   and vals[num_col(c, "YTD")] + 1e-9 < vals[c]]
            if low:
                issues.append(("Warning", "YTD is smaller than the present month for: " + ", ".join(low[:6])))
            if prior_idx is not None and br is not None and br["Branch Code"] in prior_idx.index:
                pr = prior_idx.loc[br["Branch Code"]]
                diff = [c for c in KRI_CODES if vals.get(num_col(c, "Previous")) is not None
                        and pd.notna(pr[c]) and abs(vals[num_col(c, "Previous")] - float(pr[c])) > 1e-6]
                if diff:
                    issues.append(("Warning", "'Previous Month' differs from what this branch reported as "
                                              "present last month for: " + ", ".join(diff[:6])))

        level = "Error" if any(l == "Error" for l, _ in issues) else ("Warning" if issues else "OK")
        rec = {
            "Include": level in ("OK", "Warning"),
            "Source File": source,
            "Source ID": _get(row, found["id"]) or str(idx + 1),
            "Region Code": br["Region Code"] if br is not None else "",
            "Region Name": reg_names.get(br["Region Code"], br["Region Name"]) if br is not None else "",
            "Branch Code": br["Branch Code"] if br is not None else "",
            "Branch Name": br["Branch Name"] if br is not None else in_name,
            "Status": status, "Responder Email": email,
            "Submitted At": submitted.strftime("%Y-%m-%d %H:%M") if pd.notna(submitted) else "",
            "Prepared By": _get(row, found["prep_name"]), "Prepared Designation": _get(row, found["prep_des"]),
            "Prepared Date": _get(row, found["prep_date"]), "Approved By": _get(row, found["appr_name"]),
            "Approved Designation": _get(row, found["appr_des"]), "Approved Date": _get(row, found["appr_date"]),
            "_ts": submitted, "_level": level, "_issues": issues,
        }
        rec.update(vals)
        rows.append(rec)
    df = pd.DataFrame(rows)
    for c in NUM_COLS:
        if len(df) and c not in df.columns:
            df[c] = np.nan
    return df, skipped


def parse_many(files: list[tuple[str, pd.DataFrame]], regions: pd.DataFrame, branches: pd.DataFrame,
               month_key: str, branch_due_day: int = 5, regional_due_day: int = 10, prior: pd.DataFrame | None = None):
    """Parse one or more Forms exports into ONE reviewable staging table.

    `prior` = last month's saved rows (optional) - used to cross-check 'Previous Month' figures.
    Returns (stage, per_file_summary). Raises ValueError for structural problems."""
    frames, summary = [], []
    for name, raw in files:
        df, skipped = _parse_one(raw, regions, branches, month_key, name, branch_due_day, regional_due_day, prior)
        frames.append(df)
        summary.append({"File": name, "Rows in file": len(raw), "For this month": len(df),
                        "Other months (ignored)": skipped})
    frames = [f for f in frames if not f.empty]
    if not frames:
        return pd.DataFrame(columns=STAGE_COLS), pd.DataFrame(summary)
    stage = pd.concat(frames, ignore_index=True)

    # Several submissions for one branch -> only the latest stays ticked
    stage["_ord"] = pd.to_datetime(stage["_ts"], errors="coerce")
    for code, grp in stage[stage["Branch Code"] != ""].groupby("Branch Code"):
        if len(grp) > 1:
            latest = grp.sort_values("_ord", na_position="first", kind="stable").index[-1]
            for i in grp.index:
                if i != latest:
                    stage.at[i, "Include"] = False
                    stage.at[i, "_level"] = "Superseded"
                    stage.at[i, "_issues"] = stage.at[i, "_issues"] + [
                        ("Info", "Superseded by a later submission from the same branch")]
                else:
                    stage.at[i, "_issues"] = stage.at[i, "_issues"] + [
                        ("Info", f"Latest of {len(grp)} submissions from this branch")]
    stage["Check"] = stage["_level"]
    stage["Issues"] = stage["_issues"].apply(lambda L: "; ".join(t for _, t in L))
    return stage[STAGE_COLS].reset_index(drop=True), pd.DataFrame(summary)


# --------------------------------------------------------------------------- #
# Stored records: one tab per month ("Oct-2026"), one row per branch
# --------------------------------------------------------------------------- #
def _normalise_month_df(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for c in MONTH_COLS:
        if c not in df.columns:
            df[c] = np.nan if c in NUM_COLS else ""
    df = df[MONTH_COLS]
    for c in TEXT_COLS:
        df[c] = df[c].fillna("").astype(str).str.strip().replace({"nan": ""})
    for c in NUM_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    return df.reset_index(drop=True)


def derive_prev_ytd(records: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """Fill the Previous-month and YTD columns from the PRESENT-month figures of the saved months.

    Branches now report the present month only. For every saved month and branch:
      Previous = the branch's Present figure of the calendar month before (0 if none saved,
                 e.g. September 2026, the month before the first report).
      YTD      = sum of the branch's Present figures from January to this month of the same
                 calendar year (months with nothing saved, e.g. Jan-Sep 2026, count as 0).
                 Balance indicators (YTD rule 'latest') use the present month's figure.
    Recomputed every time, so a late or corrected month flows through to later months.
    """
    out = {}
    keys = sorted(records)
    for key in keys:
        d = records[key].copy()
        d = d.reset_index(drop=True)
        yr = key[:4]
        pk = prev_key(key)
        prev_df = records.get(pk)
        prev_map = {}
        if prev_df is not None and not prev_df.empty:
            prev_map = prev_df.drop_duplicates("Branch Code").set_index("Branch Code")
        same_year = [k for k in keys if k[:4] == yr and k <= key]
        for code in KRI_CODES:
            pres = pd.to_numeric(d[code], errors="coerce")
            if len(prev_map):
                pv = d["Branch Code"].map(pd.to_numeric(prev_map[code], errors="coerce"))
            else:
                pv = pd.Series(np.nan, index=d.index)
            d[num_col(code, "Previous")] = pv.fillna(0.0)
            if IND_BY_CODE[code]["ytd"] == "latest":
                d[num_col(code, "YTD")] = pres.fillna(0.0)
            else:
                tot = pres.fillna(0.0).copy()
                for k in same_year:
                    if k == key:
                        continue
                    o = records[k].drop_duplicates("Branch Code").set_index("Branch Code")
                    tot = tot + d["Branch Code"].map(pd.to_numeric(o[code], errors="coerce")).fillna(0.0)
                d[num_col(code, "YTD")] = tot
        out[key] = d
    return out


def load_records() -> dict[str, pd.DataFrame]:
    if not cfg.RECORDS_FILE.exists():
        return {}
    xl = pd.ExcelFile(cfg.RECORDS_FILE)
    out = {}
    for sheet in xl.sheet_names:
        key = tab_to_key(sheet)
        if key is None or sheet in RESERVED_SHEETS:
            continue
        df = pd.read_excel(xl, sheet_name=sheet, dtype={c: str for c in TEXT_COLS})
        if "Branch Code" in df.columns:
            df = df[df["Region Code"].notna() | df["Branch Code"].notna()]
        df = df[df["Region Code"].astype(str).str.strip().str.upper() != "TOTAL"]
        out[key] = _normalise_month_df(df)
    return derive_prev_ytd(dict(sorted(out.items())))


def load_log() -> pd.DataFrame:
    if cfg.RECORDS_FILE.exists():
        xl = pd.ExcelFile(cfg.RECORDS_FILE)
        if "Log" in xl.sheet_names:
            df = pd.read_excel(xl, sheet_name="Log", dtype=str).fillna("")
            for c in LOG_COLS:
                if c not in df.columns:
                    df[c] = ""
            return df[LOG_COLS]
    return pd.DataFrame(columns=LOG_COLS)


def save_records(records: dict[str, pd.DataFrame], log: pd.DataFrame) -> None:
    cfg.DATA_DIR.mkdir(parents=True, exist_ok=True)
    records = derive_prev_ytd({k: _normalise_month_df(v) for k, v in records.items()})
    _backup(cfg.RECORDS_FILE)
    tmp = cfg.RECORDS_FILE.with_name(cfg.RECORDS_FILE.stem + ".tmp.xlsx")
    legend = pd.DataFrame([{"Code": i["code"], "Indicator": i["full"], "Unit": i["unit"]} for i in INDICATORS])
    with pd.ExcelWriter(tmp, engine="openpyxl") as xw:
        for key in sorted(records):
            _normalise_month_df(records[key]).to_excel(xw, sheet_name=key_to_tab(key), index=False)
        legend.to_excel(xw, sheet_name="Legend", index=False)
        log[LOG_COLS].to_excel(xw, sheet_name="Log", index=False)
    wb = load_workbook(tmp)
    first_kri_col = len(TEXT_COLS) + 1
    for ws in wb.worksheets:
        if ws.title == "Legend":
            _style_sheet(ws, widths={"Code": 10, "Indicator": 120, "Unit": 26}, freeze="A2")
            continue
        if ws.title == "Log":
            _style_sheet(ws, widths={"Timestamp": 18, "Source File": 38, "Notes": 60}, freeze="A2")
            continue
        _style_sheet(ws, widths={"Region Name": 24, "Branch Name": 26, "Responder Email": 28, "Submitted At": 17,
                                 "Prepared By": 20, "Prepared Designation": 20, "Approved By": 20,
                                 "Approved Designation": 20, "RHO Confirmed By": 22, "Source File": 26},
                     freeze="E2")
        n = ws.max_row - 1
        total_row = ws.max_row + 1
        ws.cell(total_row, 1, "TOTAL")
        ws.cell(total_row, 2, "All branches")
        for c in range(first_kri_col, first_kri_col + len(NUM_COLS)):
            L = get_column_letter(c)
            ws.cell(total_row, c, f"=SUM({L}2:{L}{n + 1})")
        for cell in ws[total_row]:
            cell.font = Font(name=FONT, bold=True, size=10)
            cell.fill = PatternFill("solid", fgColor="DCE6F1")
            cell.border = BORDER
        for c, col in enumerate(NUM_COLS, start=first_kri_col):
            fmt = "#,##0.00" if IND_BY_CODE[base_code(col)]["kind"] == "value" else "#,##0"
            for r in range(2, total_row + 1):
                ws.cell(r, c).number_format = fmt
                ws.cell(r, c).alignment = Alignment(horizontal="right")
            ws.column_dimensions[get_column_letter(c)].width = 10 if col in KRI_CODES else 9
    wb.save(tmp)
    _replace(tmp, cfg.RECORDS_FILE)


def blank_month(branches: pd.DataFrame) -> pd.DataFrame:
    df = pd.DataFrame({c: branches[c] for c in ("Region Code", "Region Name", "Branch Code", "Branch Name")})
    df["Status"] = cfg.STATUS_NONE
    return _normalise_month_df(df)


def stage_to_rows(stage: pd.DataFrame, original: pd.DataFrame | None = None,
                  rho_confirmed: dict | None = None) -> pd.DataFrame:
    """Ticked staging rows -> month-tab rows. 'Edited' = 'Yes' when a KRI was changed by hand."""
    d = stage[stage["Include"] == True].copy()  # noqa: E712
    d["Edited"] = ""
    if original is not None:
        for i in d.index:
            if i in original.index:
                a = pd.to_numeric(d.loc[i, NUM_COLS], errors="coerce").fillna(-1).values
                b = pd.to_numeric(original.loc[i, NUM_COLS], errors="coerce").fillna(-1).values
                if not np.allclose(a, b):
                    d.at[i, "Edited"] = "Yes"
    d["RHO Confirmed By"] = d["Region Code"].map(rho_confirmed or {}).fillna("")
    return _normalise_month_df(d)


def merge_month(existing: pd.DataFrame | None, new_rows: pd.DataFrame,
                branches: pd.DataFrame, replace: bool = False) -> pd.DataFrame:
    base = blank_month(branches)
    ex = None
    if existing is not None and not replace:
        ex = _normalise_month_df(existing)
        base = base[~base["Branch Code"].isin(ex["Branch Code"])]
        base = pd.concat([ex, base], ignore_index=True)
    new_rows = _normalise_month_df(new_rows)
    if ex is not None and not new_rows.empty:   # keep an RHO confirmation already recorded for the region
        prior = ex[ex["RHO Confirmed By"] != ""].drop_duplicates("Region Code").set_index("Region Code")["RHO Confirmed By"]
        blank = new_rows["RHO Confirmed By"] == ""
        new_rows.loc[blank, "RHO Confirmed By"] = new_rows.loc[blank, "Region Code"].map(prior).fillna("")
    base = base[~base["Branch Code"].isin(new_rows["Branch Code"])]
    out = pd.concat([base, new_rows], ignore_index=True)
    order = {c: i for i, c in enumerate(branches["Branch Code"])}
    out["_o"] = out["Branch Code"].map(order).fillna(100_000)
    return _normalise_month_df(out.sort_values("_o", kind="stable").drop(columns="_o"))


def append_log(log: pd.DataFrame, **kw) -> pd.DataFrame:
    entry = {c: "" for c in LOG_COLS}
    entry["Timestamp"] = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    entry.update({k: str(v) for k, v in kw.items()})
    return pd.concat([log, pd.DataFrame([entry])], ignore_index=True)


def delete_month(records: dict, key: str) -> dict:
    records = dict(records)
    records.pop(key, None)
    return records


# --------------------------------------------------------------------------- #
# Calculations
# --------------------------------------------------------------------------- #
def _subset(records, key, region_codes=None, branch_codes=None):
    d = records.get(key)
    if d is None:
        return None
    if region_codes is not None:
        d = d[d["Region Code"].isin(region_codes)]
    if branch_codes is not None:
        d = d[d["Branch Code"].isin(branch_codes)]
    return d


def _totals(df: pd.DataFrame) -> pd.Series:
    return df[KRI_CODES].apply(pd.to_numeric, errors="coerce").fillna(0.0).sum()


def _sum_or_nan(d: pd.DataFrame | None, col: str) -> float:
    if d is None or d.empty or col not in d.columns:
        return np.nan
    s = pd.to_numeric(d[col], errors="coerce")
    return float(s.sum()) if s.notna().any() else np.nan


def annex1(records: dict[str, pd.DataFrame], key: str, region_codes=None, branch_codes=None):
    """Annex I figures = the sum of what the branches reported (Present, Previous month, Year to date)."""
    d = _subset(records, key, region_codes, branch_codes)
    rows = []
    for ind in INDICATORS:
        c = ind["code"]
        present = _sum_or_nan(d, c)
        rows.append({
            "No": ind["no"], "Code": c, "Indicator": ind["full"], "Short": ind["short"],
            "Unit": ind["unit"], "Kind": ind["kind"], "Group": ind["group"],
            "Present": 0.0 if np.isnan(present) else present,
            "Previous": _sum_or_nan(d, num_col(c, "Previous")),
            "YTD": _sum_or_nan(d, num_col(c, "YTD")),
        })
    meta = {"previous_key": prev_key(key), "ytd_months": [], "source": "branches"}
    return pd.DataFrame(rows), meta


MATRIX_COLS = ["Region Code", "Region Name", "Branches", "Reported", "Nil", "Status"]


def _measure_cols(measure: str) -> dict:
    """{stored column: indicator code} for the chosen measure (Present / Previous / YTD)."""
    return {num_col(c, measure): c for c in KRI_CODES}


def region_matrix(records, key, region_codes=None, branch_codes=None, measure: str = "Present") -> pd.DataFrame:
    """Region totals calculated from the branch rows (columns are named by indicator code)."""
    d = _subset(records, key, region_codes, branch_codes)
    if d is None or d.empty:
        return pd.DataFrame(columns=MATRIX_COLS + KRI_CODES)
    cols = _measure_cols(measure)
    d = d.copy()
    d["_n"] = 1
    d["_rep"] = (d["Status"] != cfg.STATUS_NONE).astype(int)
    d["_nil"] = (d["Status"] == cfg.STATUS_NIL).astype(int)
    agg = d.groupby(["Region Code", "Region Name"], sort=False).agg(
        {**{c: "sum" for c in cols}, "_n": "sum", "_rep": "sum", "_nil": "sum"}).reset_index()
    agg.loc[agg["_rep"] == 0, list(cols)] = np.nan   # nobody reported -> blank, not zero
    out = agg.rename(columns={"_n": "Branches", "_rep": "Reported", "_nil": "Nil", **cols})
    out["Status"] = out["Reported"].astype(int).astype(str) + " of " + out["Branches"].astype(int).astype(str) + " branches"
    return out[MATRIX_COLS + KRI_CODES]


def branch_matrix(records, key, region_codes=None, branch_codes=None, measure: str = "Present") -> pd.DataFrame:
    d = _subset(records, key, region_codes, branch_codes)
    cols = _measure_cols(measure)
    ident = ["Region Code", "Region Name", "Branch Code", "Branch Name", "Status", "Submitted At"]
    if d is None:
        return pd.DataFrame(columns=ident + KRI_CODES)
    return d[ident + list(cols)].rename(columns=cols).reset_index(drop=True)


def month_rows(records, key, region_codes=None, branch_codes=None) -> pd.DataFrame | None:
    """The saved branch rows of a month (all columns: present, previous, YTD)."""
    d = _subset(records, key, region_codes, branch_codes)
    return None if d is None else d.reset_index(drop=True)


def trend_table(records, region_codes=None, branch_codes=None) -> pd.DataFrame:
    rows = {month_label(k): _totals(_subset(records, k, region_codes, branch_codes)) for k in sorted(records)}
    return pd.DataFrame(rows).T if rows else pd.DataFrame(columns=KRI_CODES)


def branch_status_table(records, key, regions, branches, bdue=5, rdue=10) -> pd.DataFrame:
    """Every branch in the master (plus any extra stored branch) with its submission status."""
    names = dict(zip(regions["Region Code"], regions["Region Name"]))
    empty = pd.DataFrame(columns=MONTH_COLS)
    full = merge_month(records.get(key), empty, branches)
    d_branch, d_region = due_date(key, bdue), due_date(key, rdue)
    ts = pd.to_datetime(full["Submitted At"], errors="coerce")
    tl = np.where(ts.isna(), "-", np.where(ts > d_region, f"After {ordinal(rdue)}",
                                           np.where(ts > d_branch, f"After {ordinal(bdue)}", "On time")))
    emails = dict(zip(branches["Branch Code"], branches["Branch Email"]))
    out = pd.DataFrame({
        "Region Code": full["Region Code"],
        "Region Name": full["Region Code"].map(names).fillna(full["Region Name"]),
        "Branch Code": full["Branch Code"], "Branch Name": full["Branch Name"],
        "Status": full["Status"], "Submitted At": full["Submitted At"], "Timeliness": tl,
        "Prepared By": full["Prepared By"],
        "Responder Email": full["Responder Email"],
        "Branch Email": full["Branch Code"].map(emails).fillna(""),
        "RHO Confirmed By": full["RHO Confirmed By"],
    })
    return out.reset_index(drop=True)


def region_status_table(records, key, regions, branches, bdue=5, rdue=10) -> pd.DataFrame:
    bs = branch_status_table(records, key, regions, branches, bdue, rdue)
    rows = []
    for _, r in regions.iterrows():
        sub = bs[bs["Region Code"] == r["Region Code"]]
        n = len(sub)
        sub_n = int((sub["Status"] == cfg.STATUS_SUBMITTED).sum())
        nil_n = int((sub["Status"] == cfg.STATUS_NIL).sum())
        none_n = int((sub["Status"] == cfg.STATUS_NONE).sum())
        conf = sub["RHO Confirmed By"][sub["RHO Confirmed By"] != ""]
        rep = sub_n + nil_n
        status = ("No branches in master" if n == 0 else "Complete" if rep == n
                  else "Not started" if rep == 0 else "Partial")
        rows.append({
            "Region Code": r["Region Code"], "Region Name": r["Region Name"], "RHO Email": r["Officer Email"],
            "Branches": n, "Submitted": sub_n, "Nil": nil_n, "Not Submitted": none_n,
            "% Reported": round(rep / n * 100, 1) if n else 0.0,
            "After due date": int((~sub["Timeliness"].isin(["On time", "-"])).sum()),
            "RHO Confirmed By": conf.iloc[0] if len(conf) else "", "Status": status,
        })
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- #
# Multi-month combined ("master") report
# --------------------------------------------------------------------------- #
def months_in_range(records: dict, start: str, end: str) -> list[str]:
    return [k for k in sorted(records) if start <= k <= end]


def missing_months(records: dict, start: str, end: str) -> list[str]:
    """Calendar months between start and end (inclusive) that have no saved tab."""
    out, k = [], start
    while k <= end:
        if k not in records:
            out.append(k)
        k = f"{next_month_start(k).year:04d}-{next_month_start(k).month:02d}"
    return out


def combined_flat(records: dict, keys: list[str], region_codes=None) -> pd.DataFrame:
    """Every branch row of every selected month in ONE table (a 'Month' column added)."""
    frames = []
    for k in keys:
        d = _subset(records, k, region_codes)
        if d is None or d.empty:
            continue
        d = d.copy()
        d.insert(0, "Month Key", k)
        d.insert(0, "Month", month_label(k))
        frames.append(d)
    if not frames:
        return pd.DataFrame(columns=["Month", "Month Key"] + MONTH_COLS)
    return pd.concat(frames, ignore_index=True)


def _range_total(flat: pd.DataFrame, group_cols: list[str], last_key: str) -> pd.DataFrame:
    """Sum over the months - except period-end balances (YTD rule 'latest'), which take the last month."""
    total = flat.groupby(group_cols, sort=False)[KRI_CODES].sum(min_count=1)
    last = flat[flat["Month Key"] == last_key].groupby(group_cols, sort=False)[KRI_CODES].sum(min_count=1)
    for ind in INDICATORS:
        if ind["ytd"] == "latest":
            total[ind["code"]] = last[ind["code"]].reindex(total.index)
    return total.reset_index()


def combined_tables(records: dict, keys: list[str], region_codes=None) -> dict:
    flat = combined_flat(records, keys, region_codes)
    out = {"flat": flat}
    if flat.empty:
        return out
    last_key = keys[-1]
    # bank-wide: one row per month + the period total
    by_month = flat.groupby(["Month Key", "Month"], sort=True)[KRI_CODES].sum().reset_index()
    bank = by_month.drop(columns="Month Key")
    flat2 = flat.assign(_all="All regions")
    tot = _range_total(flat2, ["_all"], last_key).drop(columns="_all")
    tot.insert(0, "Month", f"PERIOD TOTAL ({month_label(keys[0])} - {month_label(last_key)})")
    out["bank_by_month"] = pd.concat([bank, tot], ignore_index=True)
    # region by month (sum of branches)
    rows = []
    for k in keys:
        mx = region_matrix(records, k, region_codes)
        if not mx.empty:
            mx = mx.copy()
            mx.insert(0, "Month", month_label(k))
            rows.append(mx.drop(columns=["Status"]))
    out["region_by_month"] = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()
    out["region_total"] = _range_total(flat, ["Region Code", "Region Name"], last_key)
    out["branch_total"] = _range_total(flat, ["Region Code", "Region Name", "Branch Code", "Branch Name"], last_key)
    return out
