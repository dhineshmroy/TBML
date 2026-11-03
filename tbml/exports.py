"""Excel and PDF report builders (Annex I format) - People's Bank branding (gold & charcoal)."""
from __future__ import annotations

import datetime as dt
import io

import numpy as np
import pandas as pd
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name

from . import config as cfg
from .config import INDICATORS, KRI_CODES
from .core import find_logo, fmt_val, key_to_tab, month_label

# ---- brand palette (taken from the People's Bank logo) -------------------------------------------
INK, GOLD, GOLD_D, GOLD_L, CREAM = "#1C1C1C", "#C9A227", "#A67C00", "#F3E3A8", "#FBF6E6"
GREY, LINE = "#6B6B6B", "#D9D2BC"
GREEN, GREEN_L = "#2E7D32", "#DCEBDD"
RED, RED_L = "#B23A3A", "#F3D9D9"
AMBER, AMBER_L = "#E39B00", "#FBF0C8"
BLUE_L = "#DCE8F3"
GROUP_STRONG = {
    "Alerts & STRs": "#A67C00",
    "High-risk jurisdictions": "#8E3B3B",
    "Customer & transaction behaviour": "#2F5D80",
    "Relationships & high-risk customers": "#3E7A4A",
}
GROUP_TINT = {
    "Alerts & STRs": "#FBF0C8",
    "High-risk jurisdictions": "#F3D9D9",
    "Customer & transaction behaviour": "#DCE8F3",
    "Relationships & high-risk customers": "#DCEBDD",
}
IND = cfg.IND_BY_CODE


def _logo_bytes(px: int = 260):
    """Small PNG copy of the bank logo for reports (None if no logo file is present)."""
    p = find_logo()
    if p is None:
        return None
    try:
        from PIL import Image
        im = Image.open(p).convert("RGBA")
        im.thumbnail((px, px))
        out = io.BytesIO()
        im.save(out, format="PNG", optimize=True)
        return out.getvalue()
    except Exception:
        return None


def _gs(code):  # strong colour of an indicator's group
    return GROUP_STRONG.get(IND[code]["group"], INK)


def _gt(code):
    return GROUP_TINT.get(IND[code]["group"], CREAM)


def _filter_note(all_regions: int, shown: int) -> str:
    if shown < all_regions:
        return (f"FILTERED VIEW - {shown} of {all_regions} regions only. "
                "This is NOT the official bank-wide report.")
    return ""


def _change_text(present, previous, prev_label, kind="count"):
    """('+12 vs Sep 2026', colour) - an increase in a risk indicator is shown in red."""
    if previous is None or (isinstance(previous, float) and np.isnan(previous)):
        return "no previous-month figure reported", GREY
    d = present - previous
    if abs(d) < 1e-9:
        return f"no change vs {prev_label}", GREY
    s = f"{d:+,.2f}" if kind == "value" else f"{d:+,.0f}"
    return f"{s} vs {prev_label}", (RED if d > 0 else GREEN)


# =============================================================================================
# EXCEL helpers
# =============================================================================================
class XF:
    """Cached xlsxwriter format factory (Arial 10 base)."""

    def __init__(self, wb):
        self.wb, self.cache = wb, {}

    def __call__(self, **kw):
        base = {"font_name": "Arial", "font_size": 10, "valign": "vcenter"}
        base.update(kw)
        k = tuple(sorted(base.items()))
        if k not in self.cache:
            self.cache[k] = self.wb.add_format(base)
        return self.cache[k]


def _banner(ws, xf, last_col, title, subtitle, logo):
    band = xf(bg_color=INK)
    for r in (0, 1, 2):
        ws.set_row(r, 24, band)
    ws.set_row(3, 4, xf(bg_color=GOLD))
    ws.merge_range(0, 0, 2, last_col, "", band)
    ws.write_rich_string(
        0, 0, xf(bold=True, font_size=16, font_color="#FFFFFF"), title + "\n",
        xf(font_size=10, font_color=GOLD_L), subtitle,
        xf(bg_color=INK, text_wrap=True, indent=9, valign="vcenter"))
    if logo:
        ws.insert_image(0, 0, "logo.png", {"image_data": io.BytesIO(logo), "x_scale": 0.24, "y_scale": 0.24,
                                           "x_offset": 8, "y_offset": 4, "object_position": 3})
    ws.hide_gridlines(2)


def _page_setup(ws, landscape=True):
    if landscape:
        ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.set_margins(0.4, 0.4, 0.7, 0.6)
    ws.set_header('&L&"Arial,Bold"People\'s Bank&R&"Arial"STRICTLY CONFIDENTIAL')
    ws.set_footer('&L&"Arial"Generated &D &T&R&"Arial"Page &P of &N')


def _kri_header_fmt(xf, code):
    return xf(bold=True, font_color="#FFFFFF", bg_color=_gs(code), border=1, border_color="#FFFFFF",
              align="center", text_wrap=True)


def _num_fmt(xf, kind, **kw):
    return xf(border=1, border_color=LINE, num_format="#,##0.00" if kind == "value" else "#,##0",
              align="right", **kw)


def _status_cf(ws, xf, r1, c1, r2, c2):
    for text, fg, bg in (("Submitted", GREEN, GREEN_L), ("Nil Report", GOLD_D, AMBER_L), ("Not Submitted", RED, RED_L),
                         ("Complete", GREEN, GREEN_L), ("Partial", GOLD_D, AMBER_L), ("Not started", RED, RED_L)):
        ws.conditional_format(r1, c1, r2, c2, {"type": "cell", "criteria": "==", "value": f'"{text}"',
                                                "format": xf(bold=True, font_color=fg, bg_color=bg)})


def _timeliness_cf(ws, xf, r1, c1, r2, c2):
    ws.conditional_format(r1, c1, r2, c2, {"type": "text", "criteria": "begins with", "value": "After",
                                            "format": xf(bold=True, font_color="#9C4A00", bg_color="#FFE3C2")})
    ws.conditional_format(r1, c1, r2, c2, {"type": "cell", "criteria": "==", "value": '"On time"',
                                            "format": xf(font_color=GREEN, bg_color=GREEN_L)})


def _write_status_df(ws, xf, df, top):
    hdr = xf(bold=True, font_color="#FFFFFF", bg_color=INK, border=1, border_color=GOLD, align="center", text_wrap=True)
    txt = xf(border=1, border_color=LINE)
    for c, h in enumerate(df.columns):
        ws.write(top, c, h, hdr)
    ws.set_row(top, 30)
    for r, (_, row) in enumerate(df.iterrows(), start=top + 1):
        for c, v in enumerate(row):
            if isinstance(v, (int, float, np.integer, np.floating)) and not pd.isna(v):
                ws.write_number(r, c, float(v), xf(border=1, border_color=LINE, align="center",
                                                  num_format="0.0" if "%" in df.columns[c] else "0"))
            else:
                ws.write(r, c, "" if pd.isna(v) else str(v), txt)
    last = top + len(df)
    for c, name in enumerate(df.columns):
        if name == "Status":
            _status_cf(ws, xf, top + 1, c, last, c)
        if name == "Timeliness":
            _timeliness_cf(ws, xf, top + 1, c, last, c)
        if name == "% Reported":
            ws.conditional_format(top + 1, c, last, c, {"type": "data_bar", "bar_color": GOLD, "bar_solid": True,
                                                         "min_type": "num", "min_value": 0, "max_type": "num", "max_value": 100})
    ws.autofilter(top, 0, last, len(df.columns) - 1)


# =============================================================================================
# EXCEL - monthly Annex I report
# =============================================================================================
def build_excel(key: str, annex: pd.DataFrame, meta: dict, matrix: pd.DataFrame,
                status: pd.DataFrame, trend: pd.DataFrame, info: dict,
                branch_df: pd.DataFrame | None = None, branch_status: pd.DataFrame | None = None,
                branch_full: pd.DataFrame | None = None) -> bytes:
    s = info["settings"]
    logo = _logo_bytes()
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    xf = XF(wb)
    banner_note = _filter_note(info["total_regions"], len(matrix))
    warn = xf(bold=True, font_color=RED)
    note = xf(italic=True, font_color=GREY, text_wrap=True, valign="top")
    label = xf(bold=True, font_color=INK)
    prev_label = month_label(meta["previous_key"]) if meta["previous_key"] else ""

    # create sheets up-front in final order
    wd = wb.add_worksheet("Dashboard")
    wa = wb.add_worksheet("Annex I")
    rw = wb.add_worksheet("Region-wise")
    bw = wb.add_worksheet("Branch-wise")
    rp = wb.add_worksheet("Region Present-Prev-YTD") if branch_full is not None else None
    bp = wb.add_worksheet("Branch Present-Prev-YTD") if branch_full is not None else None
    st = wb.add_worksheet("Region Status")
    bs = wb.add_worksheet("Branch Status") if branch_status is not None else None
    tr = wb.add_worksheet("Monthly Trend")
    lg = wb.add_worksheet("Legend")
    for w, colr in ((wd, GOLD), (wa, INK), (rw, GOLD_D), (bw, GOLD_D), (rp, GOLD_D), (bp, GOLD_D), (st, "#2F5D80"), (bs, "#2F5D80"),
                    (tr, "#3E7A4A"), (lg, GREY)):
        if w is not None:
            w.set_tab_color(colr)

    # ------------------------------------------------------------------ Region-wise
    HR = 5
    _banner(rw, xf, 3 + len(KRI_CODES) - 1, f"Region-wise TBML indicators - {month_label(key)}",
            "Region figures are the sum of the branch submissions", logo)
    if banner_note:
        rw.write(4, 0, banner_note, warn)
    heads = ["Region Code", "Region Name", "Branches reported"]
    hdr = xf(bold=True, font_color="#FFFFFF", bg_color=INK, border=1, border_color=GOLD, align="center", text_wrap=True)
    for c, h in enumerate(heads):
        rw.write(HR, c, h, hdr)
    for c, code in enumerate(KRI_CODES, start=3):
        rw.write(HR, c, f"{code}\n{IND[code]['short']}", _kri_header_fmt(xf, code))
    rw.set_row(HR, 82)
    rw.set_column(0, 0, 11)
    rw.set_column(1, 1, 28)
    rw.set_column(2, 2, 22)
    rw.set_column(3, 3 + len(KRI_CODES) - 1, 13)
    txt = xf(border=1, border_color=LINE, text_wrap=True)
    txt_c = xf(border=1, border_color=LINE, align="center")
    for r, (_, row) in enumerate(matrix.iterrows(), start=HR + 1):
        rw.write(r, 0, row["Region Code"], txt_c)
        rw.write(r, 1, row["Region Name"], txt)
        full = row["Reported"] >= row["Branches"]
        none = row["Reported"] == 0
        rw.write(r, 2, row["Status"], xf(border=1, border_color=LINE, align="center", bold=True,
                                         font_color=GREEN if full else (RED if none else GOLD_D),
                                         bg_color=GREEN_L if full else (RED_L if none else AMBER_L)))
        for c, code in enumerate(KRI_CODES, start=3):
            v = row[code]
            f = _num_fmt(xf, IND[code]["kind"])
            rw.write_blank(r, c, None, f) if pd.isna(v) else rw.write_number(r, c, float(v), f)
    first, last = HR + 1, HR + len(matrix)
    total_row = last + 1
    tot_t = xf(border=1, border_color=GOLD, bold=True, bg_color=GOLD_L)
    rw.write(total_row, 0, "TOTAL", tot_t)
    rw.write(total_row, 1, "All regions shown", tot_t)
    rw.write(total_row, 2, "", tot_t)
    for c, code in enumerate(KRI_CODES, start=3):
        col = xl_col_to_name(c)
        val = float(pd.to_numeric(matrix[code], errors="coerce").fillna(0).sum())
        rw.write_formula(total_row, c, f"=SUM({col}{first + 1}:{col}{last + 1})",
                         _num_fmt(xf, IND[code]["kind"], bold=True, bg_color=GOLD_L), val)
        rw.conditional_format(first, c, last, c, {"type": "2_color_scale", "min_color": "#FFFFFF", "max_color": "#E8B923"})
    rw.freeze_panes(HR + 1, 2)
    _page_setup(rw)

    # ------------------------------------------------------------------ Branch-wise
    _banner(bw, xf, 4 + len(KRI_CODES), f"Branch-wise TBML indicators - {month_label(key)}",
            "One row per branch (as submitted via the regional Forms)", logo)
    if banner_note:
        bw.write(4, 0, banner_note, warn)
    bcols = ["Region Code", "Region Name", "Branch Code", "Branch Name", "Status"]
    for c, h in enumerate(bcols):
        bw.write(HR, c, h, hdr)
    for c, code in enumerate(KRI_CODES, start=len(bcols)):
        bw.write(HR, c, f"{code}\n{IND[code]['short']}", _kri_header_fmt(xf, code))
    bw.set_row(HR, 82)
    for c, w in enumerate([11, 24, 12, 28, 15]):
        bw.set_column(c, c, w)
    bw.set_column(5, 5 + len(KRI_CODES) - 1, 13)
    if branch_df is not None and len(branch_df):
        for r, (_, row) in enumerate(branch_df.iterrows(), start=HR + 1):
            for c, col in enumerate(bcols):
                bw.write(r, c, row[col], txt if col in ("Region Name", "Branch Name") else txt_c)
            for c, code in enumerate(KRI_CODES, start=len(bcols)):
                v = row[code]
                f = _num_fmt(xf, IND[code]["kind"])
                bw.write_blank(r, c, None, f) if pd.isna(v) else bw.write_number(r, c, float(v), f)
        lastb = HR + len(branch_df)
        _status_cf(bw, xf, HR + 1, 4, lastb, 4)
        for c, code in enumerate(KRI_CODES, start=len(bcols)):
            bw.conditional_format(HR + 1, c, lastb, c, {"type": "2_color_scale", "min_color": "#FFFFFF", "max_color": "#E8B923"})
        bw.autofilter(HR, 0, lastb, len(bcols) + len(KRI_CODES) - 1)
    bw.freeze_panes(HR + 1, 4)
    _page_setup(bw)

    # ------------------------------------------------------------------ Present, Previous & YTD sheets
    if branch_full is not None:
        pcols = [cfg.num_col(c, p) for c in KRI_CODES for p in ("Present", "Previous", "YTD")]   # per indicator: Present, Previous, YTD (as in Annex I)

        def prev_ytd_sheet(ws, df, ident, title, sub):
            _banner(ws, xf, len(ident) + len(pcols) - 1, title, sub, logo)
            if banner_note:
                ws.write(4, 0, banner_note, warn)
            for c, h in enumerate(ident):
                ws.write(HR, c, h, hdr)
            for k, col in enumerate(pcols, start=len(ident)):
                code = col[:6]
                label = f"{col} Present" if col == code else col
                ws.write(HR, k, f"{label}\n{IND[code]['short']}", _kri_header_fmt(xf, code))
            ws.set_row(HR, 82)
            for c, w_ in enumerate([11, 26, 12, 28][:len(ident)]):
                ws.set_column(c, c, w_)
            ws.set_column(len(ident), len(ident) + len(pcols) - 1, 12)
            for r, (_, row) in enumerate(df.iterrows(), start=HR + 1):
                for c, col in enumerate(ident):
                    ws.write(r, c, row[col], txt)
                for k, col in enumerate(pcols, start=len(ident)):
                    v = row[col]
                    f = _num_fmt(xf, IND[col[:6]]["kind"])
                    ws.write_blank(r, k, None, f) if pd.isna(v) else ws.write_number(r, k, float(v), f)
            if len(df):
                ws.autofilter(HR, 0, HR + len(df), len(ident) + len(pcols) - 1)
            ws.freeze_panes(HR + 1, len(ident))
            _page_setup(ws)

        reg_df = (branch_full.groupby(["Region Code", "Region Name"], sort=False)[pcols].sum(min_count=1).reset_index())
        prev_ytd_sheet(rp, reg_df, ["Region Code", "Region Name"], f"Region totals - Present month, Previous month and Year to date - {month_label(key)}",
                       "Sum of the figures reported by the branches")
        prev_ytd_sheet(bp, branch_full[["Region Code", "Region Name", "Branch Code", "Branch Name"] + pcols],
                       ["Region Code", "Region Name", "Branch Code", "Branch Name"],
                       f"Branch figures - Present month, Previous month and Year to date - {month_label(key)}", "As reported by each branch")

    # ------------------------------------------------------------------ Annex I
    AH = 9
    A_FIRST = AH + 1
    _banner(wa, xf, 5, "Annex I - Monthly TBML Key Risk Indicator Reporting Format",
            f"{s['bank_name']}  |  {s['department']}", logo)
    wa.merge_range(5, 0, 5, 2, f"Reporting Month:   {month_label(key)}",
                   xf(bold=True, font_size=12, font_color="#FFFFFF", bg_color=GOLD_D, indent=1))
    wa.merge_range(6, 0, 6, 2, f"Branch/RHO/Department/Unit:   {s['unit_label']}",
                   xf(bold=True, font_color=INK, bg_color=CREAM, indent=1))
    if banner_note:
        wa.write(7, 0, banner_note, warn)
    heads = [("No.", INK), ("TBML Key Risk Indicator", INK), ("Unit", INK), ("Present Period Figure", GOLD_D),
             ("Previous Period Figure", GREY), ("YTD", "#3A3A3A")]
    for c, (h, colr) in enumerate(heads):
        wa.write(AH, c, h, xf(bold=True, font_color="#FFFFFF", bg_color=colr, border=1, border_color="#FFFFFF",
                              align="center", text_wrap=True))
    wa.set_row(AH, 34)
    wa.set_column(0, 0, 7)
    wa.set_column(1, 1, 88)
    wa.set_column(2, 2, 22)
    wa.set_column(3, 5, 17)
    na = xf(border=1, border_color=LINE, align="right", font_color="#8A8A8A", italic=True)
    for i, (_, r) in enumerate(annex.iterrows()):
        row = A_FIRST + i
        code = r["Code"]
        wa.set_row(row, 50 if len(r["Indicator"]) > 150 else 36)
        wa.write_number(row, 0, int(r["No"]), xf(bold=True, font_color="#FFFFFF", bg_color=_gs(code), border=1,
                                                border_color="#FFFFFF", align="center"))
        wa.write(row, 1, r["Indicator"], xf(border=1, border_color=LINE, text_wrap=True, bg_color=_gt(code)))
        wa.write(row, 2, r["Unit"], xf(border=1, border_color=LINE, text_wrap=True, bg_color=_gt(code)))
        col = xl_col_to_name(3 + i)
        wa.write_formula(row, 3, f"='Region-wise'!{col}{total_row + 1}",
                         _num_fmt(xf, r["Kind"], bold=True, bg_color=GOLD_L, font_size=11), float(r["Present"]))
        if pd.isna(r["Previous"]):
            wa.write(row, 4, "N/A", na)
        else:
            wa.write_number(row, 4, float(r["Previous"]), _num_fmt(xf, r["Kind"]))
        if pd.isna(r["YTD"]):
            wa.write(row, 5, "N/A", na)
        else:
            wa.write_number(row, 5, float(r["YTD"]), _num_fmt(xf, r["Kind"], bg_color=CREAM))
    r0 = A_FIRST + len(annex) + 1
    prev_txt = (f"Present period = {month_label(key)}; Previous period = {month_label(meta['previous_key'])}. "
                "Present, Previous and YTD are the totals of the figures reported by the branches in their Forms "
                "(KRI-05 and KRI-08 in Rs. Million).")
    ytd_txt = ""
    wa.merge_range(r0, 0, r0, 5, prev_txt + " " + ytd_txt, note)
    wa.set_row(r0, 44)
    # indicator-group legend
    lr = r0 + 1
    for gi, (g, colr) in enumerate(GROUP_STRONG.items()):
        wa.write(lr + gi, 0, "", xf(bg_color=colr))
        wa.write(lr + gi, 1, g, xf(font_color=GREY, italic=True))
    sr = lr + len(GROUP_STRONG) + 1
    sign_l = xf(bold=True, font_color=INK, bg_color=CREAM, border=1, border_color=LINE)
    sign_v = xf(border=1, border_color=LINE)
    for off, (who, name, des, date) in enumerate([
            ("Prepared by", info.get("prepared_by", ""), info.get("prepared_designation", ""), info.get("prepared_date", "")),
            ("Approved by", info.get("approved_by", ""), info.get("approved_designation", ""), info.get("approved_date", ""))]):
        base = sr + off * 4
        wa.merge_range(base, 0, base, 2, f"{who}:  {name}", sign_l)
        wa.merge_range(base + 1, 0, base + 1, 2, f"Designation:  {des}", sign_v)
        wa.merge_range(base + 2, 0, base + 2, 2, f"Date:  {date}", sign_v)
    _page_setup(wa)
    wa.set_first_sheet()

    # ------------------------------------------------------------------ Region / Branch status
    _banner(st, xf, 10, f"Submission status by region - {month_label(key)}",
            "Calculated from the branch submissions", logo)
    _write_status_df(st, xf, status, 5)
    st.set_column(0, 0, 11)
    st.set_column(1, 2, 28)
    st.set_column(3, len(status.columns), 15)
    st.set_column(len(status.columns) - 2, len(status.columns) - 2, 28)
    st.freeze_panes(6, 2)
    _page_setup(st)
    if bs is not None:
        _banner(bs, xf, 10, f"Submission status by branch - {month_label(key)}",
                "Every branch in the master list", logo)
        _write_status_df(bs, xf, branch_status, 5)
        bs.set_column(0, 0, 11)
        bs.set_column(1, 3, 26)
        bs.set_column(4, len(branch_status.columns), 18)
        bs.freeze_panes(6, 4)
        _page_setup(bs)

    # ------------------------------------------------------------------ Trend
    TH = 5
    _banner(tr, xf, len(KRI_CODES), "Bank-wide monthly totals (all saved months)",
            "Sum of all branch submissions", logo)
    tr.write(TH, 0, "Month", hdr)
    for c, code in enumerate(KRI_CODES, start=1):
        tr.write(TH, c, f"{code}\n{IND[code]['short']}", _kri_header_fmt(xf, code))
    tr.set_row(TH, 82)
    tr.set_column(0, 0, 16)
    tr.set_column(1, len(KRI_CODES), 13)
    for r, (lab, row) in enumerate(trend.iterrows(), start=TH + 1):
        tr.write(r, 0, lab, txt)
        for c, code in enumerate(KRI_CODES, start=1):
            tr.write_number(r, c, float(row[code]), _num_fmt(xf, IND[code]["kind"]))
    tr.freeze_panes(TH + 1, 1)
    _page_setup(tr)

    # ------------------------------------------------------------------ Legend
    _banner(lg, xf, 2, "Indicator legend", "Wording as in Circular No. 1996/2026, Annex I", logo)
    for c, h in enumerate(["Code", "Indicator", "Unit"]):
        lg.write(5, c, h, hdr)
    for r, ind in enumerate(INDICATORS, start=6):
        lg.write(r, 0, ind["code"], xf(bold=True, font_color="#FFFFFF", bg_color=_gs(ind["code"]), border=1,
                                        border_color="#FFFFFF", align="center"))
        lg.write(r, 1, ind["full"], xf(border=1, border_color=LINE, text_wrap=True, bg_color=_gt(ind["code"])))
        lg.write(r, 2, ind["unit"], xf(border=1, border_color=LINE, bg_color=_gt(ind["code"])))
        lg.set_row(r, 44 if len(ind["full"]) > 150 else 30)
    lg.set_column(0, 0, 10)
    lg.set_column(1, 1, 110)
    lg.set_column(2, 2, 24)

    # ------------------------------------------------------------------ Dashboard (first tab)
    _banner(wd, xf, 11, f"TBML Key Risk Indicators - {month_label(key)}",
            f"{s['bank_name']}  |  {s['department']}  |  Executive summary", logo)
    if banner_note:
        wd.write(4, 0, banner_note, warn)
    wd.set_column(0, 11, 13.5)

    def tile(r, c, lab, formula, cached, kind, delta_txt, delta_col, strong=GOLD):
        tl = xf(font_size=9, bold=True, font_color=GREY, bg_color=CREAM, left=5, left_color=strong, indent=1)
        tv = xf(font_size=22, bold=True, font_color=INK, bg_color=CREAM, left=5, left_color=strong, indent=1,
                align="left", num_format={"value": "#,##0.00", "pct": '0.0"%"'}.get(kind, "#,##0"))
        td = xf(font_size=9, bold=True, font_color=delta_col, bg_color=CREAM, left=5, left_color=strong, indent=1)
        wd.merge_range(r, c, r, c + 2, lab, tl)
        wd.merge_range(r + 1, c, r + 1, c + 2, "", tv)
        if formula:
            wd.write_formula(r + 1, c, formula, tv, cached)
        else:
            wd.write_number(r + 1, c, cached, tv)
        wd.merge_range(r + 2, c, r + 2, c + 2, delta_txt, td)
        wd.set_row(r, 18)
        wd.set_row(r + 1, 36)
        wd.set_row(r + 2, 18)

    by = {r["Code"]: (i, r) for i, (_, r) in enumerate(annex.iterrows())}

    def risk_tile(r, c, code, strong):
        i, a = by[code]
        txt_, colr = _change_text(a["Present"], a["Previous"], prev_label, a["Kind"])
        tile(r, c, f"{code}  {a['Short']}", f"='Annex I'!D{A_FIRST + i + 1}", float(a["Present"]), a["Kind"], txt_, colr, strong)

    row1 = [("KRI-01", _gs("KRI-01")), ("KRI-04", _gs("KRI-04")), ("KRI-06", _gs("KRI-06")), ("KRI-05", _gs("KRI-05"))]
    for k, (code, colr) in enumerate(row1):
        risk_tile(5, k * 3, code, colr)
    wd.set_row(8, 8)
    row2 = [("KRI-07", _gs("KRI-07")), ("KRI-14", _gs("KRI-14")), ("KRI-16", _gs("KRI-16"))]
    for k, (code, colr) in enumerate(row2):
        risk_tile(9, k * 3, code, colr)
    tot = status[["Branches", "Submitted", "Nil", "Not Submitted", "After due date"]].sum()
    rep = int(tot["Submitted"] + tot["Nil"])
    pct = rep / tot["Branches"] * 100 if tot["Branches"] else 0
    tile(9, 9, "BRANCH REPORTING COMPLETENESS", None, round(pct, 1), "pct",
         f"{rep} of {int(tot['Branches'])} branches reported", GREEN if pct >= 99.95 else (AMBER if pct >= 80 else RED), GREEN)
    wd.set_row(12, 8)

    # chart data (status counts)
    wd.write(60, 0, "Chart data", xf(font_color=GREY, italic=True))
    for k, (lab_, v) in enumerate([("Submitted", tot["Submitted"]), ("Nil report", tot["Nil"]),
                                   ("Not submitted", tot["Not Submitted"])]):
        wd.write(61 + k, 0, lab_, xf(font_color=GREY))
        wd.write_number(61 + k, 1, float(v), xf(font_color=GREY))

    n_reg = len(matrix)
    c1 = wb.add_chart({"type": "column"})
    for code, colr in (("KRI-01", INK), ("KRI-04", GOLD)):
        col = 3 + KRI_CODES.index(code)
        c1.add_series({"name": f"{code} {IND[code]['short']}", "categories": ["Region-wise", first, 0, last, 0],
                       "values": ["Region-wise", first, col, last, col], "fill": {"color": colr}, "gap": 60})
    c1.set_title({"name": "Alerts generated vs suspected TBML alerts, by region", "name_font": {"size": 11, "color": INK}})
    c1.set_legend({"position": "bottom"})
    c1.set_y_axis({"major_gridlines": {"visible": True, "line": {"color": "#E6E0CC"}}})
    c1.set_chartarea({"border": {"color": LINE}})
    c1.set_size({"width": 700, "height": 330})
    wd.insert_chart(13, 0, c1)

    c2 = wb.add_chart({"type": "doughnut"})
    c2.add_series({"name": "Branch submissions", "categories": ["Dashboard", 61, 0, 63, 0],
                   "values": ["Dashboard", 61, 1, 63, 1],
                   "points": [{"fill": {"color": GREEN}}, {"fill": {"color": GOLD}}, {"fill": {"color": RED}}],
                   "data_labels": {"value": True, "font": {"color": "#FFFFFF", "bold": True}}})
    c2.set_title({"name": "Branch submissions", "name_font": {"size": 11, "color": INK}})
    c2.set_legend({"position": "bottom"})
    c2.set_hole_size(55)
    c2.set_chartarea({"border": {"color": LINE}})
    c2.set_size({"width": 440, "height": 330})
    wd.insert_chart(13, 7, c2, {"x_offset": 60})

    if len(trend) >= 2:
        c3 = wb.add_chart({"type": "line"})
        for code, colr in (("KRI-01", INK), ("KRI-04", GOLD_D), ("KRI-06", RED)):
            col = 1 + KRI_CODES.index(code)
            c3.add_series({"name": f"{code} {IND[code]['short']}", "categories": ["Monthly Trend", TH + 1, 0, TH + len(trend), 0],
                           "values": ["Monthly Trend", TH + 1, col, TH + len(trend), col],
                           "line": {"color": colr, "width": 2.5},
                           "marker": {"type": "circle", "size": 7, "fill": {"color": colr}, "border": {"color": colr}}})
        c3.set_title({"name": "Monthly trend (bank-wide)", "name_font": {"size": 11, "color": INK}})
        c3.set_legend({"position": "bottom"})
        c3.set_y_axis({"major_gridlines": {"visible": True, "line": {"color": "#E6E0CC"}}})
        c3.set_chartarea({"border": {"color": LINE}})
        c3.set_size({"width": 1150, "height": 300})
        wd.insert_chart(31, 0, c3)
    wd.write(44 if len(trend) >= 2 else 31, 0,
             "Figures are the sum of all branch submissions. Arrows/colours: an increase in a risk indicator is shown in red.",
             xf(italic=True, font_color=GREY))
    _page_setup(wd)
    wd.print_area(0, 0, 46 if len(trend) >= 2 else 33, 11)
    wd.activate()
    wb.close()
    return buf.getvalue()


# =============================================================================================
# EXCEL - combined multi-month master file
# =============================================================================================
def build_combined_excel(keys: list[str], tables: dict, missing: list[str], info: dict) -> bytes:
    s = info["settings"]
    logo = _logo_bytes()
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    xf = XF(wb)
    warn = xf(bold=True, font_color=RED)
    note = xf(italic=True, font_color=GREY, text_wrap=True, valign="top")
    hdr = xf(bold=True, font_color="#FFFFFF", bg_color=INK, border=1, border_color=GOLD, align="center", text_wrap=True)
    txt = xf(border=1, border_color=LINE)
    period = f"{month_label(keys[0])} to {month_label(keys[-1])}" if keys else ""

    def _kb(col):
        return col[:6] if str(col).startswith("KRI-") and col[:6] in IND else None

    def write_table(ws, df, top, label_cols, total_last_row=False):
        for c, col in enumerate(df.columns):
            if _kb(col):
                ws.write(top, c, f"{col}\n{IND[_kb(col)]['short']}", _kri_header_fmt(xf, _kb(col)))
            else:
                ws.write(top, c, col, hdr)
        ws.set_row(top, 82)
        for r, (_, row) in enumerate(df.iterrows(), start=top + 1):
            is_tot = total_last_row and r == top + len(df)
            for c, col in enumerate(df.columns):
                v = row[col]
                if _kb(col):
                    f = _num_fmt(xf, IND[_kb(col)]["kind"], bold=is_tot, bg_color=GOLD_L if is_tot else "#FFFFFF")
                    ws.write_blank(r, c, None, f) if pd.isna(v) else ws.write_number(r, c, float(v), f)
                else:
                    ws.write(r, c, "" if pd.isna(v) else v,
                             xf(border=1, border_color=GOLD if is_tot else LINE, bold=is_tot,
                                bg_color=GOLD_L if is_tot else "#FFFFFF"))
        if len(df):
            ws.autofilter(top, 0, top + len(df), len(df.columns) - 1)
            for c, col in enumerate(df.columns):
                if _kb(col) and not total_last_row:
                    ws.conditional_format(top + 1, c, top + len(df), c,
                                          {"type": "2_color_scale", "min_color": "#FFFFFF", "max_color": "#E8B923"})
        ws.freeze_panes(top + 1, label_cols)

    def sheet(name, title, sub, ncols, colr):
        w = wb.add_worksheet(name)
        w.set_tab_color(colr)
        _banner(w, xf, ncols, title, sub, logo)
        _page_setup(w)
        return w

    ws0 = sheet("Summary", f"Combined TBML report: {period}", f"{s['bank_name']}  |  {s['department']}", len(KRI_CODES), GOLD)
    r = 4
    if missing:
        ws0.write(r, 0, "WARNING - months with no saved data in this period: " + ", ".join(month_label(k) for k in missing), warn)
        r += 1
    if info.get("filtered"):
        ws0.write(r, 0, "FILTERED VIEW - selected regions only; NOT the official bank-wide figures.", warn)
        r += 1
    write_table(ws0, tables["bank_by_month"], r + 1, 1, total_last_row=True)
    ws0.set_column(0, 0, 46)
    ws0.set_column(1, len(KRI_CODES), 13)
    nrow = r + 3 + len(tables["bank_by_month"])
    ws0.merge_range(nrow, 0, nrow, 8,
                    "Bank-wide totals = sum of all branch submissions. PERIOD TOTAL adds the months together, except KRI-02 "
                    "and KRI-03 (period-end balances), which show the last month's figure. KRI-05 and KRI-08 are in Rs. Million. "
                    "Branches that had not submitted are not in the figures.", note)
    ws0.set_row(nrow, 44)
    # trend chart of the monthly rows (exclude the period-total row)
    nm = len(tables["bank_by_month"]) - 1
    if nm >= 2:
        ch = wb.add_chart({"type": "line"})
        for code, colr in (("KRI-01", INK), ("KRI-04", GOLD_D), ("KRI-06", RED)):
            col = 1 + list(tables["bank_by_month"].columns).index(code) - 1 + 1 - 1
            col = list(tables["bank_by_month"].columns).index(code)
            ch.add_series({"name": f"{code} {IND[code]['short']}",
                           "categories": ["Summary", r + 2, 0, r + 1 + nm, 0],
                           "values": ["Summary", r + 2, col, r + 1 + nm, col],
                           "line": {"color": colr, "width": 2.5},
                           "marker": {"type": "circle", "size": 7, "fill": {"color": colr}, "border": {"color": colr}}})
        ch.set_title({"name": "Monthly trend (bank-wide)", "name_font": {"size": 11, "color": INK}})
        ch.set_legend({"position": "bottom"})
        ch.set_chartarea({"border": {"color": LINE}})
        ch.set_size({"width": 900, "height": 300})
        ws0.insert_chart(nrow + 2, 0, ch)

    rm = sheet("Region by Month", f"Region figures month by month - {period}", "Sum of the branches in each region", 6 + len(KRI_CODES), "#2F5D80")
    write_table(rm, tables["region_by_month"], 5, 3)
    rm.set_column(0, 0, 16)
    rm.set_column(1, 1, 12)
    rm.set_column(2, 2, 26)
    rm.set_column(3, 5, 10)
    rm.set_column(6, 6 + len(KRI_CODES), 13)

    rt = sheet("Region Total", f"Region totals for the whole period - {period}", "Period totals per region", 2 + len(KRI_CODES), "#2F5D80")
    write_table(rt, tables["region_total"], 5, 2)
    rt.set_column(0, 0, 12)
    rt.set_column(1, 1, 28)
    rt.set_column(2, 2 + len(KRI_CODES), 13)

    bt = sheet("Branch Total", f"Branch totals for the whole period - {period}", "Period totals per branch", 4 + len(KRI_CODES), "#3E7A4A")
    write_table(bt, tables["branch_total"], 5, 4)
    bt.set_column(0, 0, 12)
    bt.set_column(1, 1, 24)
    bt.set_column(2, 2, 12)
    bt.set_column(3, 3, 28)
    bt.set_column(4, 4 + len(KRI_CODES), 13)

    flat = tables["flat"].drop(columns=["Month Key"])
    ad = sheet("All Data", "Master file - every branch, every month in one table", "Filter or pivot as needed", len(flat.columns), GREY)
    write_table(ad, flat, 5, 4)
    ad.set_column(0, 0, 16)
    ad.set_column(1, 1, 12)
    ad.set_column(2, 3, 14)
    ad.set_column(4, len(flat.columns), 13)
    wb.close()
    return buf.getvalue()


# =============================================================================================
# PDF
# =============================================================================================
def _mpl():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    return plt


def _png(fig) -> bytes:
    out = io.BytesIO()
    fig.savefig(out, format="png", dpi=150, bbox_inches="tight", transparent=False, facecolor="white")
    _mpl().close(fig)
    return out.getvalue()


def _donut_png(sub, nil, none):
    plt = _mpl()
    fig, ax = plt.subplots(figsize=(3.9, 3.0))
    vals = [sub, nil, none]
    if sum(vals) == 0:
        vals = [0, 0, 1]
    ax.pie(vals, colors=[GREEN, GOLD, RED], startangle=90, counterclock=False,
           wedgeprops=dict(width=0.38, edgecolor="white", linewidth=2))
    tot = sub + nil + none
    ax.text(0, 0.07, f"{(sub + nil) / tot * 100:.0f}%" if tot else "-", ha="center", va="center", fontsize=20,
            fontweight="bold", color=INK)
    ax.text(0, -0.2, "reported", ha="center", va="center", fontsize=9, color=GREY)
    labels = [f"Submitted ({sub})", f"Nil report ({nil})", f"Not submitted ({none})"]
    ax.legend(labels, loc="upper center", bbox_to_anchor=(0.5, 0.02), ncol=1, frameon=False, fontsize=8)
    return _png(fig)


def _topbar_png(matrix, code="KRI-01", n=10):
    plt = _mpl()
    d = matrix[["Region Code", "Region Name", code]].copy()
    d["v"] = pd.to_numeric(d[code], errors="coerce").fillna(0)
    d = d.sort_values("v", ascending=False).head(n).iloc[::-1]
    fig, ax = plt.subplots(figsize=(5.2, 3.0))
    bars = ax.barh(d["Region Code"] + "  " + d["Region Name"].str.slice(0, 22), d["v"], color=GOLD, edgecolor=GOLD_D)
    for b_, v in zip(bars, d["v"]):
        ax.text(b_.get_width(), b_.get_y() + b_.get_height() / 2, f" {v:,.0f}", va="center", fontsize=7.5, color=INK)
    ax.tick_params(axis="y", labelsize=7.5)
    ax.tick_params(axis="x", labelsize=7.5)
    ax.spines[["top", "right"]].set_visible(False)
    ax.xaxis.grid(True, alpha=0.25)
    ax.set_axisbelow(True)
    ax.set_xlim(0, max(d["v"].max() * 1.15, 1))
    return _png(fig)


def _region_chart_png(matrix):
    plt = _mpl()
    d = matrix[matrix["Region Code"] != ""]
    if d.empty:
        return None
    x = np.arange(len(d))
    fig, ax = plt.subplots(figsize=(11, 4.4))
    w = 0.4
    ax.bar(x - w / 2, pd.to_numeric(d["KRI-01"], errors="coerce").fillna(0), w, label="KRI-01 Alerts generated", color=INK)
    ax.bar(x + w / 2, pd.to_numeric(d["KRI-04"], errors="coerce").fillna(0), w, label="KRI-04 Suspected TBML alerts", color=GOLD)
    ax.set_xticks(x)
    ax.set_xticklabels(d["Region Code"], rotation=60, ha="right", fontsize=8)
    ax.set_ylabel("Number of alerts", fontsize=9)
    ax.set_title("Alerts generated vs suspected TBML alerts, by region", fontsize=11, color=INK, loc="left", fontweight="bold")
    ax.spines[["top", "right"]].set_visible(False)
    ax.legend(fontsize=8, frameon=False)
    ax.grid(axis="y", alpha=0.25)
    ax.set_axisbelow(True)
    return _png(fig)


def build_pdf(key: str, annex: pd.DataFrame, meta: dict, matrix: pd.DataFrame,
              status: pd.DataFrame, info: dict) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.lib.utils import ImageReader
    from reportlab.platypus import (BaseDocTemplate, Frame, Image, NextPageTemplate, PageBreak,
                                    PageTemplate, Paragraph, Spacer, Table, TableStyle)

    s = info["settings"]
    C = colors.HexColor
    ink, gold, gold_d, gold_l, cream = C(INK), C(GOLD), C(GOLD_D), C(GOLD_L), C(CREAM)
    logo = _logo_bytes()
    ss = getSampleStyleSheet()
    mk = lambda name, **kw: ParagraphStyle(name, parent=ss["BodyText"], **kw)  # noqa: E731
    body = mk("b", fontName="Helvetica", fontSize=8.5, leading=11)
    small = mk("s", fontName="Helvetica", fontSize=7.5, leading=9.5)
    note = mk("n", fontName="Helvetica", fontSize=7.5, leading=10, textColor=C(GREY))
    warn = mk("w", fontName="Helvetica-Bold", fontSize=8.5, textColor=C(RED))
    cell_c = mk("c", fontName="Helvetica", fontSize=7.5, leading=9.5, alignment=TA_CENTER)
    cell_cb = mk("cb", fontName="Helvetica-Bold", fontSize=8, leading=10, alignment=TA_CENTER)
    cell_cw = mk("cw", fontName="Helvetica-Bold", fontSize=8, leading=10, alignment=TA_CENTER, textColor=colors.white)
    head = mk("hd", fontName="Helvetica-Bold", fontSize=7.5, leading=9.5, alignment=TA_CENTER, textColor=colors.white)
    h_title = mk("ht", fontName="Helvetica-Bold", fontSize=19, leading=23, textColor=ink, alignment=TA_LEFT)
    h_month = mk("hm", fontName="Helvetica-Bold", fontSize=13, leading=17, textColor=gold_d)
    h_unit = mk("hu", fontName="Helvetica", fontSize=9, leading=12, textColor=C(GREY))
    h2 = mk("h2", fontName="Helvetica-Bold", fontSize=11, leading=14, textColor=ink, spaceBefore=6, spaceAfter=4)
    t_lab = mk("tl", fontName="Helvetica-Bold", fontSize=7, leading=8.5, textColor=C(GREY))
    t_val = mk("tv", fontName="Helvetica-Bold", fontSize=19, leading=22, textColor=ink)
    generated = dt.datetime.now().strftime("%d %b %Y %H:%M")
    ml = month_label(key)
    buf = io.BytesIO()

    def decorate(canvas, doc):
        w, h = doc.pagesize
        canvas.saveState()
        canvas.setFillColor(ink)
        canvas.rect(0, h - 12 * mm, w, 12 * mm, stroke=0, fill=1)
        canvas.setFillColor(gold)
        canvas.rect(0, h - 13 * mm, w, 1 * mm, stroke=0, fill=1)
        if logo:
            canvas.drawImage(ImageReader(io.BytesIO(logo)), 10 * mm, h - 11 * mm, width=9.5 * mm, height=9.5 * mm, mask="auto")
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica-Bold", 10)
        canvas.drawString(22 * mm, h - 7.2 * mm, s["bank_name"])
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(gold_l)
        canvas.drawRightString(w - 12 * mm, h - 7.2 * mm, f"Monthly TBML Key Risk Indicator Report  |  {ml}")
        canvas.setStrokeColor(gold)
        canvas.setLineWidth(1.2)
        canvas.line(14 * mm, 14 * mm, w - 14 * mm, 14 * mm)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(C(GREY))
        canvas.drawString(14 * mm, 9.5 * mm, f"STRICTLY CONFIDENTIAL | {s['bank_name']} - {s['department']}")
        canvas.drawRightString(w - 14 * mm, 9.5 * mm, f"Generated {generated}  |  Page {doc.page}")
        canvas.restoreState()

    doc = BaseDocTemplate(buf, pagesize=A4, title=f"TBML KRI Report - {ml}", author=s["department"],
                          leftMargin=14 * mm, rightMargin=14 * mm, topMargin=18 * mm, bottomMargin=18 * mm)
    pw, ph = A4
    lw, lh = landscape(A4)
    doc.addPageTemplates([
        PageTemplate("portrait", frames=[Frame(14 * mm, 18 * mm, pw - 28 * mm, ph - 36 * mm, id="p")], onPage=decorate, pagesize=A4),
        PageTemplate("landscape", frames=[Frame(12 * mm, 18 * mm, lw - 24 * mm, lh - 36 * mm, id="l")], onPage=decorate, pagesize=landscape(A4)),
    ])

    el = []
    # ---------------- Page 1: executive summary ----------------
    if logo:
        title_tbl = Table([[Image(io.BytesIO(logo), width=24 * mm, height=24 * mm),
                            [Paragraph("Monthly TBML Key Risk Indicator Report", h_title),
                             Paragraph(ml, h_month),
                             Paragraph(f"{s['bank_name']}  |  {s['unit_label']}", h_unit)]]],
                          colWidths=[30 * mm, None])
    else:
        title_tbl = Table([[[Paragraph("Monthly TBML Key Risk Indicator Report", h_title), Paragraph(ml, h_month),
                             Paragraph(f"{s['bank_name']}  |  {s['unit_label']}", h_unit)]]])
    title_tbl.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                                   ("LINEBELOW", (0, 0), (-1, 0), 2, gold), ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
    el += [title_tbl, Spacer(1, 6)]
    banner_txt = _filter_note(info["total_regions"], len(matrix))
    if banner_txt:
        el += [Paragraph(banner_txt, warn), Spacer(1, 4)]

    prev_label = month_label(meta["previous_key"]) if meta["previous_key"] else ""
    by = {r["Code"]: r for _, r in annex.iterrows()}

    def tile(code):
        a = by[code]
        txt_, colr = _change_text(a["Present"], a["Previous"], prev_label, a["Kind"])
        return [Paragraph(f"{code}  {a['Short'].upper()}", t_lab), Paragraph(fmt_val(a["Present"], a["Kind"]), t_val),
                Paragraph(f'<font color="{colr}">{txt_}</font>', small)]

    tot = status[["Branches", "Submitted", "Nil", "Not Submitted", "After due date"]].sum()
    rep = int(tot["Submitted"] + tot["Nil"])
    pct = rep / tot["Branches"] * 100 if tot["Branches"] else 0
    comp_col = GREEN if pct >= 99.95 else (AMBER if pct >= 80 else RED)
    comp = [Paragraph("BRANCH REPORTING COMPLETENESS", t_lab), Paragraph(f"{pct:.1f}%", t_val),
            Paragraph(f'<font color="{comp_col}">{rep} of {int(tot["Branches"])} branches reported</font>', small)]
    cells1 = [tile(c) for c in ("KRI-01", "KRI-04", "KRI-06", "KRI-05")]
    cells2 = [tile(c) for c in ("KRI-07", "KRI-14", "KRI-16")] + [comp]
    gap = ""
    grid = [[cells1[0], gap, cells1[1], gap, cells1[2], gap, cells1[3]],
            [gap] * 7,
            [cells2[0], gap, cells2[1], gap, cells2[2], gap, cells2[3]]]
    g = Table(grid, colWidths=[41 * mm, 6 * mm, 41 * mm, 6 * mm, 41 * mm, 6 * mm, 41 * mm],
              rowHeights=[25 * mm, 4 * mm, 25 * mm])
    style = [("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 8), ("RIGHTPADDING", (0, 0), (-1, -1), 4)]
    codes1 = ("KRI-01", "KRI-04", "KRI-06", "KRI-05")
    codes2 = ("KRI-07", "KRI-14", "KRI-16", None)
    for k, col in enumerate((0, 2, 4, 6)):
        for rowi, codes in ((0, codes1), (2, codes2)):
            strong = _gs(codes[k]) if codes[k] else GOLD
            style += [("BACKGROUND", (col, rowi), (col, rowi), cream), ("LINEBEFORE", (col, rowi), (col, rowi), 3.5, C(strong)),
                      ("BOX", (col, rowi), (col, rowi), 0.4, C(LINE))]
    g.setStyle(TableStyle(style))
    el += [g, Spacer(1, 10)]

    charts = Table([[Paragraph("Branch submissions", h2), Paragraph("Top 10 regions - alerts generated (KRI-01)", h2)],
                    [Image(io.BytesIO(_donut_png(int(tot["Submitted"]), int(tot["Nil"]), int(tot["Not Submitted"]))), width=76 * mm, height=76 * mm * 0.86),
                     Image(io.BytesIO(_topbar_png(matrix)), width=98 * mm, height=98 * mm * 0.56)]],
                   colWidths=[80 * mm, 102 * mm])
    charts.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    el += [charts, Spacer(1, 6)]
    nums = matrix.copy()
    for c_ in ("KRI-01", "KRI-04"):
        nums[c_] = pd.to_numeric(nums[c_], errors="coerce").fillna(0)
    pts = []
    if len(nums) and nums["KRI-01"].max() > 0:
        t1 = nums.sort_values("KRI-01", ascending=False).iloc[0]
        pts.append(f"Highest alert volume: <b>{t1['Region Code']} {t1['Region Name']}</b> with <b>{t1['KRI-01']:,.0f}</b> alerts generated.")
    if len(nums) and nums["KRI-04"].max() > 0:
        t4 = nums.sort_values("KRI-04", ascending=False).iloc[0]
        pts.append(f"Most suspected TBML alerts: <b>{t4['Region Code']} {t4['Region Name']}</b> with <b>{t4['KRI-04']:,.0f}</b>.")
    inc = int((status["Status"].isin(["Partial", "Not started"])).sum())
    pts.append(f"<b>{inc}</b> of {len(status)} regions have not yet reported for every branch; "
               f"<b>{int(tot['Not Submitted'])}</b> branch(es) outstanding, <b>{int(tot['After due date'])}</b> submitted after the due date.")
    hl = Table([[Paragraph("Highlights", h2)]] + [[Paragraph("&bull; " + p_, body)] for p_ in pts], colWidths=[182 * mm])
    hl.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), cream), ("LINEBEFORE", (0, 0), (0, -1), 3.5, gold),
                            ("BOX", (0, 0), (-1, -1), 0.4, C(LINE)), ("LEFTPADDING", (0, 0), (-1, -1), 10),
                            ("TOPPADDING", (0, 1), (-1, -1), 3), ("BOTTOMPADDING", (0, 1), (-1, -1), 3)]))
    el += [hl, Spacer(1, 8)]
    el.append(Paragraph(
        "Figures are the sum of all branch submissions received through the regional Forms. "
        "A rise in a risk indicator compared with the previous month is shown in red, a fall in green. "
        "Detailed Annex I figures follow on the next page.", note))

    # ---------------- Page 2: Annex I ----------------
    el += [PageBreak(), Paragraph("Annex I - Monthly TBML Key Risk Indicator Reporting Format", h2)]
    sub = Table([[Paragraph("<b>Reporting Month:</b>", body), Paragraph(ml, body)],
                 [Paragraph("<b>Branch/RHO/Department/Unit:</b>", body), Paragraph(s["unit_label"], body)]],
                colWidths=[52 * mm, None])
    sub.setStyle(TableStyle([("LINEBELOW", (0, -1), (-1, -1), 0.6, gold), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                             ("BOTTOMPADDING", (0, 0), (-1, -1), 3), ("TOPPADDING", (0, 0), (-1, -1), 2)]))
    el += [sub, Spacer(1, 6)]
    data = [[Paragraph(t_, head) for t_ in ["No.", "TBML Key Risk Indicator", "Unit", "Present Period", "Previous Period", "YTD"]]]
    for _, r in annex.iterrows():
        data.append([Paragraph(str(int(r["No"])), cell_cw), Paragraph(r["Indicator"], small), Paragraph(r["Unit"], small),
                     Paragraph(fmt_val(r["Present"], r["Kind"]), cell_cb), Paragraph(fmt_val(r["Previous"], r["Kind"]), cell_c),
                     Paragraph(fmt_val(r["YTD"], r["Kind"]), cell_c)])
    t = Table(data, colWidths=[9 * mm, 85 * mm, 25 * mm, 21 * mm, 21 * mm, 21 * mm], repeatRows=1)
    ts = [("BACKGROUND", (0, 0), (-1, 0), ink), ("LINEBELOW", (0, 0), (-1, 0), 1.6, gold),
          ("BACKGROUND", (3, 0), (3, 0), gold_d), ("BACKGROUND", (4, 0), (4, 0), C(GREY)),
          ("GRID", (0, 0), (-1, -1), 0.4, C(LINE)), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]
    for i, (_, r) in enumerate(annex.iterrows(), start=1):
        ts += [("BACKGROUND", (0, i), (0, i), C(_gs(r["Code"]))), ("BACKGROUND", (1, i), (2, i), C(_gt(r["Code"]))),
               ("BACKGROUND", (3, i), (3, i), gold_l), ("BACKGROUND", (5, i), (5, i), cream)]
    t.setStyle(TableStyle(ts))
    el += [t, Spacer(1, 4)]
    legend = Table([[Paragraph(g_, mk("lg", fontName="Helvetica", fontSize=6.8, leading=8, textColor=colors.white, alignment=TA_CENTER))
                     for g_ in GROUP_STRONG]], colWidths=[45.5 * mm] * 4)
    legend.setStyle(TableStyle([("BACKGROUND", (i, 0), (i, 0), C(c_)) for i, c_ in enumerate(GROUP_STRONG.values())]
                               + [("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
                                  ("LINEAFTER", (0, 0), (-2, -1), 1.5, colors.white)]))
    el += [legend, Spacer(1, 4)]
    el.append(Paragraph(
        f"Present period = {ml}; Previous period = {prev_label}. Present, Previous and YTD are the totals of the figures "
        "reported by the branches in their Forms. Rupee values are in Rs. Million, converted by each branch at the mid "
        "rate of the last day of the period.", note))
    el.append(Spacer(1, 10))

    def sign_block(lab, name, des, date):
        rows = [[Paragraph(f"<b>{lab}</b>", body), Paragraph(name or "", body)],
                [Paragraph("Designation:", body), Paragraph(des or "", body)],
                [Paragraph("Date:", body), Paragraph(date or "", body)]]
        tb = Table(rows, colWidths=[28 * mm, 60 * mm])
        tb.setStyle(TableStyle([("LINEBELOW", (1, 0), (1, -1), 0.4, colors.grey), ("TOPPADDING", (0, 0), (-1, -1), 5)]))
        return tb

    sig = Table([[sign_block("Prepared by:", info.get("prepared_by"), info.get("prepared_designation"), info.get("prepared_date")),
                  sign_block("Approved by:", info.get("approved_by"), info.get("approved_designation"), info.get("approved_date"))]],
                colWidths=[92 * mm, 92 * mm])
    sig.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    el.append(sig)

    # ---------------- Submission status by region ----------------
    el += [Spacer(1, 14), Paragraph(f"Submission status by region - {ml}", h2)]
    el.append(Paragraph(
        f"Branches in total: <b>{int(tot['Branches'])}</b> &nbsp;|&nbsp; Submitted: <b>{int(tot['Submitted'])}</b> "
        f"&nbsp;|&nbsp; Nil reports: <b>{int(tot['Nil'])}</b> &nbsp;|&nbsp; Not submitted: <b>{int(tot['Not Submitted'])}</b> "
        f"&nbsp;|&nbsp; Reported: <b>{pct:.1f}%</b> &nbsp;|&nbsp; After due date: <b>{int(tot['After due date'])}</b>", body))
    el.append(Spacer(1, 4))
    cols = ["Region Code", "Region Name", "Branches", "Submitted", "Nil", "Not Submitted", "% Reported", "After due date",
            "Status", "RHO Confirmed By"]
    labels = ["Code", "Region Name", "Branches", "Submitted", "Nil", "Not submitted", "% reported", "After due", "Status",
              "RHO confirmed by"]
    srows = [[Paragraph(c_, head) for c_ in labels]]
    st_style = [("BACKGROUND", (0, 0), (-1, 0), ink), ("LINEBELOW", (0, 0), (-1, 0), 1.6, gold),
                ("GRID", (0, 0), (-1, -1), 0.3, C(LINE)), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]
    for i, (_, r) in enumerate(status.iterrows(), start=1):
        srows.append([Paragraph(str(r[c_]) if str(r[c_]) != "nan" else "", small) for c_ in cols])
        bg = {"Complete": GREEN_L, "Partial": AMBER_L, "Not started": RED_L}.get(r["Status"])
        if bg:
            st_style.append(("BACKGROUND", (8, i), (8, i), C(bg)))
    stt = Table(srows, colWidths=[11 * mm, 30 * mm, 18 * mm, 19 * mm, 9 * mm, 17 * mm, 15 * mm, 15 * mm, 18 * mm, 30 * mm], repeatRows=1)
    stt.setStyle(TableStyle(st_style))
    el += [stt, Paragraph("Region figures are the sum of branch submissions. Branch-level detail is in the Excel report.", note)]

    # ---------------- Region-wise (landscape) ----------------
    el += [NextPageTemplate("landscape"), PageBreak(), Paragraph(f"Region-wise figures - {ml}", h2)]
    if banner_txt:
        el.append(Paragraph(banner_txt, warn))
    hdr_row = [Paragraph("Region", head)] + [Paragraph(c.replace("KRI-", ""), head) for c in KRI_CODES]
    rows = [hdr_row]
    num = matrix[KRI_CODES].apply(pd.to_numeric, errors="coerce")
    for (_, r), (_, nr) in zip(matrix.iterrows(), num.iterrows()):
        label = f"{r['Region Code']} {r['Region Name']}"
        if r["Reported"] < r["Branches"]:
            label += f" ({int(r['Reported'])}/{int(r['Branches'])} branches)"
        rows.append([Paragraph(label, small)] + [
            Paragraph(fmt_val(nr[c], IND[c]["kind"]) if not pd.isna(nr[c]) else "-", cell_c) for c in KRI_CODES])
    tt = num.fillna(0).sum()
    rows.append([Paragraph("<b>TOTAL</b>", small)] + [Paragraph("<b>" + fmt_val(tt[c], IND[c]["kind"]) + "</b>", cell_c) for c in KRI_CODES])
    cw = [52 * mm] + [(lw - 24 * mm - 52 * mm) / len(KRI_CODES)] * len(KRI_CODES)
    rt = Table(rows, colWidths=cw, repeatRows=1)
    rs = [("BACKGROUND", (0, 0), (0, 0), ink), ("GRID", (0, 0), (-1, -1), 0.3, C(LINE)),
          ("BACKGROUND", (0, -1), (-1, -1), gold_l), ("LINEABOVE", (0, -1), (-1, -1), 1.2, gold_d),
          ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, C("#FAF7EE")]),
          ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
          ("LEFTPADDING", (0, 0), (-1, -1), 2), ("RIGHTPADDING", (0, 0), (-1, -1), 2)]
    for j, code in enumerate(KRI_CODES, start=1):
        rs.append(("BACKGROUND", (j, 0), (j, 0), C(_gs(code))))
    rt.setStyle(TableStyle(rs))
    el += [rt, Spacer(1, 4),
           Paragraph("Column numbers 1-18 correspond to KRI-01 to KRI-18 (see Annex I table); figures are summed from branch submissions. "
                     "KRI-05 and KRI-08 are in Rs. Million.", note)]
    png = _region_chart_png(matrix)
    if png:
        el += [PageBreak(), Paragraph("Region comparison", h2),
               Image(io.BytesIO(png), width=lw - 40 * mm, height=(lw - 40 * mm) * 0.40)]
    doc.build(el)
    return buf.getvalue()
