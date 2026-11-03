
"""TBML Reporting Portal - People's Bank (Channel Management Department).

Run with:  streamlit run app.py
"""

from __future__ import annotations
import sys
from pathlib import Path

# Force the directory containing app.py into the Python path
APP_DIR = Path(__file__).resolve().parent
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

import datetime as dt
import io

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from tbml import config as cfg
from tbml import core, exports
from tbml.config import INDICATORS, KRI_CODES, IND_BY_CODE, NUM_COLS, base_code

def _page_icon():
    p = core.find_logo()
    if p is None:
        return ":shield:"
    try:
        from PIL import Image
        return Image.open(p)
    except Exception:
        return ":shield:"


st.set_page_config(page_title="TBML Reporting Portal | People's Bank", page_icon=_page_icon(), layout="wide")

SETTINGS = cfg.load_settings()
BDUE, RDUE = int(SETTINGS["branch_due_day"]), int(SETTINGS["regional_due_day"])
INK, GOLD, GOLD_D, GOLD_L, CREAM = "#1C1C1C", "#C9A227", "#A67C00", "#F3E3A8", "#FBF6E6"
NAVY, ORANGE = INK, GOLD          # chart colours below use these two names
COLORWAY = [GOLD, INK, "#2F5D80", "#8E3B3B", "#3E7A4A", "#8A8A8A"]
ALL_STATUS = [cfg.STATUS_SUBMITTED, cfg.STATUS_NIL, cfg.STATUS_NONE]

st.markdown(
    f"""
<style>
.block-container {{padding-top: 1.4rem; max-width: 1500px;}}
.hero {{display: flex; align-items: center; gap: 20px; background: linear-gradient(115deg, {INK} 0%, #2B2B2B 70%, #3A3320 100%);
        padding: 18px 26px; border-radius: 14px; margin-bottom: 20px; border-bottom: 5px solid {GOLD};
        box-shadow: 0 4px 14px rgba(0,0,0,.18);}}
.hero img {{height: 66px; width: 66px; border-radius: 50%; box-shadow: 0 0 0 3px rgba(201,162,39,.35);}}
.hero h1 {{color: #FFFFFF; margin: 0; font-size: 1.65rem; font-weight: 700; letter-spacing: .2px;}}
.hero p {{color: {GOLD_L}; margin: 4px 0 0 0; font-size: .95rem;}}
div[data-testid="stMetric"] {{background: {CREAM}; border: 1px solid #E8DFC0; border-left: 6px solid {GOLD};
        padding: 12px 16px; border-radius: 10px; box-shadow: 0 2px 6px rgba(28,28,28,.07);}}
div[data-testid="stMetricLabel"] p {{font-weight: 700; color: #6B6B6B; font-size: .8rem; text-transform: uppercase; letter-spacing: .3px;}}
div[data-testid="stMetricValue"] {{color: {INK}; font-weight: 800;}}
section[data-testid="stSidebar"] {{background: {CREAM}; border-right: 1px solid #E8DFC0;}}
.sb-brand {{text-align: center; padding: 6px 0 4px 0;}}
.sb-brand img {{width: 92px; height: 92px; border-radius: 50%; box-shadow: 0 3px 10px rgba(0,0,0,.25);}}
.sb-brand h3 {{margin: 10px 0 0 0; color: {INK}; font-size: 1.15rem;}}
.sb-brand p {{margin: 2px 0 0 0; color: #6B6B6B; font-size: .8rem;}}
.stMarkdown h5 {{color: {INK}; border-left: 5px solid {GOLD}; padding-left: 10px; margin-top: 1.2rem;}}
.stTabs [data-baseweb="tab-list"] {{gap: 6px;}}
.stTabs [data-baseweb="tab"] {{background: {CREAM}; border-radius: 8px 8px 0 0; padding: 8px 18px; font-weight: 600;}}
.stTabs [aria-selected="true"] {{background: {INK}; color: {GOLD_L} !important;}}
.stButton > button, .stDownloadButton > button {{border-radius: 8px; font-weight: 600;}}
.stButton > button[kind="primary"] {{background: {INK}; border: 1px solid {GOLD}; color: {GOLD_L};}}
.stButton > button[kind="primary"]:hover {{background: {GOLD_D}; color: #FFFFFF; border-color: {GOLD_D};}}
[data-testid="stDataFrame"] {{border: 1px solid #E8DFC0; border-radius: 8px;}}
</style>
""",
    unsafe_allow_html=True,
)


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _logo_b64(px: int = 140) -> str:
    p = core.find_logo()
    if p is None:
        return ""
    try:
        import base64
        from PIL import Image
        im = Image.open(p).convert("RGBA")
        im.thumbnail((px, px))
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        return base64.b64encode(buf.getvalue()).decode()
    except Exception:
        return ""


LOGO_B64 = _logo_b64()


def banner(title: str, subtitle: str = "") -> None:
    img = f'<img src="data:image/png;base64,{LOGO_B64}"/>' if LOGO_B64 else ""
    st.markdown(f'<div class="hero">{img}<div><h1>{title}</h1><p>{subtitle}</p></div></div>', unsafe_allow_html=True)


_STATUS_CSS = {
    "Submitted": "background-color:#DCEBDD;color:#1B5E20;font-weight:600",
    "Complete": "background-color:#DCEBDD;color:#1B5E20;font-weight:600",
    "Nil Report": "background-color:#FBF0C8;color:#7A5B00;font-weight:600",
    "Partial": "background-color:#FBF0C8;color:#7A5B00;font-weight:600",
    "Not Submitted": "background-color:#F3D9D9;color:#8E2222;font-weight:600",
    "Not started": "background-color:#F3D9D9;color:#8E2222;font-weight:600",
}


def style_status(df: pd.DataFrame):
    """Colour the Status / Timeliness columns (green / amber / red)."""
    def status(v):
        return _STATUS_CSS.get(v, "")

    def timely(v):
        v = str(v)
        return ("background-color:#FFE3C2;color:#9C4A00;font-weight:600" if v.startswith("After")
                else "color:#2E7D32;font-weight:600" if v == "On time" else "")

    sty = df.style
    apply = sty.map if hasattr(sty, "map") else sty.applymap
    if "Status" in df.columns:
        sty = apply(status, subset=["Status"])
        apply = sty.map if hasattr(sty, "map") else sty.applymap
    if "Timeliness" in df.columns:
        sty = apply(timely, subset=["Timeliness"])
    return sty


def gate() -> None:
    pw = str(SETTINGS.get("password", ""))
    if not pw or st.session_state.get("auth"):
        return
    banner("TBML Reporting Portal", "Please sign in")
    entered = st.text_input("Password", type="password")
    if st.button("Sign in", type="primary"):
        if entered == pw:
            st.session_state["auth"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    st.stop()


def load_all():
    regions, branches = core.load_master()
    return regions, branches, core.load_records()


def kri_column_config() -> dict:
    return {c: st.column_config.NumberColumn(
        c, help=IND_BY_CODE[base_code(c)]["short"],
        format="%.2f" if IND_BY_CODE[base_code(c)]["kind"] == "value" else "%.0f")
        for c in NUM_COLS}


MEASURES = {"Present month": "Present", "Previous month": "Previous", "Year to date": "YTD"}


def to_xlsx_bytes(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for name, df in sheets.items():
            df.to_excel(xw, sheet_name=name[:31], index=False)
    return buf.getvalue()


def pick_stored_month(records, label="Reporting month", key="month_pick"):
    keys = sorted(records, reverse=True)
    if not keys:
        st.info("No monthly data has been saved yet. Go to **Upload & Validate** to add the first month.")
        return None
    return st.selectbox(label, keys, format_func=core.month_label, key=key)


def region_filter(month_df: pd.DataFrame, container=st, key="rf"):
    names = dict(zip(month_df["Region Code"], month_df["Region Name"]))
    return container.multiselect("Regions (empty = all)", list(names),
                                 format_func=lambda c: f"{c} - {names[c]}", key=key)


def need_master(regions, branches) -> bool:
    if regions.empty or branches.empty:
        st.warning("The Regions and Branches lists are not filled in yet. Open **Master Data** first.")
        return True
    return False


# --------------------------------------------------------------------------- #
# Pages
# --------------------------------------------------------------------------- #
def page_dashboard():
    regions, branches, records = load_all()
    banner("TBML Monitoring Dashboard", "Trade-Based Money Laundering key risk indicators - region and branch level")
    if not records:
        st.info("No data yet. Open **Upload & Validate**, load the Forms exports for the first month, and the dashboard will appear here.")
        return

    c1, c2, c3, c4 = st.columns([1.1, 2, 2, 1.6])
    key = c1.selectbox("Reporting month", sorted(records, reverse=True), format_func=core.month_label, key="dash_month")
    month_df = records[key]
    rsel = region_filter(month_df, c2, key="dash_regions")
    pool = month_df[month_df["Region Code"].isin(rsel)] if rsel else month_df
    bnames = dict(zip(pool["Branch Code"], pool["Branch Name"]))
    bsel = c3.multiselect("Branches (empty = all)", list(bnames), format_func=lambda c: f"{c} - {bnames[c]}", key="dash_branches")
    statuses = c4.multiselect("Branch status", ALL_STATUS, default=ALL_STATUS, key="dash_status")
    groups_all = list(dict.fromkeys(i["group"] for i in INDICATORS))
    g1, g2 = st.columns([3, 1])
    groups = g1.multiselect("Indicator groups", groups_all, default=groups_all, key="dash_groups")
    measure = MEASURES[g2.selectbox("Figures shown in charts and tables", list(MEASURES), key="dash_measure")]

    view = pool[pool["Branch Code"].isin(bsel)] if bsel else pool
    view = view[view["Status"].isin(statuses)]
    bcodes = list(view["Branch Code"])
    if not bcodes:
        st.warning("No branches match the filters.")
        return
    shown = [i["code"] for i in INDICATORS if i["group"] in groups] or KRI_CODES
    annex, meta = core.annex1(records, key, None, bcodes)
    mx = core.region_matrix(records, key, None, bcodes, measure=measure)
    bm = core.branch_matrix(records, key, None, bcodes, measure=measure)
    tr = core.trend_table(records, None, bcodes)
    bst = core.branch_status_table(records, key, regions, branches, BDUE, RDUE)
    bst = bst[bst["Branch Code"].isin(bcodes)]

    s = st.columns(6)
    s[0].metric("Regions in view", mx["Region Code"].nunique())
    s[1].metric("Branches in view", len(bcodes))
    s[2].metric("Submitted", int((bst["Status"] == cfg.STATUS_SUBMITTED).sum()))
    s[3].metric("Nil reports", int((bst["Status"] == cfg.STATUS_NIL).sum()))
    s[4].metric("Not submitted", int((bst["Status"] == cfg.STATUS_NONE).sum()))
    s[5].metric("After due date", int((~bst["Timeliness"].isin(["On time", "-"])).sum()))

    def metric(col, code):
        r = annex[annex["Code"] == code].iloc[0]
        delta = None if pd.isna(r["Previous"]) else r["Present"] - r["Previous"]
        d = None if delta is None else (f"{delta:+,.2f}" if r["Kind"] == "value" else f"{delta:+,.0f}")
        col.metric(r["Short"], core.fmt_val(r["Present"], r["Kind"]), d, delta_color="inverse",
                   help=f"{r['Code']} - change vs previous month")

    k = st.columns(4)
    for col, code in zip(k, ["KRI-01", "KRI-04", "KRI-06", "KRI-05"]):
        metric(col, code)
    st.caption("Arrows compare with the previous month reported by the branches (a rise in a risk indicator is shown in red). "
               "Filters apply to every chart and table; the trend line always shows the present month of each saved month.")

    opts = [c for c in KRI_CODES if c in shown]
    fmt_ind = lambda c: f"{c} - {IND_BY_CODE[c]['short']}"  # noqa: E731
    t1, t2, t3, t4, t5, t6, t7 = st.tabs(["Regions", "Top branches", "Monthly trend", "Heatmap",
                                          "Month-on-month", "Region table", "Branch table"])
    with t1:
        code = st.selectbox("Indicator", opts, format_func=fmt_ind, key="cmp_ind")
        d = mx[["Region Code", "Region Name", code]].copy()
        d["Value"] = pd.to_numeric(d[code], errors="coerce").fillna(0)
        d = d.sort_values("Value", ascending=False)
        fig = px.bar(d, x="Region Code", y="Value", hover_data=["Region Name"], text="Value", color_discrete_sequence=[ORANGE])
        fig.update_traces(texttemplate="%{text:,.2f}" if IND_BY_CODE[code]["kind"] == "value" else "%{text:,.0f}", textposition="outside")
        if len(d):
            fig.add_hline(y=float(d["Value"].mean()), line_dash="dot", line_color=NAVY, annotation_text="average", annotation_position="top left")
        fig.update_layout(height=430, margin=dict(t=30, b=10), plot_bgcolor="white", yaxis_title=IND_BY_CODE[code]["unit"], xaxis_title="")
        st.plotly_chart(fig, use_container_width=True)
    with t2:
        a, b = st.columns([3, 1])
        code = a.selectbox("Indicator", opts, format_func=fmt_ind, key="top_ind")
        n = b.slider("How many branches", 5, 50, 15)
        d = bm[["Region Code", "Branch Code", "Branch Name", code]].copy()
        d["Value"] = pd.to_numeric(d[code], errors="coerce").fillna(0)
        d = d.sort_values("Value", ascending=False).head(n).iloc[::-1]
        d["Label"] = d["Branch Code"] + " " + d["Branch Name"]
        fig = px.bar(d, x="Value", y="Label", orientation="h", text="Value", hover_data=["Region Code"], color_discrete_sequence=[NAVY])
        fig.update_traces(texttemplate="%{text:,.2f}" if IND_BY_CODE[code]["kind"] == "value" else "%{text:,.0f}", textposition="outside")
        fig.update_layout(height=max(360, 26 * len(d) + 120), margin=dict(t=20, b=10), plot_bgcolor="white", yaxis_title="", xaxis_title=IND_BY_CODE[code]["unit"])
        st.plotly_chart(fig, use_container_width=True)
    with t3:
        defaults = [c for c in ("KRI-01", "KRI-04") if c in opts]
        sel_k = st.multiselect("Indicators to plot", opts, default=defaults, format_func=fmt_ind, key="trend_ind")
        if sel_k and len(tr):
            plot = tr[sel_k].rename(columns={c: f"{c} {IND_BY_CODE[c]['short']}" for c in sel_k})
            fig = px.line(plot, markers=True, color_discrete_sequence=COLORWAY)
            fig.update_layout(height=430, margin=dict(t=30, b=10), plot_bgcolor="white", xaxis_title="Month",
                              yaxis_title="Total (branches in view)", legend_title="")
            st.plotly_chart(fig, use_container_width=True)
            if len(tr) < 2:
                st.caption("Only one month is stored so far - the trend line appears as more months are saved.")
        else:
            st.info("Select at least one indicator.")
    with t4:
        mat = mx.set_index("Region Code")[opts].apply(pd.to_numeric, errors="coerce").fillna(0)
        norm = mat.div(mat.max().replace(0, np.nan), axis=1).fillna(0)
        text = [[core.fmt_val(mat.iloc[i, j], IND_BY_CODE[mat.columns[j]]["kind"]) for j in range(mat.shape[1])] for i in range(mat.shape[0])]
        fig = go.Figure(go.Heatmap(z=norm.values, x=list(mat.columns), y=list(mat.index), text=text, texttemplate="%{text}",
                                   colorscale=[[0, "#FFFFFF"], [0.5, "#E8B923"], [1, "#8E3B3B"]], showscale=False, hovertemplate="Region %{y}<br>%{x}: %{text}<extra></extra>"))
        fig.update_layout(height=max(380, 24 * len(mat) + 120), margin=dict(t=20, b=10), yaxis=dict(autorange="reversed"), xaxis=dict(side="top"))
        st.plotly_chart(fig, use_container_width=True)
        st.caption("Colour intensity is relative within each indicator (column); the number is the actual figure.")
    with t5:
        if not meta["previous_key"]:
            st.info("The previous month has not been saved, so month-on-month movement is not available.")
        else:
            m = annex[annex["Code"].isin(shown)][["Code", "Short", "Unit", "Present", "Previous", "Kind"]].copy()
            m["Change"] = m["Present"] - m["Previous"]
            m["Change %"] = np.where(m["Previous"] > 0, m["Change"] / m["Previous"] * 100, np.nan)
            m = m.sort_values("Change %", ascending=False, na_position="last")
            st.dataframe(pd.DataFrame({
                "Indicator": m["Code"] + "  " + m["Short"], "Unit": m["Unit"],
                f"Present ({core.month_label(key)})": [core.fmt_val(v, k) for v, k in zip(m["Present"], m["Kind"])],
                f"Previous ({core.month_label(meta['previous_key'])})": [core.fmt_val(v, k) for v, k in zip(m["Previous"], m["Kind"])],
                "Change": [core.fmt_val(v, k) if v >= 0 else "-" + core.fmt_val(abs(v), k) for v, k in zip(m["Change"], m["Kind"])],
                "Change %": [("N/A" if pd.isna(v) else f"{v:+.1f}%") for v in m["Change %"]],
            }), hide_index=True, use_container_width=True)
    with t6:
        st.dataframe(mx, hide_index=True, use_container_width=True, column_config=kri_column_config())
    with t7:
        st.dataframe(style_status(bm), hide_index=True, use_container_width=True, column_config=kri_column_config(), height=520)


def page_upload():
    regions, branches, records = load_all()
    banner("Upload & Validate", "Load the region Forms exports (after each RHO has reviewed them), check every flag, then save")
    if need_master(regions, branches):
        return
    opts = core.month_options()
    today = dt.date.today()
    default_key = core.prev_key(f"{today.year:04d}-{today.month:02d}")
    key = st.selectbox("Reporting month", opts, index=opts.index(default_key) if default_key in opts else 0,
                       format_func=core.month_label)
    if key in records:
        have = int((records[key]["Status"] != cfg.STATUS_NONE).sum())
        st.caption(f"{core.month_label(key)} already has saved data for {have} of {len(branches)} branches.")

    ups = st.file_uploader("Excel files from the regions (select all 25 at once, or a few at a time)", type=["xlsx"],
                           accept_multiple_files=True)
    mode = st.radio("If this month already has saved data",
                    ["Update only the branches found in these files (recommended)", "Replace the whole month"], index=0)
    if not ups:
        st.markdown("**Before you upload:** each RHO sends the Excel they downloaded from their region's Form "
                    "(Responses -> Open in Excel), after checking or correcting the branch figures. Responses for other "
                    "months inside a file are ignored automatically.")
        return
    files = []
    for u in ups:
        try:
            files.append((u.name, pd.read_excel(u)))
        except Exception as exc:
            st.error(f"Could not read {u.name}: {exc}")
            return
    try:
        stage, per_file = core.parse_many(files, regions, branches, key, BDUE, RDUE, prior=records.get(core.prev_key(key)))
    except ValueError as exc:
        st.error(str(exc))
        return
    st.dataframe(per_file, hide_index=True, use_container_width=True)
    if stage.empty:
        st.warning("None of the files contain responses for this month.")
        return

    vc = stage["Check"].value_counts()
    c = st.columns(6)
    c[0].metric("Responses", len(stage))
    c[1].metric("OK", int(vc.get("OK", 0)))
    c[2].metric("Warnings", int(vc.get("Warning", 0)))
    c[3].metric("Errors", int(vc.get("Error", 0)))
    c[4].metric("Superseded", int(vc.get("Superseded", 0)))
    c[5].metric("Branches covered", f"{stage[stage['Include']]['Branch Code'].nunique()} / {len(branches)}")

    st.markdown("##### Review")
    view_mode = st.radio("Show", ["Needs attention (warnings, errors, superseded)", "All rows"], horizontal=True)
    attention = view_mode.startswith("Needs")
    view = stage[stage["Check"] != "OK"] if attention else stage
    st.caption("Rows with **Errors** are un-ticked automatically. Fix the figure after confirming with the RHO (edits are "
               "recorded) or leave the row out. A branch with no row is saved as 'Not Submitted'. "
               "Changing this view clears edits you have not saved.")
    ed = None
    if view.empty:
        st.success("Every response passed all checks.")
    else:
        cfgcols = {"Include": st.column_config.CheckboxColumn("Include", help="Tick to save this row"),
                   "Issues": st.column_config.TextColumn("Issues", width="large")}
        cfgcols.update(kri_column_config())
        sig = "|".join(f"{u.name}{u.size}" for u in ups)
        ed = st.data_editor(view, key=f"stage_{sig}_{key}_{'a' if attention else 'b'}", hide_index=True,
                            use_container_width=True, column_config=cfgcols, height=460,
                            disabled=[c for c in view.columns if c != "Include" and c not in NUM_COLS])
    edited = stage.copy()
    if ed is not None:
        cols_ed = ["Include"] + NUM_COLS
        edited.loc[ed.index, cols_ed] = ed[cols_ed].values

    inc = edited[edited["Include"] == True]  # noqa: E712
    problems = core.validate_kri_values(inc)
    if (inc["Branch Code"] == "").any():
        problems.append("A ticked row has no valid branch.")
    dup = inc["Branch Code"][inc["Branch Code"].duplicated() & (inc["Branch Code"] != "")]
    if len(dup):
        problems.append("More than one ticked row for branch(es): " + ", ".join(sorted(set(dup))))
    for p in problems[:12]:
        st.error(p)
    forced = inc[inc["Check"] == "Error"]
    if len(forced):
        st.warning(f"{len(forced)} row(s) flagged as Error are ticked - make sure you have verified them.")

    st.markdown("##### RHO confirmation")
    st.caption("Type the name of the Regional Manager / RHO officer who confirmed each region's file (from their e-mail).")
    reg_rows = (inc[inc["Region Code"] != ""].groupby(["Region Code", "Region Name"])
                .agg(Branches=("Branch Code", "nunique"), Files=("Source File", lambda s: ", ".join(sorted(set(s)))))
                .reset_index())
    reg_rows["RHO Confirmed By"] = ""
    conf = st.data_editor(reg_rows, key=f"conf_{'|'.join(f'{u.name}{u.size}' for u in ups)}_{key}", hide_index=True,
                          use_container_width=True, disabled=["Region Code", "Region Name", "Branches", "Files"])
    rho = {r["Region Code"]: str(r["RHO Confirmed By"] or "").strip() for r in conf.to_dict("records")}
    missing = [c for c, v in rho.items() if not v]
    if missing:
        st.warning("No RHO confirmation typed for: " + ", ".join(missing) + " (you can still save).")

    if st.button("Save to monthly records", type="primary", disabled=bool(problems) or inc.empty):
        new_rows = core.stage_to_rows(edited, original=stage, rho_confirmed=rho)
        replace = mode.startswith("Replace")
        records[key] = core.merge_month(records.get(key), new_rows, branches, replace=replace)
        orig = stage.loc[inc.index, NUM_COLS].apply(pd.to_numeric, errors="coerce").fillna(-1)
        now = inc[NUM_COLS].apply(pd.to_numeric, errors="coerce").fillna(-1)
        n_edit = int((orig.values != now.values).sum())
        log = core.append_log(core.load_log(), **{
            "Reporting Month": core.month_label(key), "Source File": "; ".join(u.name for u in ups)[:250],
            "Mode": "Replace month" if replace else "Update branches", "Rows in File": len(stage),
            "Branches Saved": inc["Branch Code"].nunique(), "Cells Edited Manually": n_edit,
            "Notes": f"{int((stage['Check'] == 'Error').sum())} error row(s); RHO confirmation missing for {len(missing)} region(s)"})
        try:
            core.save_records(records, log)
            st.success(f"Saved {inc['Branch Code'].nunique()} branch(es) from {inc['Region Code'].nunique()} region(s) "
                       f"to the {core.key_to_tab(key)} tab. {n_edit} figure(s) edited by hand.")
        except core.FileLockedError as exc:
            st.error(str(exc))


def page_report():
    regions, branches, records = load_all()
    banner("Annex I Report", "Bank-wide Present / Previous / YTD (sum of all branches) - download as Excel or PDF")
    key = pick_stored_month(records, key="rep_month")
    if key is None:
        return
    month_df = records[key]
    sel = region_filter(month_df, st, key="rep_regions")
    codes = sel or None
    annex, meta = core.annex1(records, key, codes)
    mx = core.region_matrix(records, key, codes)
    bm = core.branch_matrix(records, key, codes)
    bfull = core.month_rows(records, key, codes)
    rst = core.region_status_table(records, key, regions, branches, BDUE, RDUE)
    bst = core.branch_status_table(records, key, regions, branches, BDUE, RDUE)
    if codes:
        rst, bst = rst[rst["Region Code"].isin(codes)], bst[bst["Region Code"].isin(codes)]
    tr = core.trend_table(records, codes)
    n_regions = month_df["Region Code"].nunique()

    if codes and len(codes) < n_regions:
        st.warning(f"Filtered view: {len(codes)} of {n_regions} regions. The exported files will say "
                   "'not the official bank-wide report'. Clear the region filter for the official report.")
    n_missing = int((month_df["Status"] == cfg.STATUS_NONE).sum())
    if n_missing:
        st.warning(f"{n_missing} branch(es) have not submitted for {core.month_label(key)} - they are not in the totals. "
                   "See **Submission Tracker**.")

    st.dataframe(pd.DataFrame({
        "No.": annex["No"], "TBML Key Risk Indicator": annex["Indicator"], "Unit": annex["Unit"],
        "Present Period": [core.fmt_val(v, k) for v, k in zip(annex["Present"], annex["Kind"])],
        "Previous Period": [core.fmt_val(v, k) for v, k in zip(annex["Previous"], annex["Kind"])],
        "YTD": [core.fmt_val(v, k) for v, k in zip(annex["YTD"], annex["Kind"])],
    }), hide_index=True, use_container_width=True, height=640)
    st.caption(f"Present = {core.month_label(key)}; Previous = {core.month_label(meta['previous_key'])}. Present, Previous and YTD are "
               "the totals of the figures reported by the branches in their Forms. N/A = no branch reported that figure.")

    st.markdown("##### Sign-off details (printed on the report)")
    a, b = st.columns(2)
    prepared_by = a.text_input("Prepared by", SETTINGS.get("prepared_by", ""))
    prepared_des = a.text_input("Designation (prepared)", SETTINGS.get("prepared_designation", ""))
    prepared_date = a.date_input("Date (prepared)", value=dt.date.today())
    approved_by = b.text_input("Approved by", SETTINGS.get("approved_by", ""))
    approved_des = b.text_input("Designation (approved)", SETTINGS.get("approved_designation", ""))
    approved_date = b.date_input("Date (approved)", value=dt.date.today())

    if st.button("Prepare Excel & PDF", type="primary"):
        info = dict(settings=SETTINGS, key=key, total_regions=n_regions,
                    prepared_by=prepared_by, prepared_designation=prepared_des, prepared_date=prepared_date.strftime("%d %b %Y"),
                    approved_by=approved_by, approved_designation=approved_des, approved_date=approved_date.strftime("%d %b %Y"))
        try:
            st.session_state["export"] = {
                "key": key,
                "xlsx": exports.build_excel(key, annex, meta, mx, rst, tr, info, branch_df=bm, branch_status=bst, branch_full=bfull),
                "pdf": exports.build_pdf(key, annex, meta, mx, rst, info),
            }
        except Exception as exc:
            st.error(f"Could not build the report: {exc}")
    ex = st.session_state.get("export")
    if ex:
        st.success(f"Files ready for {core.month_label(ex['key'])}.")
        d1, d2 = st.columns(2)
        d1.download_button("Download Excel (Annex I + region + branch sheets)", ex["xlsx"],
                           file_name=f"TBML_KRI_Report_{ex['key']}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        d2.download_button("Download PDF", ex["pdf"], file_name=f"TBML_KRI_Report_{ex['key']}.pdf", mime="application/pdf")


def page_combined():
    regions, branches, records = load_all()
    banner("Combined Report", "Merge several saved months (e.g. October to December) into one master Excel file")
    keys_all = sorted(records)
    if not keys_all:
        st.info("No monthly data has been saved yet.")
        return
    a, b = st.columns(2)
    start = a.selectbox("From month", keys_all, index=0, format_func=core.month_label, key="cmb_from")
    end = b.selectbox("To month", keys_all, index=len(keys_all) - 1, format_func=core.month_label, key="cmb_to")
    if start > end:
        st.error("'From month' must not be later than 'To month'.")
        return
    keys = core.months_in_range(records, start, end)
    missing = core.missing_months(records, start, end)
    if missing:
        st.warning("No saved data for: " + ", ".join(core.month_label(k) for k in missing)
                   + ". They are left out of the combined figures.")
    sample = records[keys[-1]]
    sel = region_filter(sample, st, key="cmb_regions")
    codes = sel or None
    tables = core.combined_tables(records, keys, codes)
    if tables["flat"].empty:
        st.warning("Nothing to combine for this selection.")
        return
    st.caption(f"{len(keys)} month(s) combined: " + ", ".join(core.key_to_tab(k) for k in keys)
               + f"  |  {len(tables['flat'])} branch rows in the master table.")
    st.markdown("##### Bank-wide totals by month")
    st.dataframe(tables["bank_by_month"], hide_index=True, use_container_width=True, column_config=kri_column_config())
    st.caption("PERIOD TOTAL adds the months together, except KRI-02 and KRI-03 (period-end balances), which show the last month.")
    t1, t2, t3, t4 = st.tabs(["Region by month", "Region total", "Branch total", "Master table (all data)"])
    with t1:
        st.dataframe(tables["region_by_month"], hide_index=True, use_container_width=True, column_config=kri_column_config())
    with t2:
        st.dataframe(tables["region_total"], hide_index=True, use_container_width=True, column_config=kri_column_config())
    with t3:
        st.dataframe(tables["branch_total"], hide_index=True, use_container_width=True, column_config=kri_column_config(), height=480)
    with t4:
        st.dataframe(tables["flat"].drop(columns=["Month Key"]), hide_index=True, use_container_width=True,
                     column_config=kri_column_config(), height=480)
    if st.button("Prepare combined Excel", type="primary"):
        info = dict(settings=SETTINGS, filtered=bool(codes) and len(codes) < sample["Region Code"].nunique())
        try:
            st.session_state["combined"] = {
                "name": f"TBML_Combined_{keys[0]}_to_{keys[-1]}.xlsx",
                "xlsx": exports.build_combined_excel(keys, tables, missing, info)}
        except Exception as exc:
            st.error(f"Could not build the file: {exc}")
    cb = st.session_state.get("combined")
    if cb:
        st.success("Combined file ready.")
        st.download_button("Download combined Excel (master file)", cb["xlsx"], file_name=cb["name"],
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def page_tracker():
    regions, branches, records = load_all()
    banner("Submission Tracker", "Which regions and branches have reported, who is late, who is outstanding")
    if need_master(regions, branches):
        return
    key = pick_stored_month(records, key="trk_month")
    if key is None:
        return
    rst = core.region_status_table(records, key, regions, branches, BDUE, RDUE)
    bst = core.branch_status_table(records, key, regions, branches, BDUE, RDUE)
    total = len(bst)
    done = int((bst["Status"] != cfg.STATUS_NONE).sum())
    st.progress(done / total if total else 0.0, text=f"{done} of {total} branches have reported for {core.month_label(key)}")
    c = st.columns(5)
    c[0].metric("Submitted", int((bst["Status"] == cfg.STATUS_SUBMITTED).sum()))
    c[1].metric("Nil reports", int((bst["Status"] == cfg.STATUS_NIL).sum()))
    c[2].metric("Not submitted", int((bst["Status"] == cfg.STATUS_NONE).sum()))
    c[3].metric("After due date", int((~bst["Timeliness"].isin(["On time", "-"])).sum()))
    c[4].metric("Regions complete", f"{int((rst['Status'] == 'Complete').sum())} / {len(rst)}")

    t1, t2, t3 = st.tabs(["By region", "By branch", "Reminders"])
    with t1:
        st.dataframe(style_status(rst), hide_index=True, use_container_width=True, height=620, column_config={
            "% Reported": st.column_config.ProgressColumn("% Reported", min_value=0, max_value=100, format="%.0f%%")})
    with t2:
        a, b = st.columns(2)
        rf = a.multiselect("Regions", list(regions["Region Code"]), key="trk_reg")
        sf = b.multiselect("Status", ALL_STATUS, default=ALL_STATUS, key="trk_st")
        v = bst[bst["Status"].isin(sf)]
        if rf:
            v = v[v["Region Code"].isin(rf)]
        st.dataframe(style_status(v), hide_index=True, use_container_width=True, height=560)
    with t3:
        out = bst[bst["Status"] == cfg.STATUS_NONE]
        if out.empty:
            st.success("All branches have reported.")
        else:
            regs = rst[rst["Region Code"].isin(out["Region Code"])]
            st.markdown("**RHOs to remind** (copy into the e-mail 'To' field)")
            st.code("; ".join(e for e in regs["RHO Email"] if e), language=None)
            pick = st.selectbox("Outstanding branches of region", list(regs["Region Code"]),
                                format_func=lambda c: f"{c} - {regs.loc[regs['Region Code'] == c, 'Region Name'].iloc[0]}")
            sub = out[out["Region Code"] == pick]
            st.code("\n".join(f"{c} | {n}" for c, n in zip(sub["Branch Code"], sub["Branch Name"])), language=None)
    st.download_button("Download tracker (Excel)", to_xlsx_bytes({"By region": rst, "By branch": bst}),
                       file_name=f"TBML_Tracker_{key}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def page_records():
    regions, branches, records = load_all()
    banner("Monthly Records", "Everything saved so far - one Excel tab per month, one row per branch")
    if not records:
        st.info("Nothing saved yet.")
        return
    summary = pd.DataFrame([{
        "Month": core.month_label(k), "Tab": core.key_to_tab(k), "Branches": len(d),
        "Submitted": int((d["Status"] == cfg.STATUS_SUBMITTED).sum()),
        "Nil": int((d["Status"] == cfg.STATUS_NIL).sum()),
        "Not submitted": int((d["Status"] == cfg.STATUS_NONE).sum()),
        "Rows edited by hand": int((d["Edited"] == "Yes").sum()),
    } for k, d in sorted(records.items())])
    st.dataframe(summary, hide_index=True, use_container_width=True)
    key = st.selectbox("Open month", sorted(records, reverse=True), format_func=core.month_label, key="rec_month")
    d = records[key]
    rsel = region_filter(d, st, key="rec_regions")
    st.dataframe(d[d["Region Code"].isin(rsel)] if rsel else d, hide_index=True, use_container_width=True,
                 column_config=kri_column_config(), height=480)
    if cfg.RECORDS_FILE.exists():
        st.download_button("Download the full records workbook (all months)", cfg.RECORDS_FILE.read_bytes(),
                           file_name="TBML_Records.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption(f"Stored at: {cfg.RECORDS_FILE}  |  automatic backups: {cfg.BACKUP_DIR}")
    st.markdown("##### Upload log")
    st.dataframe(core.load_log().iloc[::-1], hide_index=True, use_container_width=True)
    with st.expander("Delete a month (cannot be undone - a backup copy is kept)"):
        ok = st.checkbox(f"I want to delete {core.month_label(key)}")
        if st.button("Delete month", disabled=not ok):
            try:
                log = core.append_log(core.load_log(), **{"Reporting Month": core.month_label(key),
                                                          "Mode": "DELETE", "Notes": "Month deleted by user"})
                core.save_records(core.delete_month(records, key), log)
                st.success("Deleted.")
                st.rerun()
            except core.FileLockedError as exc:
                st.error(str(exc))


def page_master():
    banner("Master Data", "Regions (with the RHO's e-mail) and branches")
    regions, branches = core.load_master()
    tpl = cfg.TEMPLATES_DIR / "Regions_Master_Template.xlsx"
    up = st.file_uploader("Import from the Excel template (replaces what is shown below)", type=["xlsx"])
    if up is not None:
        try:
            xl = pd.ExcelFile(up)
            if "Regions" in xl.sheet_names:
                regions = pd.read_excel(xl, sheet_name="Regions", dtype=str).fillna("")
            if "Branches" in xl.sheet_names:
                branches = pd.read_excel(xl, sheet_name="Branches", dtype=str).fillna("")
            st.info("File loaded - review the tables and press **Save master data**.")
        except Exception as exc:
            st.error(f"Could not read the file: {exc}")
    if tpl.exists():
        st.download_button("Download blank template", tpl.read_bytes(), file_name="Regions_Master_Template.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    sig = f"{len(regions)}_{len(branches)}_{up.name if up else ''}"
    t1, t2 = st.tabs(["Regions", "Branches"])
    with t1:
        r_ed = st.data_editor(core.clean_regions(regions), num_rows="dynamic", hide_index=True,
                              use_container_width=True, key=f"reg_{sig}")
        st.caption("Officer Email = the RHO / Regional Manager who owns the region's Form.")
    with t2:
        b_ed = st.data_editor(core.clean_branches(branches), num_rows="dynamic", hide_index=True,
                              use_container_width=True, key=f"br_{sig}", height=420)
        st.caption("Region Name is filled in automatically from the Region Code. Branch Email is optional "
                   "(if filled, a different sender is shown as a warning).")
    errs = core.validate_master(r_ed, b_ed)
    for e in errs:
        st.error(e)
    if st.button("Save master data", type="primary", disabled=bool(errs)):
        try:
            core.save_master(r_ed, b_ed)
            st.success("Master data saved.")
        except core.FileLockedError as exc:
            st.error(str(exc))

    rc, bc = core.clean_regions(r_ed), core.clean_branches(b_ed)
    if not rc.empty and not bc.empty and not errs:
        st.markdown("##### Dropdown lists for each region's Form")
        st.caption("Each region's Form has four dropdowns: Region, Region Code, Branch Name and Branch Code. "
                   "Choose the region, copy each list, and in the Form use **Add multiple options** to paste it. "
                   "Keep the Branch Name list and the Branch Code list in the same order.")
        pick = st.selectbox("Region", list(rc["Region Code"]),
                            format_func=lambda c: f"{c} - {rc.loc[rc['Region Code'] == c, 'Region Name'].iloc[0]}")
        sub = bc[bc["Region Code"] == pick]
        a1, a2 = st.columns(2)
        a1.markdown("**Region** (one option)")
        a1.code(rc.loc[rc["Region Code"] == pick, "Region Name"].iloc[0], language=None)
        a2.markdown("**Region Code** (one option)")
        a2.code(pick, language=None)
        b1, b2 = st.columns(2)
        b1.markdown("**Branch Name** list")
        b1.code("\n".join(sub["Branch Name"]) or "(no branches)", language=None)
        b2.markdown("**Branch Code** list")
        b2.code("\n".join(sub["Branch Code"]) or "(no branches)", language=None)
        st.download_button("Download all regions' branch lists (one sheet per region)", core.branch_lists_workbook(rc, bc),
                           file_name="Branch_Dropdown_Lists.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def page_guide():
    banner("Monthly Guide", "How the monthly cycle works (Circular No. 1996/2026)")
    st.markdown(
        f"""
**Set-up (once):** 25 Forms - one per region - each with that region's Region, Region Code, Branch Name and Branch Code lists, shared with the RHO
as collaborator. Branches get the link from their RHO and keep using the same link every month.

**Every month**
1. **1st** - send the reminder e-mail to the 25 RHOs (the *Submission Tracker -> Reminders* tab gives you the addresses).
2. **By the {core.ordinal(BDUE)}** - branches submit their figures in their region's Form (choosing the reporting month and their branch).
3. **{core.ordinal(BDUE)}-{core.ordinal(RDUE)}** - the RHO opens *Responses -> Open in Excel*, checks the branches (asks a branch to resubmit, or corrects the figure),
   saves the file and e-mails it to Channel Management with their confirmation.
4. **After the {core.ordinal(RDUE)}** - **Upload & Validate**: select all the files together, review the flags, type the RHO's name for each region, save.
5. **Submission Tracker** - chase branches/regions that are outstanding. Late files can be uploaded later ("Update only the branches found").
6. **Annex I Report** - check, fill in the sign-off names, download Excel / PDF. **By the 15th** send to IBD and Compliance.
7. **Combined Report** (any time) - pick From/To months (e.g. October to December) and download one master Excel file. You never have to merge months by hand: every saved month stays in the records and is picked up automatically.

**Rules built into the portal**
- Counts must be whole numbers, never negative. KRI-05 and KRI-08 are in Rs. Million (2 decimals), converted by the branch at the mid rate of the last day of the month.
- Region and bank-wide figures are the **sum of the branch rows**. Branches that have not submitted are not in the totals.
- A branch that submits twice: the latest is kept, earlier ones are *Superseded*.
- Responses for other months inside a file are ignored automatically.
- Each branch reports three figures per indicator (previous month, present month, year to date). Annex I shows the totals of these, as the circular requires.
- The portal also cross-checks: the branch's 'Previous Month' against what it reported as present last month, and YTD against the present month (warnings only).
"""
    )
    st.markdown("##### The 18 indicators")
    st.dataframe(pd.DataFrame([{"Code": i["code"], "Indicator": i["full"], "Unit": i["unit"], "Group": i["group"]}
                               for i in INDICATORS]), hide_index=True, use_container_width=True)
    st.markdown("##### Reporting-month list for the Form dropdown")
    st.code("\n".join(core.month_label(k) for k in core.month_options(15)), language=None)


# --------------------------------------------------------------------------- #
# Navigation
# --------------------------------------------------------------------------- #
gate()
img = f'<img src="data:image/png;base64,{LOGO_B64}"/>' if LOGO_B64 else ""
st.sidebar.markdown(
    f'<div class="sb-brand">{img}<h3>{SETTINGS["bank_name"]}</h3><p>{SETTINGS["department"]}</p></div>',
    unsafe_allow_html=True)
st.sidebar.divider()
PAGES = {
    "📊  Dashboard": page_dashboard,
    "📥  Upload & Validate": page_upload,
    "📄  Annex I Report": page_report,
    "🗂️  Combined Report": page_combined,
    "✅  Submission Tracker": page_tracker,
    "🗓️  Monthly Records": page_records,
    "⚙️  Master Data": page_master,
    "📘  Monthly Guide": page_guide,
}
choice = st.sidebar.radio("Navigate", list(PAGES), label_visibility="collapsed")
st.sidebar.divider()
st.sidebar.caption("STRICTLY CONFIDENTIAL - TBML data. Stored locally on this computer only.")
PAGES[choice]()
