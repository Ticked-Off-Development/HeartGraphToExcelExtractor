"""
HRM Pacing Dashboard
====================
Interactive Streamlit dashboard for visualising heart rate zone data,
PEM/STABLE status, and pacing patterns from the HRM Daily Excel file.

Usage:
    streamlit run hrm_dashboard.py

Then upload your HRM Daily .xlsx file via the sidebar file uploader.
"""

import io
import warnings
from datetime import date, datetime, timedelta

# openpyxl emits noisy UserWarnings for unsupported Excel extensions
# (e.g. conditional-formatting extensions).  They are harmless for our use.
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

import numpy as np
import pandas as pd
import plotly.basedatatypes as _pbd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st

# Plotly 6 encodes numpy/pandas arrays as base64 binary data ({"dtype": …,
# "bdata": …}).  Streamlit's bundled plotly.js does not support this format,
# causing every chart to render blank.  Disabling the conversion keeps data
# as plain JSON arrays, which both plotly.js 2.x and 3.x understand.
_pbd.convert_to_base64 = lambda obj: None

# ── Page configuration ────────────────────────────────────────────────────────
st.set_page_config(
    page_title="HRM Pacing Dashboard",
    page_icon="❤️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── Constants ─────────────────────────────────────────────────────────────────
ZONE_ORDER = ["red", "orange", "yellow", "green", "blue"]  # Zone 5 → Zone 1

ZONE_META = {
    "red":    {"label": "Zone 5 (Red)",    "color": "#C62828"},
    "orange": {"label": "Zone 4 (Orange)", "color": "#EF6C00"},
    "yellow": {"label": "Zone 3 (Yellow)", "color": "#F9A825"},
    "green":  {"label": "Zone 2 (Green)",  "color": "#2E7D32"},
    "blue":   {"label": "Zone 1 (Blue)",   "color": "#1565C0"},
}

PEM_SHADE = "rgba(198, 40, 40, 0.12)"

# ── Utility helpers ───────────────────────────────────────────────────────────

def hex_to_rgba(hex_color: str, alpha: float = 0.15) -> str:
    """Convert a #RRGGBB hex string to an rgba() CSS colour string."""
    h = hex_color.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"


def excel_time_to_seconds(val) -> float:
    """Convert an Excel time value (any form openpyxl may return) to seconds."""
    if val is None:
        return 0.0
    try:
        if isinstance(val, float):
            return val * 86400 if not np.isnan(val) else 0.0
        if isinstance(val, int):
            return float(val) * 86400
        if isinstance(val, timedelta):
            return val.total_seconds()
        if isinstance(val, datetime):
            # openpyxl sometimes wraps a time as a datetime on the epoch date
            return val.hour * 3600 + val.minute * 60 + val.second
        if hasattr(val, "hour"):          # datetime.time
            return val.hour * 3600 + val.minute * 60 + val.second
    except Exception:
        pass
    return 0.0


def seconds_to_hms(seconds) -> str:
    """Format a total-seconds number as H:MM:SS."""
    if seconds is None or (isinstance(seconds, float) and np.isnan(seconds)):
        return "0:00:00"
    s = int(seconds)
    h, rem = divmod(s, 3600)
    m, sec = divmod(rem, 60)
    return f"{h}:{m:02d}:{sec:02d}"


def detect_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """Return the first df column whose lowercase name contains any candidate."""
    for col in df.columns:
        lc = str(col).lower().strip()
        for c in candidates:
            if c in lc:
                return col
    return None


def add_pem_shading(fig: go.Figure, pem_dates):
    """Add semi-transparent red vertical rectangles for each PEM day.

    All shapes are added in a single update_layout() call.  The previous
    approach called add_vrect() once per PEM day, and each call deepcopies
    the entire figure — with 400+ PEM days that took 40+ seconds.
    """
    pem_list = list(pem_dates)
    if not pem_list:
        return
    existing = list(fig.layout.shapes) if fig.layout.shapes else []
    new_shapes = [
        dict(
            type="rect",
            x0=(d - timedelta(hours=12)).strftime("%Y-%m-%dT%H:%M:%S"),
            x1=(d + timedelta(hours=12)).strftime("%Y-%m-%dT%H:%M:%S"),
            y0=0, y1=1,
            xref="x", yref="paper",
            fillcolor=PEM_SHADE,
            layer="below",
            line_width=0,
        )
        for d in pem_list
    ]
    fig.update_layout(shapes=existing + new_shapes)


# ── Data loading ──────────────────────────────────────────────────────────────

@st.cache_data(show_spinner="Loading HRM data…")
def load_data(file_bytes: bytes) -> pd.DataFrame:
    """Read the HRM Daily sheet and return a processed DataFrame."""
    buf = io.BytesIO(file_bytes)
    raw = pd.read_excel(buf, sheet_name="HRM Daily", header=1, engine="openpyxl")

    df = raw.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)

    # ── Date ──────────────────────────────────────────────────────────────────
    date_col = df.columns[0]

    # Step 1: try letting pandas auto-detect (handles Excel datetime objects,
    # ISO strings, and many other common formats).
    date_series = pd.to_datetime(df[date_col], errors="coerce")

    # Step 2: if most rows are still NaT, the column likely contains text in
    # the "Day, Month D, YYYY" format used in the HRM Daily sheet — try that.
    if date_series.isna().mean() > 0.5:
        date_series = pd.to_datetime(
            df[date_col].astype(str).str.strip(),
            format="%A, %B %d, %Y",
            errors="coerce",
        )

    df["date"] = date_series
    df = df.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)

    # ── Zone times → seconds, fractional hours, % of session ─────────────────
    for z in ZONE_ORDER:
        src_col = detect_column(df, [z])
        df[f"{z}_sec"] = df[src_col].apply(excel_time_to_seconds) if src_col else 0.0

    zone_sec_cols = [f"{z}_sec" for z in ZONE_ORDER]
    df["total_sec"] = df[zone_sec_cols].sum(axis=1)

    for z in ZONE_ORDER:
        df[f"{z}_hr"] = df[f"{z}_sec"] / 3600
        df[f"{z}_pct"] = np.where(
            df["total_sec"] > 0,
            df[f"{z}_sec"] / df["total_sec"] * 100,
            np.nan,
        )

    df["total_hr"] = df["total_sec"] / 3600

    # Safe zone = Zone 1 (blue) + Zone 2 (green)
    df["safe_pct"] = np.where(
        df["total_sec"] > 0,
        (df["green_sec"] + df["blue_sec"]) / df["total_sec"] * 100,
        np.nan,
    )

    # ── HR columns ────────────────────────────────────────────────────────────
    max_col = detect_column(df, ["max hr", "max_hr"])
    if max_col:
        df["max_hr"] = pd.to_numeric(df[max_col], errors="coerce")

    mean_col = detect_column(df, ["mean 24", "mean_24", "mean hr"])
    if mean_col:
        df["mean_hr"] = pd.to_numeric(df[mean_col], errors="coerce")

    # ── 7-day average HR ──────────────────────────────────────────────────────
    avg7_col = detect_column(df, ["7-day", "7 day", "7day"])
    if avg7_col:
        df["hr_7d"] = pd.to_numeric(df[avg7_col], errors="coerce")
    elif "mean_hr" in df.columns:
        df["hr_7d"] = df["mean_hr"].rolling(7, min_periods=1).mean().round(1)

    # ── Tags / PEM status ─────────────────────────────────────────────────────
    tags_col = detect_column(df, ["tag", "status"])
    df["tag"] = (
        df[tags_col].fillna("").astype(str).str.strip()
        if tags_col
        else pd.Series("", index=df.index)
    )
    df["is_pem"] = df["tag"].str.upper().str.contains("PEM", na=False)

    # ── Other Events ──────────────────────────────────────────────────────────
    evt_col = detect_column(df, ["event"])
    df["events"] = (
        df[evt_col].fillna("").astype(str).str.strip()
        if evt_col
        else pd.Series("", index=df.index)
    )

    # ── Drop date-only rows with no meaningful data ────────────────────────────
    # Keep a row only if it has at least one of: zone time, mean HR, max HR,
    # or a non-empty status tag.  Rows that only have a date (e.g. future dates
    # pre-filled in the sheet, or blank filler rows) are skipped.
    has_zone    = df["total_sec"] > 0
    has_mean_hr = df["mean_hr"].notna() if "mean_hr" in df.columns else pd.Series(False, index=df.index)
    has_max_hr  = df["max_hr"].notna()  if "max_hr"  in df.columns else pd.Series(False, index=df.index)
    has_tag     = df["tag"] != ""

    df = df[has_zone | has_mean_hr | has_max_hr | has_tag].reset_index(drop=True)

    return df


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.title("❤️ HRM Pacing")
    st.caption("Heart Rate Monitor · Pacing Dashboard")

    uploaded = st.file_uploader(
        "Upload HRM Daily Excel file", type=["xlsx", "xls"]
    )

    if uploaded is None:
        st.info("Upload your **HRM Daily .xlsx** file to begin.")
        st.stop()

    df_full = load_data(uploaded.getvalue())

    st.divider()

    # Date range
    min_date = df_full["date"].min().date()
    max_date = df_full["date"].max().date()
    today = date.today()

    preset = st.selectbox(
        "Date range",
        [
            "Last 30 days", "Last 60 days", "Last 90 days",
            "Last 6 months", "This year", "All time", "Custom",
        ],
        index=2,
    )

    if preset == "Last 30 days":
        d_start, d_end = today - timedelta(days=30), today
    elif preset == "Last 60 days":
        d_start, d_end = today - timedelta(days=60), today
    elif preset == "Last 90 days":
        d_start, d_end = today - timedelta(days=90), today
    elif preset == "Last 6 months":
        d_start, d_end = today - timedelta(days=182), today
    elif preset == "This year":
        d_start, d_end = date(today.year, 1, 1), today
    elif preset == "All time":
        d_start, d_end = min_date, max_date
    else:
        d_start = st.date_input("From", value=min_date, min_value=min_date, max_value=max_date)
        d_end   = st.date_input("To",   value=max_date, min_value=min_date, max_value=max_date)

    d_start = max(d_start, min_date)
    d_end   = min(d_end,   max_date)

    st.divider()

    # Tag filter
    all_tags_raw = df_full["tag"].unique().tolist()
    display_tags = [t if t != "" else "Untagged" for t in all_tags_raw]
    tag_display_map = {(t if t != "" else "Untagged"): t for t in all_tags_raw}

    selected_display = st.multiselect(
        "Filter by status tag",
        options=sorted(display_tags),
        default=sorted(display_tags),
    )
    selected_raw = [tag_display_map[d] for d in selected_display]

# ── Apply filters ─────────────────────────────────────────────────────────────

df = df_full[
    (df_full["date"].dt.date >= d_start)
    & (df_full["date"].dt.date <= d_end)
    & (df_full["tag"].isin(selected_raw))
].copy()

if df.empty:
    st.warning("No data for the selected date range / tag filter.")
    st.stop()

pem_dates  = df[df["is_pem"]]["date"]
pem_days   = int(df["is_pem"].sum())
stable_days = int((df["tag"].str.upper() == "STABLE").sum())

# ── Top KPI row ───────────────────────────────────────────────────────────────

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Days Tracked", f"{len(df)}")
k2.metric(
    "PEM Days",
    f"{pem_days}",
    delta=f"{pem_days / len(df) * 100:.0f}% of period" if len(df) else None,
    delta_color="inverse",
)
k3.metric("STABLE Days", f"{stable_days}")
k4.metric(
    "Avg Mean HR",
    f"{df['mean_hr'].mean():.1f} bpm" if "mean_hr" in df.columns and df["mean_hr"].notna().any() else "—",
)
k5.metric(
    "Avg Safe Zone %",
    f"{df['safe_pct'].mean():.1f}%" if df["safe_pct"].notna().any() else "—",
)

st.divider()

# ── Debug expander ────────────────────────────────────────────────────────────

with st.expander("🔍 Data debug info (expand if charts are blank)", expanded=False):
    st.markdown(f"**Rows loaded:** {len(df_full)}  |  **After date filter:** {len(df)}")
    st.markdown(f"**Date range in file:** {df_full['date'].min().date()} → {df_full['date'].max().date()}")

    st.markdown("**Columns found in Excel:**")
    st.code(", ".join(df_full.columns.tolist()))

    zone_detection = {}
    for z in ZONE_ORDER:
        col = detect_column(df_full, [z])
        zone_detection[z] = col or "❌ NOT FOUND"
    st.markdown("**Zone column mapping:**")
    st.json(zone_detection)

    hr_cols = {
        "max_hr": "max_hr" in df_full.columns,
        "mean_hr": "mean_hr" in df_full.columns,
        "hr_7d": "hr_7d" in df_full.columns,
    }
    st.markdown("**HR columns:**")
    st.json(hr_cols)

    st.markdown("**Sample data (first 3 rows):**")
    preview_cols = ["date"] + [f"{z}_sec" for z in ZONE_ORDER]
    if "mean_hr" in df_full.columns:
        preview_cols.append("mean_hr")
    st.dataframe(df_full[preview_cols].head(3), width="stretch")

st.divider()

# ── Tabs ──────────────────────────────────────────────────────────────────────

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Overview",
    "🔵 Zone Analysis",
    "📈 HR Trends",
    "🛡️ Pacing & PEM",
    "💊 Events",
])


# ═════════════════════════════════════════════════════════════════════════════
# TAB 1 — Overview
# ═════════════════════════════════════════════════════════════════════════════

with tab1:
    col_left, col_right = st.columns([2, 1])

    with col_left:
        st.subheader("Mean HR over time")
        fig_ov = go.Figure()
        add_pem_shading(fig_ov, pem_dates)

        if "mean_hr" in df.columns:
            fig_ov.add_trace(go.Scatter(
                x=df["date"], y=df["mean_hr"],
                mode="lines+markers",
                name="Mean 24hr HR",
                line=dict(color="#1565C0", width=1.5),
                marker=dict(size=4),
                customdata=np.stack([df["tag"], df["events"]], axis=-1),
                hovertemplate=(
                    "%{x|%a %d %b %Y}<br>"
                    "Mean HR: %{y:.1f} bpm<br>"
                    "Tag: %{customdata[0]}<br>"
                    "Events: %{customdata[1]}<extra></extra>"
                ),
            ))

        if "hr_7d" in df.columns:
            fig_ov.add_trace(go.Scatter(
                x=df["date"], y=df["hr_7d"],
                mode="lines",
                name="7-day avg",
                line=dict(color="#EF6C00", width=2.5, dash="dot"),
                hovertemplate="%{x|%a %d %b %Y}<br>7-day avg: %{y:.1f} bpm<extra></extra>",
            ))

        fig_ov.update_layout(
            height=300,
            margin=dict(l=0, r=0, t=10, b=0),
            hovermode="x unified",
            yaxis_title="bpm",
            xaxis_showgrid=False,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        st.plotly_chart(fig_ov, width="stretch")

    with col_right:
        st.subheader("Avg zone distribution")
        avg_zone_hrs = [df[f"{z}_hr"].mean() for z in ZONE_ORDER]
        zone_labels  = [ZONE_META[z]["label"] for z in ZONE_ORDER]
        zone_colors  = [ZONE_META[z]["color"] for z in ZONE_ORDER]

        fig_donut = go.Figure(go.Pie(
            labels=zone_labels,
            values=avg_zone_hrs,
            marker_colors=zone_colors,
            hole=0.45,
            direction="clockwise",
            textinfo="percent",
            hovertemplate="%{label}<br>%{value:.2f} hrs avg<br>%{percent}<extra></extra>",
        ))
        fig_donut.update_layout(
            height=300,
            margin=dict(l=0, r=10, t=10, b=0),
            showlegend=True,
            legend=dict(orientation="v", x=0.75),
        )
        st.plotly_chart(fig_donut, width="stretch")

    # Recent 14 days table
    st.subheader("Recent days (last 14)")
    recent = df.tail(14).copy()
    recent["Date"] = recent["date"].dt.strftime("%a %d %b %Y")
    recent["Tag"]  = recent["tag"]
    for z in ZONE_ORDER:
        recent[ZONE_META[z]["label"]] = recent[f"{z}_sec"].apply(seconds_to_hms)
    if "max_hr" in df.columns:
        recent["Max HR"] = recent["max_hr"].apply(
            lambda x: f"{int(x)} bpm" if pd.notna(x) else "—"
        )
    if "mean_hr" in df.columns:
        recent["Mean HR"] = recent["mean_hr"].apply(
            lambda x: f"{x:.1f} bpm" if pd.notna(x) else "—"
        )

    display_cols = ["Date", "Tag"] + [ZONE_META[z]["label"] for z in ZONE_ORDER]
    if "max_hr" in df.columns:
        display_cols.append("Max HR")
    if "mean_hr" in df.columns:
        display_cols.append("Mean HR")

    st.dataframe(
        recent[display_cols].set_index("Date"),
        width="stretch",
    )


# ═════════════════════════════════════════════════════════════════════════════
# TAB 2 — Zone Analysis
# ═════════════════════════════════════════════════════════════════════════════

with tab2:
    view_mode = st.radio(
        "Display as", ["Hours", "% of session"],
        horizontal=True, key="zone_view",
    )
    y_suffix = "_hr" if view_mode == "Hours" else "_pct"
    y_label  = "Hours in zone" if view_mode == "Hours" else "% of session"

    # Stacked bar — daily zone time
    st.subheader("Daily zone time — stacked")
    fig_bar = go.Figure()

    for z in reversed(ZONE_ORDER):   # reversed → Zone 5 visually on top of stack
        y_vals = df[f"{z}{y_suffix}"]
        fig_bar.add_trace(go.Bar(
            x=df["date"],
            y=y_vals,
            name=ZONE_META[z]["label"],
            marker_color=ZONE_META[z]["color"],
            customdata=np.stack([
                df["tag"],
                df[f"{z}_sec"].apply(seconds_to_hms),
                df[f"{z}_pct"].round(1),
            ], axis=-1),
            hovertemplate=(
                "%{x|%a %d %b %Y}<br>"
                + ZONE_META[z]["label"] + ": "
                + ("%{customdata[1]} (%{customdata[2]}%)"
                   if view_mode == "Hours"
                   else "%{y:.1f}% (%{customdata[1]})")
                + "<br>Tag: %{customdata[0]}<extra></extra>"
            ),
        ))

    fig_bar.update_layout(
        barmode="stack",
        height=360,
        margin=dict(l=0, r=0, t=10, b=0),
        yaxis_title=y_label,
        xaxis_showgrid=False,
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    st.plotly_chart(fig_bar, width="stretch")

    # Individual zone drill-down
    st.subheader("Drill-down: individual zone trend")
    col_drill, col_stats = st.columns([3, 1])

    with col_drill:
        selected_zone = st.selectbox(
            "Select zone",
            options=ZONE_ORDER,
            format_func=lambda z: ZONE_META[z]["label"],
            key="zone_drilldown",
        )
        z_color = ZONE_META[selected_zone]["color"]
        z_fill  = hex_to_rgba(z_color, 0.15)

        fig_zone = go.Figure()
        add_pem_shading(fig_zone, pem_dates)

        # 7-day rolling for selected zone
        zone_7d = df[f"{selected_zone}{y_suffix}"].rolling(7, min_periods=1).mean()

        fig_zone.add_trace(go.Scatter(
            x=df["date"],
            y=df[f"{selected_zone}{y_suffix}"],
            mode="lines+markers",
            name=ZONE_META[selected_zone]["label"],
            line=dict(color=z_color, width=2),
            marker=dict(size=5),
            fill="tozeroy",
            fillcolor=z_fill,
            hovertemplate=(
                "%{x|%a %d %b %Y}<br>"
                + ZONE_META[selected_zone]["label"]
                + ": %{y:.2f}<extra></extra>"
            ),
        ))
        fig_zone.add_trace(go.Scatter(
            x=df["date"],
            y=zone_7d,
            mode="lines",
            name="7-day avg",
            line=dict(color="#555", width=2, dash="dot"),
            hovertemplate="%{x|%a %d %b %Y}<br>7-day avg: %{y:.2f}<extra></extra>",
        ))
        fig_zone.update_layout(
            height=280,
            margin=dict(l=0, r=0, t=10, b=0),
            yaxis_title=y_label,
            xaxis_showgrid=False,
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig_zone, width="stretch")

    with col_stats:
        st.markdown("**Zone summary**")
        stats_rows = []
        for z in ZONE_ORDER:
            total_s = df[f"{z}_sec"].sum()
            stats_rows.append({
                "Zone":  ZONE_META[z]["label"],
                "Avg":   f"{df[f'{z}_hr'].mean():.2f} h",
                "Max":   f"{df[f'{z}_hr'].max():.2f} h",
                "Total": seconds_to_hms(total_s),
            })
        st.dataframe(
            pd.DataFrame(stats_rows).set_index("Zone"),
            width="stretch",
        )


# ═════════════════════════════════════════════════════════════════════════════
# TAB 3 — HR Trends
# ═════════════════════════════════════════════════════════════════════════════

with tab3:
    has_mean = "mean_hr" in df.columns and df["mean_hr"].notna().any()
    has_max  = "max_hr"  in df.columns and df["max_hr"].notna().any()

    if not has_mean and not has_max:
        st.info("No heart-rate data found for the selected range.")
    else:
        st.subheader("Heart rate over time")

        fig_hr = make_subplots(
            rows=2, cols=1,
            shared_xaxes=True,
            row_heights=[0.58, 0.42],
            vertical_spacing=0.08,
            subplot_titles=("Mean 24hr HR", "Max HR"),
        )

        # PEM shading — yref="paper" spans both subplot rows in one call
        add_pem_shading(fig_hr, pem_dates)

        if has_mean:
            fig_hr.add_trace(go.Scatter(
                x=df["date"], y=df["mean_hr"],
                mode="lines+markers",
                name="Mean HR",
                line=dict(color="#1565C0", width=1.5),
                marker=dict(size=4),
                hovertemplate="%{x|%a %d %b %Y}<br>Mean HR: %{y:.1f} bpm<extra></extra>",
            ), row=1, col=1)

            if "hr_7d" in df.columns:
                fig_hr.add_trace(go.Scatter(
                    x=df["date"], y=df["hr_7d"],
                    mode="lines",
                    name="7-day avg",
                    line=dict(color="#EF6C00", width=2.5, dash="dot"),
                    hovertemplate="%{x|%a %d %b %Y}<br>7-day avg: %{y:.1f} bpm<extra></extra>",
                ), row=1, col=1)

        if has_max:
            fig_hr.add_trace(go.Scatter(
                x=df["date"], y=df["max_hr"],
                mode="lines+markers",
                name="Max HR",
                line=dict(color="#C62828", width=1.5),
                marker=dict(size=4),
                hovertemplate="%{x|%a %d %b %Y}<br>Max HR: %{y} bpm<extra></extra>",
            ), row=2, col=1)

        fig_hr.update_yaxes(title_text="bpm", row=1, col=1)
        fig_hr.update_yaxes(title_text="bpm", row=2, col=1)
        fig_hr.update_layout(
            height=500,
            margin=dict(l=0, r=0, t=40, b=0),
            hovermode="x unified",
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=1.04, xanchor="right", x=1),
        )
        st.plotly_chart(fig_hr, width="stretch")

        # HR summary by tag
        if df["tag"].replace("", pd.NA).notna().any() and df["tag"].nunique() > 1:
            st.subheader("HR summary by status tag")
            agg_dict = {"date": "count"}
            if has_mean:
                agg_dict["mean_hr"] = lambda x: round(x.mean(), 1)
            if has_max:
                agg_dict["max_hr"] = lambda x: round(x.mean(), 1)

            tag_stats = (
                df[df["tag"] != ""]
                .groupby("tag")
                .agg(agg_dict)
                .rename(columns={
                    "date": "Days",
                    "mean_hr": "Avg Mean HR",
                    "max_hr": "Avg Max HR",
                })
            )
            st.dataframe(tag_stats, width="stretch")


# ═════════════════════════════════════════════════════════════════════════════
# TAB 4 — Pacing & PEM
# ═════════════════════════════════════════════════════════════════════════════

with tab4:
    col_trend, col_metrics = st.columns([2, 1])

    with col_trend:
        st.subheader("Safe zone % over time  (Zone 1 + Zone 2)")
        df["safe_pct_7d"] = df["safe_pct"].rolling(7, min_periods=1).mean()

        fig_safe = go.Figure()
        add_pem_shading(fig_safe, pem_dates)

        fig_safe.add_trace(go.Scatter(
            x=df["date"], y=df["safe_pct"],
            mode="markers",
            name="Daily safe %",
            marker=dict(color="#2E7D32", size=5, opacity=0.55),
            hovertemplate="%{x|%a %d %b %Y}<br>Safe zone: %{y:.1f}%<extra></extra>",
        ))
        fig_safe.add_trace(go.Scatter(
            x=df["date"], y=df["safe_pct_7d"],
            mode="lines",
            name="7-day avg",
            line=dict(color="#1565C0", width=2.5),
            hovertemplate="%{x|%a %d %b %Y}<br>7-day avg: %{y:.1f}%<extra></extra>",
        ))
        fig_safe.update_layout(
            height=280,
            margin=dict(l=0, r=0, t=10, b=0),
            yaxis_title="% of session in Zones 1–2",
            xaxis_showgrid=False,
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig_safe, width="stretch")

    with col_metrics:
        st.subheader("Pacing summary")
        overall_safe = df["safe_pct"].mean()
        if not np.isnan(overall_safe):
            st.metric("Overall avg safe %", f"{overall_safe:.1f}%")

        pem_safe = df[df["is_pem"]]["safe_pct"].mean()
        if pem_days > 0 and not np.isnan(pem_safe):
            st.metric(
                "On PEM days",
                f"{pem_safe:.1f}%",
                delta=f"{pem_safe - overall_safe:+.1f}%",
                delta_color="normal",
            )

        stable_mask = df["tag"].str.upper() == "STABLE"
        stable_safe = df[stable_mask]["safe_pct"].mean()
        if stable_days > 0 and not np.isnan(stable_safe):
            st.metric(
                "On STABLE days",
                f"{stable_safe:.1f}%",
                delta=f"{stable_safe - overall_safe:+.1f}%",
                delta_color="normal",
            )

    # Zone comparison by tag (grouped bar)
    tagged_df = df[df["tag"] != ""]
    if tagged_df["tag"].nunique() > 1:
        st.subheader("Average zone time by status tag")
        tag_zone_rows = []
        for tag_val in sorted(tagged_df["tag"].unique()):
            subset = tagged_df[tagged_df["tag"] == tag_val]
            for z in ZONE_ORDER:
                tag_zone_rows.append({
                    "Tag":  tag_val,
                    "Zone": ZONE_META[z]["label"],
                    "Avg hours": round(subset[f"{z}_hr"].mean(), 3),
                    "color": ZONE_META[z]["color"],
                })
        tz_df = pd.DataFrame(tag_zone_rows)

        fig_tz = go.Figure()
        for z in ZONE_ORDER:
            subset = tz_df[tz_df["Zone"] == ZONE_META[z]["label"]]
            fig_tz.add_trace(go.Bar(
                x=subset["Tag"],
                y=subset["Avg hours"],
                name=ZONE_META[z]["label"],
                marker_color=ZONE_META[z]["color"],
                hovertemplate="Tag: %{x}<br>" + ZONE_META[z]["label"] + ": %{y:.2f} h<extra></extra>",
            ))
        fig_tz.update_layout(
            barmode="group",
            height=320,
            margin=dict(l=0, r=0, t=10, b=0),
            yaxis_title="Avg hours",
            xaxis_showgrid=False,
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        )
        st.plotly_chart(fig_tz, width="stretch")

    # Monthly breakdown
    st.subheader("Monthly summary")
    df_m = df.copy()
    df_m["month"] = df_m["date"].dt.to_period("M").astype(str)
    agg_m: dict = {
        "date":     ("date", "count"),
        "PEM Days": ("is_pem", "sum"),
    }
    if "mean_hr" in df.columns:
        agg_m["Avg Mean HR"] = ("mean_hr", lambda x: round(x.mean(), 1))
    agg_m["Avg Safe %"] = ("safe_pct", lambda x: round(x.mean(), 1))

    monthly = df_m.groupby("month").agg(**agg_m).rename(columns={"date": "Days"})
    monthly["PEM Days"] = monthly["PEM Days"].astype(int)
    monthly["PEM %"] = (monthly["PEM Days"] / monthly["Days"] * 100).round(1)
    st.dataframe(monthly, width="stretch")

    # Pre-PEM zone pattern
    if pem_days >= 3:
        st.subheader("Pre-PEM zone pattern (avg 3 days before PEM vs baseline)")
        pem_idx_list = df[df["is_pem"]].index.tolist()
        pre_rows = []
        for idx in pem_idx_list:
            window = df.loc[max(0, idx - 3): idx - 1]
            if len(window) > 0:
                pre_rows.append(window)

        if pre_rows:
            pre_df   = pd.concat(pre_rows)
            base_df  = df[~df["is_pem"]]
            compare  = []
            for z in ZONE_ORDER:
                pre_avg  = pre_df[f"{z}_hr"].mean()
                base_avg = base_df[f"{z}_hr"].mean()
                compare.append({
                    "Zone":              ZONE_META[z]["label"],
                    "Pre-PEM avg (h)":   round(pre_avg, 3),
                    "Baseline avg (h)":  round(base_avg, 3),
                    "Difference":        round(pre_avg - base_avg, 3),
                })
            cmp_df = pd.DataFrame(compare).set_index("Zone")
            st.dataframe(cmp_df, width="stretch")


# ═════════════════════════════════════════════════════════════════════════════
# TAB 5 — Events
# ═════════════════════════════════════════════════════════════════════════════

with tab5:
    events_df = df[df["events"] != ""].copy()

    if events_df.empty:
        st.info(
            "No **Other Events** data found for the selected range. "
            "Add events in the 'Other Events' column of your Excel file."
        )
    else:
        st.subheader("Mean HR with events annotated")

        fig_evt = go.Figure()
        add_pem_shading(fig_evt, pem_dates)

        if "mean_hr" in df.columns:
            fig_evt.add_trace(go.Scatter(
                x=df["date"], y=df["mean_hr"],
                mode="lines",
                name="Mean HR",
                line=dict(color="#1565C0", width=1.5),
                hovertemplate="%{x|%a %d %b %Y}<br>Mean HR: %{y:.1f} bpm<extra></extra>",
            ))

        # Color-code event types
        unique_events = (
            events_df["events"]
            .str.split(",")
            .explode()
            .str.strip()
            .unique()
            .tolist()
        )
        palette = px.colors.qualitative.Set2
        evt_color_map = {e: palette[i % len(palette)] for i, e in enumerate(unique_events)}

        for _, row in events_df.iterrows():
            evt_name = row["events"].strip()
            color    = evt_color_map.get(evt_name, "#888888")
            fig_evt.add_vline(
                x=row["date"].strftime("%Y-%m-%dT%H:%M:%S"),
                line_width=2,
                line_dash="dash",
                line_color=color,
                annotation_text=evt_name,
                annotation_position="top right",
                annotation_font_size=10,
                annotation_font_color=color,
            )

        fig_evt.update_layout(
            height=340,
            margin=dict(l=0, r=0, t=10, b=0),
            yaxis_title="bpm",
            xaxis_showgrid=False,
            hovermode="x unified",
        )
        st.plotly_chart(fig_evt, width="stretch")

        # Event log table
        st.subheader("Event log")
        log_cols = ["date", "tag", "events"]
        if "mean_hr" in df.columns:
            log_cols.append("mean_hr")

        evt_log = events_df[log_cols].copy()
        evt_log["date"] = evt_log["date"].dt.strftime("%a %d %b %Y")
        rename_map = {"date": "Date", "tag": "Tag", "events": "Events"}
        if "mean_hr" in evt_log.columns:
            rename_map["mean_hr"] = "Mean HR"
        evt_log = evt_log.rename(columns=rename_map).set_index("Date")
        st.dataframe(evt_log, width="stretch")

        # Before / after 7-day window comparison
        if "mean_hr" in df.columns:
            st.subheader("Before / After comparison (±7-day mean HR)")
            evt_types = (
                events_df["events"]
                .str.split(",")
                .explode()
                .str.strip()
                .unique()
            )
            ba_rows = []
            for evt in evt_types:
                evt_dates = events_df[
                    events_df["events"].str.contains(evt, na=False, regex=False)
                ]["date"].tolist()
                before_vals, after_vals = [], []
                for d in evt_dates:
                    before = df[
                        (df["date"] >= d - timedelta(days=7)) & (df["date"] < d)
                    ]["mean_hr"].dropna()
                    after = df[
                        (df["date"] > d) & (df["date"] <= d + timedelta(days=7))
                    ]["mean_hr"].dropna()
                    before_vals.extend(before.tolist())
                    after_vals.extend(after.tolist())

                ba_rows.append({
                    "Event":            evt,
                    "Occurrences":      len(evt_dates),
                    "Avg HR 7d Before": round(np.mean(before_vals), 1) if before_vals else None,
                    "Avg HR 7d After":  round(np.mean(after_vals),  1) if after_vals  else None,
                })

            if ba_rows:
                ba_df = pd.DataFrame(ba_rows).set_index("Event")
                ba_df["Delta"] = (
                    ba_df["Avg HR 7d After"] - ba_df["Avg HR 7d Before"]
                ).round(1)
                st.dataframe(ba_df, width="stretch")
