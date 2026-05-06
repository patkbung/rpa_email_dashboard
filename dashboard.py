"""
dashboard.py — RPA Email Attachment Organizer · Daily Dashboard
Neumorphism (Soft UI) theme — light gray + coral pink accents

Run with:
    streamlit run dashboard.py
"""

import os
import pandas as pd
import streamlit as st
import altair as alt
from datetime import datetime, date, timedelta

from config import REPORT_FILE, REPORT_COLUMNS

# ── Page config (must be first Streamlit call) ────────────────────────────────
st.set_page_config(
    page_title="RPA Email Attachment Dashboard",
    page_icon="📬",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════════════════════════
# CUSTOM CSS — Neumorphism Soft UI
# ══════════════════════════════════════════════════════════════════════════════
ACCENT      = "#f4896b"   # coral / salmon
ACCENT_SOFT = "#ffd6cc"   # light coral tint
BG          = "#e8ecf1"   # main background
SURFACE     = "#e8ecf1"   # same as bg (neumorphic elements share bg color)
SHADOW_DARK = "#b2bac5"
SHADOW_LITE = "#ffffff"
TEXT_MAIN   = "#3a4354"
TEXT_SUB    = "#7b8498"

st.markdown(f"""
<style>
/* ── Google Font ── */
@import url('https://fonts.googleapis.com/css2?family=Nunito:wght@300;400;600;700;800&display=swap');

html, body, [class*="css"] {{
    font-family: 'Nunito', sans-serif;
    color: {TEXT_MAIN};
}}

/* ── App background ── */
.stApp {{
    background-color: {BG};
    color: {TEXT_MAIN};
}}

/* ── Remove default streamlit elements ── */
#MainMenu, footer, header {{ visibility: hidden; }}
.block-container {{
    padding-top: 1.5rem;
    padding-bottom: 2rem;
    max-width: 1400px;
}}

/* ── Sidebar ── */
section[data-testid="stSidebar"] {{
    background-color: {BG};
    border-right: none;
    box-shadow: 6px 0 20px rgba(0,0,0,0.08);
}}
section[data-testid="stSidebar"] * {{
    color: {TEXT_MAIN} !important;
}}

/* ── Sidebar neumorph inputs ── */
section[data-testid="stSidebar"] .stDateInput > div > div,
section[data-testid="stSidebar"] .stSelectbox > div > div {{
    background: {BG} !important;
    border: none !important;
    border-radius: 12px !important;
    box-shadow: 4px 4px 8px {SHADOW_DARK}, -4px -4px 8px {SHADOW_LITE} !important;
    color: {TEXT_MAIN} !important;
}}

/* ── Streamlit buttons → neumorph style ── */
.stButton > button {{
    background: {BG};
    color: {TEXT_MAIN};
    border: none;
    border-radius: 14px;
    padding: 0.55rem 1.4rem;
    font-weight: 600;
    font-size: 0.9rem;
    box-shadow: 5px 5px 10px {SHADOW_DARK}, -5px -5px 10px {SHADOW_LITE};
    transition: all 0.18s ease;
    cursor: pointer;
}}
.stButton > button:hover {{
    box-shadow: 3px 3px 6px {SHADOW_DARK}, -3px -3px 6px {SHADOW_LITE};
    color: {ACCENT};
}}
.stButton > button:active {{
    box-shadow: inset 3px 3px 6px {SHADOW_DARK}, inset -3px -3px 6px {SHADOW_LITE};
}}

/* ── Dataframe ── */
.stDataFrame {{
    background: {BG};
    border-radius: 18px;
    box-shadow: 6px 6px 14px {SHADOW_DARK}, -6px -6px 14px {SHADOW_LITE};
    overflow: hidden;
}}

/* ── Metric widget (native) ── */
[data-testid="metric-container"] {{
    background: {BG};
    border-radius: 18px;
    padding: 1rem 1.2rem;
    box-shadow: 6px 6px 14px {SHADOW_DARK}, -6px -6px 14px {SHADOW_LITE};
}}

/* ── Scrollbar ── */
::-webkit-scrollbar {{ width: 6px; height: 6px; }}
::-webkit-scrollbar-track {{ background: {BG}; }}
::-webkit-scrollbar-thumb {{
    background: {SHADOW_DARK};
    border-radius: 4px;
}}

/* ─────────────── CUSTOM COMPONENTS ─────────────── */

/* Hero card */
.neu-hero {{
    background: {BG};
    border-radius: 24px;
    padding: 2rem 2.5rem;
    margin-bottom: 1.8rem;
    box-shadow: 10px 10px 20px {SHADOW_DARK}, -10px -10px 20px {SHADOW_LITE};
    display: flex;
    align-items: center;
    gap: 1.5rem;
}}
.hero-icon {{
    font-size: 2.8rem;
    width: 72px; height: 72px;
    display: flex; align-items: center; justify-content: center;
    background: {BG};
    border-radius: 20px;
    box-shadow: 6px 6px 12px {SHADOW_DARK}, -6px -6px 12px {SHADOW_LITE};
    flex-shrink: 0;
}}
.hero-pill {{
    display: inline-block;
    background: {ACCENT};
    color: #fff;
    font-size: 0.72rem;
    font-weight: 700;
    letter-spacing: 0.08em;
    padding: 0.2rem 0.8rem;
    border-radius: 50px;
    margin-bottom: 0.45rem;
    text-transform: uppercase;
}}
.hero-title {{
    font-size: 1.75rem;
    font-weight: 800;
    color: {TEXT_MAIN};
    margin: 0 0 0.2rem 0;
    line-height: 1.2;
}}
.hero-sub {{
    font-size: 0.88rem;
    color: {TEXT_SUB};
    margin: 0;
}}

/* Metric card */
.neu-card {{
    background: {BG};
    border-radius: 20px;
    padding: 1.4rem 1.6rem;
    box-shadow: 8px 8px 16px {SHADOW_DARK}, -8px -8px 16px {SHADOW_LITE};
    height: 100%;
    transition: box-shadow 0.2s ease;
}}
.neu-card:hover {{
    box-shadow: 5px 5px 10px {SHADOW_DARK}, -5px -5px 10px {SHADOW_LITE};
}}
.card-icon-wrap {{
    width: 48px; height: 48px;
    border-radius: 14px;
    display: flex; align-items: center; justify-content: center;
    font-size: 1.4rem;
    margin-bottom: 1rem;
    background: {BG};
    box-shadow: inset 3px 3px 6px {SHADOW_DARK}, inset -3px -3px 6px {SHADOW_LITE};
}}
.card-value {{
    font-size: 2rem;
    font-weight: 800;
    line-height: 1;
    margin-bottom: 0.3rem;
}}
.card-label {{
    font-size: 0.8rem;
    color: {TEXT_SUB};
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.06em;
}}
.card-delta {{
    font-size: 0.75rem;
    margin-top: 0.5rem;
    font-weight: 600;
}}
.c-coral  {{ color: {ACCENT}; }}
.c-blue   {{ color: #6b9bf4; }}
.c-green  {{ color: #6bcfa0; }}
.c-purple {{ color: #b06bf4; }}
.c-amber  {{ color: #f4c46b; }}

/* Section label */
.section-label {{
    font-size: 0.72rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: {TEXT_SUB};
    margin-bottom: 0.6rem;
}}

/* Neumorph panel */
.neu-panel {{
    background: {BG};
    border-radius: 20px;
    padding: 1.6rem;
    box-shadow: 8px 8px 16px {SHADOW_DARK}, -8px -8px 16px {SHADOW_LITE};
    margin-bottom: 1.2rem;
}}

/* Sidebar title */
.sidebar-title {{
    font-size: 1rem;
    font-weight: 800;
    color: {TEXT_MAIN};
    margin-bottom: 0.2rem;
}}
.sidebar-sub {{
    font-size: 0.75rem;
    color: {TEXT_SUB};
    margin-bottom: 1.4rem;
}}

/* Status badge */
.badge-success {{
    background: #e6f9f0; color: #2e9e6b;
    padding: 0.15rem 0.6rem; border-radius: 50px;
    font-size: 0.72rem; font-weight: 700;
}}
.badge-failed {{
    background: #fde8e8; color: #e05050;
    padding: 0.15rem 0.6rem; border-radius: 50px;
    font-size: 0.72rem; font-weight: 700;
}}

/* Divider */
.neu-divider {{
    height: 1px;
    background: linear-gradient(to right, transparent, {SHADOW_DARK}, transparent);
    margin: 1.5rem 0;
    border: none;
}}

/* Inset pill (tag) */
.neu-tag {{
    display: inline-block;
    background: {BG};
    border-radius: 50px;
    padding: 0.2rem 0.8rem;
    font-size: 0.75rem;
    font-weight: 600;
    box-shadow: inset 2px 2px 5px {SHADOW_DARK}, inset -2px -2px 5px {SHADOW_LITE};
    color: {TEXT_SUB};
    margin-right: 0.3rem;
}}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════════════════════

FILE_TYPE_COLORS = {
    "pdf":    ACCENT,
    "word":   "#6b9bf4",
    "excel":  "#6bcfa0",
    "image":  "#b06bf4",
    "others": "#a0a8b8",
}

@st.cache_data(ttl=30)
def load_data(path: str):
    if not os.path.exists(path):
        return None
    df = pd.read_excel(path, engine="openpyxl")
    if df.empty:
        return df
    df.columns = [c.strip().lower().replace(" ", "_") for c in df.columns]
    for col in REPORT_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce").dt.date
    for col in ("status", "file_type"):
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.lower()
    return df


def metric_card(icon, value, label, color_class, delta_text=""):
    delta_html = f'<div class="card-delta {color_class}">{delta_text}</div>' if delta_text else ""
    return f"""
    <div class="neu-card">
        <div class="card-icon-wrap">{icon}</div>
        <div class="card-value {color_class}">{value}</div>
        <div class="card-label">{label}</div>
        {delta_html}
    </div>
    """

# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown('<div class="sidebar-title">📬 RPA Dashboard</div>', unsafe_allow_html=True)
    st.markdown('<div class="sidebar-sub">Email Attachment Organizer</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-label">🗓 Date Range</div>', unsafe_allow_html=True)
    today    = date.today()
    week_ago = today - timedelta(days=7)
    date_from = st.date_input("From", value=week_ago, max_value=today, key="date_from")
    date_to   = st.date_input("To",   value=today,    max_value=today, key="date_to")
    if date_from > date_to:
        st.warning("⚠️ 'From' must be before 'To'.")

    st.markdown('<hr class="neu-divider" style="margin:1rem 0"/>', unsafe_allow_html=True)
    st.markdown('<div class="section-label">🔄 Actions</div>', unsafe_allow_html=True)
    if st.button("🔄  Refresh Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.markdown('<hr class="neu-divider" style="margin:1rem 0"/>', unsafe_allow_html=True)
    st.markdown('<div class="section-label">📂 File Types</div>', unsafe_allow_html=True)
    for ft, color in FILE_TYPE_COLORS.items():
        st.markdown(
            f'<span class="neu-tag" style="color:{color};border-color:{color}">'
            f'{"📄" if ft=="pdf" else "📝" if ft=="word" else "📊" if ft=="excel" else "🖼" if ft=="image" else "📦"} {ft.upper()}'
            f'</span>',
            unsafe_allow_html=True
        )

    st.markdown('<hr class="neu-divider" style="margin:1rem 0"/>', unsafe_allow_html=True)
    st.caption(f"Last loaded: {datetime.now().strftime('%d %b %Y · %H:%M')}")

# ══════════════════════════════════════════════════════════════════════════════
# LOAD DATA
# ══════════════════════════════════════════════════════════════════════════════
df_full = load_data(str(REPORT_FILE))

if df_full is None:
    st.markdown("""
    <div class="neu-panel" style="text-align:center; padding:3rem;">
        <div style="font-size:3rem; margin-bottom:1rem;">📭</div>
        <div style="font-size:1.2rem; font-weight:700; color:#3a4354;">No report file found</div>
        <div style="font-size:0.88rem; color:#7b8498; margin-top:0.5rem;">
            Run <code>python3 bot.py</code> to generate the report.
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Apply date filter
if "date" in df_full.columns:
    mask = (df_full["date"] >= date_from) & (df_full["date"] <= date_to)
    df   = df_full[mask].copy()
else:
    df = df_full.copy()

# ══════════════════════════════════════════════════════════════════════════════
# HERO BANNER
# ══════════════════════════════════════════════════════════════════════════════
last_updated = datetime.now().strftime("%d %b %Y · %H:%M")
st.markdown(f"""
<div class="neu-hero">
    <div class="hero-icon">📬</div>
    <div>
        <div class="hero-pill">Live Dashboard</div>
        <div class="hero-title">RPA Email Attachment</div>
        <p class="hero-sub">Automated organiser &nbsp;·&nbsp; {last_updated}</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# METRICS
# ══════════════════════════════════════════════════════════════════════════════
total   = len(df)
success = len(df[df["status"] == "success"]) if "status" in df.columns else 0
failed  = len(df[df["status"] != "success"]) if "status" in df.columns else 0
pdf_cnt = len(df[df["file_type"] == "pdf"])  if "file_type" in df.columns else 0
img_cnt = len(df[df["file_type"] == "image"]) if "file_type" in df.columns else 0

success_rate = f"{int(success/total*100)}% success rate" if total else "No data"

st.markdown('<div class="section-label">📊 Overview</div>', unsafe_allow_html=True)
c1, c2, c3, c4, c5 = st.columns(5)
cols_data = [
    (c1, "📁", total,   "Total Files",  "c-blue",   success_rate),
    (c2, "✅", success, "Success",      "c-green",  f"+{success} processed"),
    (c3, "❌", failed,  "Failed",       "c-coral",  "Need attention" if failed else "All good ✨"),
    (c4, "📄", pdf_cnt, "PDF Files",    "c-amber",  ""),
    (c5, "🖼", img_cnt, "Image Files",  "c-purple", ""),
]
for col, icon, val, label, cls, delta in cols_data:
    with col:
        st.markdown(metric_card(icon, val, label, cls, delta), unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CHARTS
# ══════════════════════════════════════════════════════════════════════════════
chart_col, donut_col = st.columns([3, 2], gap="large")

with chart_col:
    st.markdown('<div class="section-label">📊 Files by Type</div>', unsafe_allow_html=True)
    st.markdown('<div class="neu-panel">', unsafe_allow_html=True)
    if "file_type" in df.columns and not df.empty:
        ft_counts = (
            df["file_type"]
            .value_counts()
            .reset_index()
            .rename(columns={"file_type": "type", "count": "count"})
        )
        ft_counts["color"] = ft_counts["type"].map(FILE_TYPE_COLORS).fillna("#a0a8b8")

        bar = (
            alt.Chart(ft_counts)
            .mark_bar(cornerRadiusTopLeft=8, cornerRadiusTopRight=8)
            .encode(
                x=alt.X("type:N", sort="-y", title="File Type",
                         axis=alt.Axis(labelColor=TEXT_SUB, titleColor=TEXT_SUB,
                                       labelFontSize=11, grid=False)),
                y=alt.Y("count:Q", title="Count",
                         axis=alt.Axis(labelColor=TEXT_SUB, titleColor=TEXT_SUB,
                                       labelFontSize=11, gridColor="#d0d6df")),
                color=alt.Color("type:N",
                    scale=alt.Scale(domain=list(FILE_TYPE_COLORS.keys()),
                                    range=list(FILE_TYPE_COLORS.values())),
                    legend=None),
                tooltip=[
                    alt.Tooltip("type:N",  title="Type"),
                    alt.Tooltip("count:Q", title="Count"),
                ],
            )
            .properties(height=240, background=BG)
            .configure_view(strokeWidth=0)
            .configure_axis(domainColor=SHADOW_DARK)
        )
        st.altair_chart(bar, use_container_width=True)
    else:
        st.info("No data in selected date range.")
    st.markdown('</div>', unsafe_allow_html=True)

with donut_col:
    st.markdown('<div class="section-label">🍩 Status Split</div>', unsafe_allow_html=True)
    st.markdown('<div class="neu-panel">', unsafe_allow_html=True)
    if "status" in df.columns and not df.empty:
        status_counts = (
            df["status"]
            .value_counts()
            .reset_index()
            .rename(columns={"status": "status", "count": "count"})
        )
        status_colors = {"success": "#6bcfa0", "failed": ACCENT, "error": "#f4c46b"}

        donut = (
            alt.Chart(status_counts)
            .mark_arc(innerRadius=60, outerRadius=100, cornerRadius=6)
            .encode(
                theta=alt.Theta("count:Q"),
                color=alt.Color("status:N",
                    scale=alt.Scale(
                        domain=list(status_colors.keys()),
                        range=list(status_colors.values())),
                    legend=alt.Legend(
                        orient="bottom",
                        labelColor=TEXT_SUB,
                        labelFontSize=11,
                        symbolSize=80,
                        titleColor=TEXT_SUB,
                    )),
                tooltip=[
                    alt.Tooltip("status:N", title="Status"),
                    alt.Tooltip("count:Q",  title="Count"),
                ],
            )
            .properties(height=240, background=BG)
            .configure_view(strokeWidth=0)
        )
        st.altair_chart(donut, use_container_width=True)
    else:
        st.info("No status data.")
    st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# TREND LINE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">📈 Daily Activity</div>', unsafe_allow_html=True)
st.markdown('<div class="neu-panel">', unsafe_allow_html=True)
if "date" in df.columns and not df.empty and df["date"].notna().any():
    daily = (
        df.groupby(["date", "status"])
        .size()
        .reset_index(name="count")
    )
    daily["date"] = pd.to_datetime(daily["date"])

    trend = (
        alt.Chart(daily)
        .mark_area(
            interpolate="monotone",
            fillOpacity=0.18,
            strokeWidth=2.5,
            point=alt.OverlayMarkDef(filled=True, size=60),
        )
        .encode(
            x=alt.X("date:T", title="Date",
                     axis=alt.Axis(format="%d %b", labelColor=TEXT_SUB,
                                   titleColor=TEXT_SUB, grid=False)),
            y=alt.Y("count:Q", title="Files",
                     axis=alt.Axis(labelColor=TEXT_SUB, titleColor=TEXT_SUB,
                                   gridColor="#d0d6df")),
            color=alt.Color("status:N",
                scale=alt.Scale(domain=["success", "failed", "error"],
                                range=["#6bcfa0", ACCENT, "#f4c46b"]),
                legend=alt.Legend(orient="top-right", labelColor=TEXT_SUB)),
            tooltip=[
                alt.Tooltip("date:T",   title="Date",   format="%Y-%m-%d"),
                alt.Tooltip("status:N", title="Status"),
                alt.Tooltip("count:Q",  title="Files"),
            ],
        )
        .properties(height=200, background=BG)
        .configure_view(strokeWidth=0)
        .configure_axis(domainColor=SHADOW_DARK)
    )
    st.altair_chart(trend, use_container_width=True)
else:
    st.info("Not enough data to show trend.")
st.markdown('</div>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# LOG TABLE
# ══════════════════════════════════════════════════════════════════════════════
st.markdown('<div class="section-label">📋 Activity Log</div>', unsafe_allow_html=True)

# Search bar
search = st.text_input("🔍  Search files, subjects or senders…", placeholder="Type to filter…", key="search")

DISPLAY_COLS = {
    "date":          "Date",
    "time":          "Time",
    "email_subject": "Subject",
    "sender":        "Sender",
    "file_name":     "File Name",
    "file_type":     "Type",
    "status":        "Status",
    "error_message": "Error",
}
available = {k: v for k, v in DISPLAY_COLS.items() if k in df.columns}
display_df = df[list(available.keys())].copy()
display_df.columns = list(available.values())

if search:
    mask = display_df.apply(
        lambda col: col.astype(str).str.contains(search, case=False, na=False)
    ).any(axis=1)
    display_df = display_df[mask]

st.markdown('<div class="neu-panel" style="padding:0.8rem 1rem;">', unsafe_allow_html=True)
st.dataframe(
    display_df,
    width="stretch",
    hide_index=True,
    height=380,
    column_config={
        "Date":    st.column_config.DateColumn("Date",   format="YYYY-MM-DD"),
        "Status":  st.column_config.TextColumn("Status"),
        "Type":    st.column_config.TextColumn("Type"),
    },
)
st.markdown('</div>', unsafe_allow_html=True)

# Summary footer
st.markdown(f"""
<div style="text-align:center; margin-top:2rem;">
    <span class="neu-tag">📁 {total} total</span>
    <span class="neu-tag" style="color:#6bcfa0;">✅ {success} success</span>
    <span class="neu-tag" style="color:{ACCENT};">❌ {failed} failed</span>
    <span class="neu-tag">📅 {date_from} → {date_to}</span>
</div>
""", unsafe_allow_html=True)
