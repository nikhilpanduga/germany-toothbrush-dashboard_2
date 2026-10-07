# app.py
# Germany Toothbrush Customer Review Dashboard
# Run with:  streamlit run app.py
# Requires:  pip install streamlit pandas plotly numpy openpyxl
from __future__ import annotations

import html
import io
import os
import re
import textwrap
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# 1. CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="Toothbrush Review Dashboard",
    page_icon="🪥",
    layout="wide",
    initial_sidebar_state="expanded",
)

DATA_FILE = "toothbrush data.xlsx"
SHEET_NAME = "Germany"

PRODUCTS = [
    "curaprox 5460",
    "Dr. Best Clean Pro Zwischenzahn",
    "elmex expert precision",
    "elmex InterX",
    "meridol base",
]

# Keywords used to map raw product names onto the five canonical products
PRODUCT_KEYWORDS = {
    "curaprox 5460": ["5460", "curaprox"],
    "Dr. Best Clean Pro Zwischenzahn": ["dr. best", "dr best", "zwischenzahn", "clean pro"],
    "elmex InterX": ["interx", "inter x", "inter-x"],
    "elmex expert precision": ["expert", "precision"],
    "meridol base": ["meridol"],
}

THEMES = {
    "Cleaning Performance": "Cleaning Performance_Present",
    "Brushing Comfort": "Brushing Comfort_Present",
    "Design & Ease of Use": "Design & Ease of Use_Present",
    "Quality & Durability": "Quality & Durability_Present",
    "Features & Technology": "Features & Technology_Present",
    "Price & Value": "Price & Value_Present",
}

SENTIMENTS = ["positive", "neutral", "negative"]
MIN_PRODUCT_REVIEWS = 10   # minimum reviews for a product to be ranked in "Key Takeaways"
MIN_THEME_REVIEWS = 10     # minimum theme mentions for a theme to be ranked as strength / pain point
CHART_HEIGHT = 400

# ---- Colour palette -------------------------------------------------------
NAVY = "#14325C"
BLUE = "#2F6DB5"
TEXT = "#2B3445"
GREY = "#8A94A6"
GRID = "#E9EEF5"
POS_COLOR = "#2E9E8A"
NEU_COLOR = "#C3CAD6"
NEG_COLOR = "#E0666B"
SENT_COLORS = {"positive": POS_COLOR, "neutral": NEU_COLOR, "negative": NEG_COLOR}
PRODUCT_COLORS = {
    "curaprox 5460": "#1B3A6B",
    "Dr. Best Clean Pro Zwischenzahn": "#3E86CF",
    "elmex expert precision": "#8EC5F0",
    "elmex InterX": "#F0A33C",
    "meridol base": "#6B7A90",
}
BLUE_SCALE = [[0, "#EEF4FB"], [1, "#6FA8E0"]]
GREEN_SCALE = [[0, "#EAF6F3"], [1, "#6CC4B0"]]
RED_SCALE = [[0, "#FBEEEE"], [1, "#EE9A9E"]]


def _st_version() -> tuple:
    m = re.match(r"(\d+)\.(\d+)", st.__version__)
    return (int(m.group(1)), int(m.group(2))) if m else (0, 0)


# Streamlit changed how "full width" is requested; support both old and new versions
STRETCH = {"width": "stretch"} if _st_version() >= (1, 50) else {"use_container_width": True}

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background-color: #F5F7FB; }
header[data-testid="stHeader"] { background: transparent; }
.block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1400px; }
footer, #MainMenu { visibility: hidden; }

/* Sidebar */
section[data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E6ECF4; }
section[data-testid="stSidebar"] h3 { color: #14325C; font-weight: 700; margin-bottom: 0.2rem; }
section[data-testid="stSidebar"] .stButton button {
    border-radius: 10px; border: 1px solid #2F6DB5; color: #2F6DB5; background: #FFFFFF; font-weight: 600;
}
section[data-testid="stSidebar"] .stButton button:hover { background: #2F6DB5; color: #FFFFFF; border-color: #2F6DB5; }
div[data-baseweb="tag"] { background-color: #2F6DB5 !important; border-radius: 6px !important; }
div[data-baseweb="tag"] span { color: #FFFFFF !important; }
div[data-testid="stSlider"] div[role="slider"] { background-color: #2F6DB5 !important; }
div[data-testid="stSliderThumbValue"], div[data-testid="stSliderTickBarMin"], div[data-testid="stSliderTickBarMax"] { color: #2F6DB5 !important; }
.filter-count {
    background: #EEF4FB; border-radius: 10px; padding: 10px 14px; margin: 14px 0 10px 0;
    color: #14325C; font-size: 0.85rem;
}

/* Header */
.dash-header { padding: 4px 0 22px 0; }
.dash-title { font-size: 2.05rem; font-weight: 700; color: #14325C; margin: 0; letter-spacing: -0.5px; line-height: 1.2; }
.dash-subtitle { font-size: 1rem; color: #6B7689; margin: 6px 0 0 0; }

/* KPI cards */
.kpi-card {
    background: #FFFFFF; border-radius: 16px; padding: 18px 20px 16px 20px;
    box-shadow: 0 1px 3px rgba(20, 50, 92, 0.07); border: 1px solid #EAEFF6;
    border-top: 4px solid var(--accent); min-height: 128px;
}
.kpi-label { font-size: 0.74rem; font-weight: 600; color: #7A8498; text-transform: uppercase; letter-spacing: 0.7px; }
.kpi-value { font-size: 2.1rem; font-weight: 700; color: #14325C; line-height: 1.15; margin-top: 8px; }
.kpi-sub { font-size: 0.78rem; color: #98A2B3; margin-top: 4px; }

/* Insight cards */
.insight-title { font-size: 0.8rem; font-weight: 600; color: #7A8498; text-transform: uppercase; letter-spacing: 0.7px; margin: 26px 0 8px 0; }
.insight-card {
    background: #FFFFFF; border-radius: 14px; padding: 14px 18px; border: 1px solid #EAEFF6;
    border-left: 5px solid var(--accent); box-shadow: 0 1px 3px rgba(20, 50, 92, 0.05); min-height: 104px;
}
.insight-label { font-size: 0.74rem; font-weight: 600; color: #7A8498; text-transform: uppercase; letter-spacing: 0.5px; }
.insight-head { font-size: 1.05rem; font-weight: 700; color: #14325C; margin-top: 5px; line-height: 1.3; }
.insight-sub { font-size: 0.82rem; color: #6B7689; margin-top: 3px; }

/* Section & chart headings */
.section-title {
    font-size: 1.3rem; font-weight: 700; color: #14325C; margin: 44px 0 2px 0;
    padding-left: 12px; border-left: 5px solid #2F6DB5; line-height: 1.3;
}
.section-caption { font-size: 0.88rem; color: #7A8498; margin: 0 0 6px 17px; }
.chart-title { font-size: 1rem; font-weight: 600; color: #2B3445; margin: 18px 0 0 0; }
.chart-caption { font-size: 0.82rem; color: #8A94A6; margin: 2px 0 8px 0; }

/* Chart & table cards */
div[data-testid="stPlotlyChart"] {
    background: #FFFFFF; border-radius: 14px; padding: 8px; border: 1px solid #EAEFF6;
    box-shadow: 0 1px 3px rgba(20, 50, 92, 0.05);
}
div[data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; border: 1px solid #EAEFF6; }
</style>
"""


# ============================================================
# 2. DATA LOADING
# ============================================================
@st.cache_data(show_spinner="Loading data...")
def load_raw(source, mtime: float = 0.0) -> pd.DataFrame:
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    return pd.read_excel(source, sheet_name=SHEET_NAME)


def find_data_source():
    """Look for the Excel file next to app.py, then in the working directory; else offer an upload."""
    for folder in (Path(__file__).parent, Path.cwd()):
        path = folder / DATA_FILE
        if path.exists():
            return str(path), os.path.getmtime(path), None
    return None, 0.0, st.sidebar.file_uploader("Upload the Excel data file", type=["xlsx"])


# ============================================================
# 3. DATA CLEANING
# ============================================================
def clean_text_key(value) -> str:
    return re.sub(r"\s+", " ", str(value)).strip()


def map_product(row_text: str):
    text = row_text.lower()
    for product in PRODUCTS:  # exact match first
        if product.lower() in text:
            return product
    for product, keywords in PRODUCT_KEYWORDS.items():
        if any(k in text for k in keywords):
            return product
    return None


def standardize_sentiment(value):
    if pd.isna(value):
        return np.nan
    v = str(value).strip().lower()
    if v.startswith("pos"):
        return "positive"
    if v.startswith("neu"):
        return "neutral"
    if v.startswith("neg"):
        return "negative"
    return np.nan


def standardize_theme(value):
    if pd.isna(value):
        return np.nan
    v = str(value).strip().lower()
    if v in {"yes", "y", "true", "1", "1.0", "present"}:
        return "Yes"
    if v in {"no", "n", "false", "0", "0.0", "absent"}:
        return "No"
    return np.nan


@st.cache_data(show_spinner="Preparing data...")
def clean_data(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw.copy()

    # Clean column names (trim, collapse spaces, align case with expected names)
    df.columns = [clean_text_key(c) for c in df.columns]
    expected = (
        ["Country", "Brand", "Product Name", "Product Name given", "SKU", "URL", "Site",
         "Title", "Text", "Posted On", "Stars", "language", "translated text", "Sentiment"]
        + list(THEMES.values())
    )
    lookup = {c.lower(): c for c in expected}
    df.columns = [lookup.get(c.lower(), c) for c in df.columns]
    for col in expected:
        if col not in df.columns:
            df[col] = np.nan

    # Types
    df["Posted On"] = pd.to_datetime(df["Posted On"], errors="coerce")
    df["Stars"] = pd.to_numeric(df["Stars"], errors="coerce")
    df["Sentiment"] = df["Sentiment"].apply(standardize_sentiment)
    for col in THEMES.values():
        df[col] = df[col].apply(standardize_theme)
    for col in ["Country", "Site"]:
        df[col] = df[col].fillna("Unknown").astype(str).str.strip()

    # Map to the five required products and drop anything else
    combined = df["Product Name"].fillna("").astype(str) + " | " + df["Product Name given"].fillna("").astype(str)
    df["Product"] = combined.apply(map_product)
    df = df[df["Product"].notna()].copy()

    # Remove duplicate reviews
    dedupe_cols = [c for c in ["Product", "Site", "Title", "Text", "Posted On", "Stars"] if c in df.columns]
    return df.drop_duplicates(subset=dedupe_cols).reset_index(drop=True)


# ============================================================
# 4. FILTER FUNCTIONS
# ============================================================
def default_filters(df: pd.DataFrame) -> dict:
    dates = df["Posted On"].dropna()
    date_range = (dates.min().date(), dates.max().date()) if not dates.empty else ()
    return {
        "f_product": ["All Products"],
        "f_country": [],
        "f_site": [],
        "f_sentiment": ["Positive", "Neutral", "Negative"],
        "f_stars": (1, 5),
        "f_dates": date_range,
    }


def reset_filters(defaults: dict):
    for key, value in defaults.items():
        st.session_state[key] = value


def render_sidebar(df: pd.DataFrame) -> dict:
    defaults = default_filters(df)
    for key, value in defaults.items():
        st.session_state.setdefault(key, value)

    st.sidebar.markdown("### Filters")
    st.sidebar.multiselect(
        "Product", ["All Products"] + [p for p in PRODUCTS if p in set(df["Product"])], key="f_product"
    )
    st.sidebar.multiselect("Country", sorted(df["Country"].unique()), key="f_country", placeholder="All countries")
    st.sidebar.multiselect("Site", sorted(df["Site"].unique()), key="f_site", placeholder="All sites")
    st.sidebar.multiselect("Sentiment", ["Positive", "Neutral", "Negative"], key="f_sentiment")
    st.sidebar.slider("Star Rating", 1, 5, key="f_stars")
    if defaults["f_dates"]:
        st.sidebar.date_input(
            "Review Date", key="f_dates", min_value=defaults["f_dates"][0], max_value=defaults["f_dates"][1]
        )
    return defaults


def render_sidebar_footer(defaults: dict, shown: int, total: int):
    st.sidebar.markdown(
        f'<div class="filter-count"><b>{shown:,}</b> of {total:,} reviews selected</div>', unsafe_allow_html=True
    )
    st.sidebar.button("Reset Filters", on_click=reset_filters, args=(defaults,), **STRETCH)


def apply_filters(df: pd.DataFrame, defaults: dict) -> pd.DataFrame:
    ss = st.session_state
    mask = pd.Series(True, index=df.index)

    products = ss["f_product"]
    if products and "All Products" not in products:
        mask &= df["Product"].isin(products)
    if ss["f_country"]:
        mask &= df["Country"].isin(ss["f_country"])
    if ss["f_site"]:
        mask &= df["Site"].isin(ss["f_site"])

    sentiments = [s.lower() for s in ss["f_sentiment"]]
    if sentiments and len(sentiments) < 3:
        mask &= df["Sentiment"].isin(sentiments)

    lo, hi = ss["f_stars"]
    if (lo, hi) != (1, 5):
        mask &= df["Stars"].between(lo, hi)

    dates = ss.get("f_dates", ())
    if defaults["f_dates"] and len(dates) == 2 and tuple(dates) != tuple(defaults["f_dates"]):
        start, end = pd.Timestamp(dates[0]), pd.Timestamp(dates[1]) + pd.Timedelta(days=1)
        mask &= (df["Posted On"] >= start) & (df["Posted On"] < end)

    return df[mask]


# ============================================================
# 5. KPI & AGGREGATION FUNCTIONS
# ============================================================
def sentiment_share(df: pd.DataFrame, label: str) -> float:
    valid = df["Sentiment"].notna().sum()
    return (df["Sentiment"] == label).sum() / valid * 100 if valid else 0.0


def ordered_products(df: pd.DataFrame) -> list:
    return [p for p in PRODUCTS if p in set(df["Product"])]


def compute_kpis(df: pd.DataFrame) -> dict:
    return {
        "total": len(df),
        "avg_rating": df["Stars"].mean() if df["Stars"].notna().any() else np.nan,
        "positive": sentiment_share(df, "positive"),
        "negative": sentiment_share(df, "negative"),
        "neutral": sentiment_share(df, "neutral"),
        "n_positive": int((df["Sentiment"] == "positive").sum()),
        "n_negative": int((df["Sentiment"] == "negative").sum()),
        "n_neutral": int((df["Sentiment"] == "neutral").sum()),
    }


def product_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for product in ordered_products(df):
        sub = df[df["Product"] == product]
        rows.append({
            "Product": product,
            "Reviews": len(sub),
            "Average Rating": sub["Stars"].mean(),
            "Positive %": sentiment_share(sub, "positive"),
            "Neutral %": sentiment_share(sub, "neutral"),
            "Negative %": sentiment_share(sub, "negative"),
        })
    return pd.DataFrame(rows).sort_values("Average Rating", ascending=False).reset_index(drop=True)


def theme_stats(df: pd.DataFrame) -> pd.DataFrame:
    """Per theme: reviews mentioning it, and share of those that are positive / neutral / negative."""
    rows = []
    for theme, col in THEMES.items():
        mentioned = df[df[col] == "Yes"]
        n = len(mentioned)
        share = lambda label: (mentioned["Sentiment"] == label).sum() / n * 100 if n else np.nan
        rows.append({
            "Theme": theme, "Reviews": n,
            "Positive %": share("positive"), "Neutral %": share("neutral"), "Negative %": share("negative"),
        })
    return pd.DataFrame(rows)


def rank_themes(stats: pd.DataFrame, by: str) -> pd.DataFrame:
    eligible = stats[stats["Reviews"] >= MIN_THEME_REVIEWS]
    if eligible.empty:
        eligible = stats[stats["Reviews"] > 0]
    return eligible.sort_values([by, "Reviews"], ascending=False)


def product_theme_matrices(df: pd.DataFrame):
    """Product x Theme matrices: penetration %, positive %, negative %, and mention counts."""
    products = ordered_products(df)
    pen = pd.DataFrame(np.nan, index=products, columns=list(THEMES))
    pos, neg, cnt = pen.copy(), pen.copy(), pen.copy()
    for product in products:
        sub = df[df["Product"] == product]
        for theme, col in THEMES.items():
            mentioned = sub[sub[col] == "Yes"]
            n = len(mentioned)
            cnt.loc[product, theme] = n
            pen.loc[product, theme] = n / len(sub) * 100 if len(sub) else np.nan
            if n:
                pos.loc[product, theme] = (mentioned["Sentiment"] == "positive").mean() * 100
                neg.loc[product, theme] = (mentioned["Sentiment"] == "negative").mean() * 100
    return pen, pos, neg, cnt


def compute_takeaways(df: pd.DataFrame) -> list:
    """Four headline insights that respond to the active filters."""
    summary = product_summary(df)
    eligible = summary[summary["Reviews"] >= MIN_PRODUCT_REVIEWS]
    if eligible.empty:
        eligible = summary

    cards = []
    ranked = eligible.dropna(subset=["Average Rating"])
    if not ranked.empty:
        best = ranked.sort_values("Average Rating", ascending=False).iloc[0]
        cards.append(("Highest Rated Product", best["Product"],
                      f'{best["Average Rating"]:.2f} ★ from {int(best["Reviews"]):,} reviews', BLUE))
    best_pos = eligible.sort_values("Positive %", ascending=False).iloc[0]
    cards.append(("Most Positive Perception", best_pos["Product"],
                  f'{best_pos["Positive %"]:.0f}% positive reviews', POS_COLOR))

    stats = theme_stats(df)
    strengths = rank_themes(stats, "Positive %").dropna(subset=["Positive %"])
    pains = rank_themes(stats, "Negative %").dropna(subset=["Negative %"])
    if not strengths.empty:
        s = strengths.iloc[0]
        cards.append(("Biggest Strength", s["Theme"],
                      f'{s["Positive %"]:.0f}% positive across {int(s["Reviews"]):,} mentions', NAVY))
    if not pains.empty:
        p = pains.iloc[0]
        cards.append(("Biggest Pain Point", p["Theme"],
                      f'{p["Negative %"]:.0f}% negative across {int(p["Reviews"]):,} mentions', NEG_COLOR))
    return cards


# ============================================================
# 6. CHART FUNCTIONS
# ============================================================
def wrap(label, width: int = 22) -> str:
    text = str(label)
    return "<br>".join(textwrap.wrap(text, width=width, break_long_words=False)) or text


def style_fig(fig: go.Figure, legend_top=None, height: int = CHART_HEIGHT) -> go.Figure:
    """Common look. Legends sit above the plot inside a reserved top margin, so they never overlap data."""
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=30, t=legend_top or 20, b=20),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(family="Inter, sans-serif", color=TEXT, size=12),
        showlegend=legend_top is not None,
        legend=dict(orientation="h", x=0, y=1, xanchor="left", yanchor="bottom",
                    font=dict(size=11), title=None, bgcolor="rgba(0,0,0,0)"),
        hoverlabel=dict(bgcolor="white", font_size=12),
    )
    fig.update_xaxes(gridcolor=GRID, zeroline=False, showline=False, automargin=True, title_standoff=10)
    fig.update_yaxes(gridcolor=GRID, zeroline=False, showline=False, automargin=True, title_standoff=10)
    return fig


def hbar(labels, values, colors, xtitle, text_fmt, x_range, hover_fmt, inside=False, tickvals=None) -> go.Figure:
    fig = go.Figure(go.Bar(
        x=values, y=[wrap(l) for l in labels], orientation="h", marker_color=colors, customdata=labels,
        text=values, texttemplate=text_fmt, textposition="inside" if inside else "outside",
        insidetextanchor="end", cliponaxis=False,
        textfont=dict(color="#FFFFFF" if inside else TEXT, size=13),
        hovertemplate="%{customdata}<br>" + xtitle + ": %{x:" + hover_fmt + "}<extra></extra>",
    ))
    fig.update_xaxes(title=xtitle, range=x_range, tickvals=tickvals)
    fig.update_yaxes(title=None)
    return style_fig(fig)


def chart_reviews_by_product(df: pd.DataFrame) -> go.Figure:
    agg = df.groupby("Product").size().sort_values()
    return hbar(list(agg.index), agg.values, [PRODUCT_COLORS.get(p, BLUE) for p in agg.index],
                "Number of Reviews", "%{text:,}", [0, agg.max() * 1.2], ",.0f")


def chart_rating_by_product(df: pd.DataFrame) -> go.Figure:
    agg = df.groupby("Product")["Stars"].mean().dropna().sort_values()
    return hbar(list(agg.index), agg.values, [PRODUCT_COLORS.get(p, BLUE) for p in agg.index],
                "Average Rating (out of 5)", "%{text:.2f}", [0, 5], ".2f", inside=True, tickvals=[0, 1, 2, 3, 4, 5])


def chart_sentiment_share(df: pd.DataFrame, group_col: str, groups: list) -> go.Figure:
    """100% stacked horizontal bars of sentiment for each group (product or site)."""
    groups = groups[::-1]  # first group appears on top
    counts = {g: int((df[group_col] == g).sum()) for g in groups}
    labels = [wrap(f"{g} (n={counts[g]:,})", 24) for g in groups]
    fig = go.Figure()
    for sentiment in SENTIMENTS:
        values = [sentiment_share(df[df[group_col] == g], sentiment) for g in groups]
        fig.add_trace(go.Bar(
            name=sentiment.title(), x=values, y=labels, orientation="h", marker_color=SENT_COLORS[sentiment],
            text=[f"{v:.0f}%" if v >= 6 else "" for v in values], textposition="inside", insidetextanchor="middle",
            textfont=dict(color=TEXT if sentiment == "neutral" else "#FFFFFF", size=12),
            hovertemplate="%{y}<br>" + sentiment.title() + ": %{x:.1f}%<extra></extra>",
        ))
    fig.update_layout(barmode="stack", bargap=0.35)
    fig.update_xaxes(title="Share of Reviews (%)", range=[0, 100], ticksuffix="%")
    fig.update_yaxes(title=None)
    return style_fig(fig, legend_top=50)


def chart_rating_distribution(df: pd.DataFrame) -> go.Figure:
    stars = df.dropna(subset=["Stars"]).copy()
    stars["Star"] = stars["Stars"].round().clip(1, 5).astype(int)
    labels = [f"{i} Star" + ("s" if i > 1 else "") for i in range(1, 6)]
    fig, ymax = go.Figure(), 10
    for product in ordered_products(stars):
        sub = stars[stars["Product"] == product]
        pct = sub["Star"].value_counts(normalize=True).reindex(range(1, 6)).fillna(0) * 100
        ymax = max(ymax, pct.max())
        fig.add_trace(go.Bar(
            name=product, x=labels, y=pct.values, marker_color=PRODUCT_COLORS.get(product, BLUE),
            text=pct.values, texttemplate="%{text:.0f}%", textposition="outside", cliponaxis=False,
            textfont=dict(size=11, color=TEXT),
            hovertemplate=product + "<br>%{x}: %{y:.1f}%<extra></extra>",
        ))
    fig.update_layout(barmode="group", bargap=0.22, bargroupgap=0.06)
    fig.update_yaxes(title="Share of Product's Reviews (%)", range=[0, ymax * 1.18], ticksuffix="%")
    fig.update_xaxes(title=None)
    return style_fig(fig, legend_top=60)


def chart_theme_importance(df: pd.DataFrame) -> go.Figure:
    data = theme_stats(df).sort_values("Reviews", ascending=True)  # highest ends up on top
    return hbar(list(data["Theme"]), data["Reviews"].values, BLUE, "Number of Reviews Mentioning Theme",
                "%{text:,}", [0, max(data["Reviews"].max() * 1.2, 5)], ",.0f")


def heatmap(pct: pd.DataFrame, counts: pd.DataFrame, colorscale, zmax: float = 100.0) -> go.Figure:
    text = [[f"{v:.0f}%" if pd.notna(v) else "–" for v in row] for row in pct.values]
    hover = [
        [f"{pct.index[i]}<br>{pct.columns[j]}: {v:.1f}% (n={int(counts.iloc[i, j])})" if pd.notna(v)
         else f"{pct.index[i]}<br>{pct.columns[j]}: no data" for j, v in enumerate(row)]
        for i, row in enumerate(pct.values)
    ]
    fig = go.Figure(go.Heatmap(
        z=pct.values, x=[wrap(c, 14) for c in pct.columns], y=[wrap(i, 18) for i in pct.index],
        colorscale=colorscale, zmin=0, zmax=zmax, showscale=False, xgap=3, ygap=3,
        text=text, texttemplate="%{text}", textfont=dict(size=13, color=NAVY),
        hovertext=hover, hoverinfo="text",
    ))
    fig.update_yaxes(autorange="reversed", showgrid=False, title=None)
    fig.update_xaxes(showgrid=False, tickangle=0, title=None)
    return style_fig(fig)


# ---- Time charts ----------------------------------------------------------
def add_period(df: pd.DataFrame, view: str):
    """Add a 'Period' column (month or quarter start) and return the complete period range."""
    freq, range_freq = ("M", "MS") if view == "Monthly" else ("Q", "QS")
    dated = df.dropna(subset=["Posted On"]).copy()
    dated["Period"] = dated["Posted On"].dt.to_period(freq).dt.to_timestamp()
    full_range = pd.date_range(dated["Period"].min(), dated["Period"].max(), freq=range_freq)
    return dated, full_range


def apply_time_axis(fig: go.Figure, full_range: pd.DatetimeIndex, view: str, ytitle: str):
    if view == "Quarterly":
        fig.update_xaxes(tickvals=list(full_range),
                         ticktext=[f"Q{(d.month - 1) // 3 + 1} {d.year}" for d in full_range])
    else:
        n = len(full_range)
        dtick = "M1" if n <= 8 else "M2" if n <= 16 else "M3" if n <= 36 else "M6"
        fig.update_xaxes(tickformat="%b %Y", dtick=dtick)
    fig.update_xaxes(title=None)
    fig.update_yaxes(title=ytitle)


def chart_review_volume(df: pd.DataFrame, view: str) -> go.Figure:
    dated, full_range = add_period(df, view)
    fig = go.Figure()
    for product in ordered_products(dated):
        counts = dated[dated["Product"] == product].groupby("Period").size().reindex(full_range, fill_value=0)
        fig.add_trace(go.Scatter(
            name=product, x=counts.index, y=counts.values, mode="lines+markers",
            line=dict(color=PRODUCT_COLORS.get(product, BLUE), width=2.5), marker=dict(size=6),
            hovertemplate=product + "<br>%{x|%b %Y}: %{y:,} reviews<extra></extra>",
        ))
    fig.update_yaxes(rangemode="tozero")
    apply_time_axis(fig, full_range, view, "Number of Reviews")
    return style_fig(fig, legend_top=90)


def chart_rating_trend(df: pd.DataFrame, view: str) -> go.Figure:
    dated, full_range = add_period(df, view)
    fig = go.Figure()
    for product in ordered_products(dated):
        avg = dated[dated["Product"] == product].groupby("Period")["Stars"].mean().reindex(full_range)
        fig.add_trace(go.Scatter(
            name=product, x=avg.index, y=avg.values, mode="lines+markers", connectgaps=True,
            line=dict(color=PRODUCT_COLORS.get(product, BLUE), width=2.5), marker=dict(size=6),
            hovertemplate=product + "<br>%{x|%b %Y}: %{y:.2f} ★<extra></extra>",
        ))
    fig.update_yaxes(range=[1, 5], dtick=1)
    apply_time_axis(fig, full_range, view, "Average Rating (out of 5)")
    return style_fig(fig, legend_top=90)


# ============================================================
# 7. TABLE HELPERS
# ============================================================
def lerp_hex(c1: str, c2: str, t: float) -> str:
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02x%02x%02x" % tuple(round(x + (y - x) * t) for x, y in zip(a, b))


def gradient(series: pd.Series, high: str, vmin: float, vmax: float, low: str = "#FFFFFF") -> list:
    styles = []
    for v in series:
        if pd.isna(v) or vmax <= vmin:
            styles.append("")
        else:
            t = min(max((v - vmin) / (vmax - vmin), 0), 1)
            styles.append(f"background-color: {lerp_hex(low, high, t)}; color: {NAVY}")
    return styles


def show_table(df: pd.DataFrame, formats: dict, gradients: dict):
    """Display a table with number formats and subtle colour gradients (no matplotlib needed)."""
    styler = df.style.format(formats, na_rep="–")
    for col, (high, vmin, vmax) in gradients.items():
        styler = styler.apply(gradient, high=high, vmin=vmin, vmax=vmax, subset=[col])
    st.dataframe(styler, hide_index=True, **STRETCH)


# ============================================================
# 8. DASHBOARD LAYOUT
# ============================================================
def show_chart(fig: go.Figure):
    st.plotly_chart(fig, config={"displayModeBar": False}, **STRETCH)


def section(title: str, caption: str = ""):
    st.markdown(f'<div class="section-title">{html.escape(title)}</div>', unsafe_allow_html=True)
    if caption:
        st.markdown(f'<div class="section-caption">{html.escape(caption)}</div>', unsafe_allow_html=True)


def chart_header(title: str, caption: str = ""):
    st.markdown(f'<div class="chart-title">{html.escape(title)}</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="chart-caption">{html.escape(caption) or "&nbsp;"}</div>', unsafe_allow_html=True)


def render_header():
    st.markdown(
        '<div class="dash-header"><p class="dash-title">Germany Toothbrush Customer Review Dashboard</p>'
        '<p class="dash-subtitle">Customer perception, ratings, sentiment and key product themes</p></div>',
        unsafe_allow_html=True,
    )


def kpi_card(label: str, value: str, sub: str, accent: str) -> str:
    return (f'<div class="kpi-card" style="--accent:{accent}"><div class="kpi-label">{html.escape(label)}</div>'
            f'<div class="kpi-value">{value}</div><div class="kpi-sub">{html.escape(sub)}</div></div>')


def render_kpis(df: pd.DataFrame):
    k = compute_kpis(df)
    avg = f"{k['avg_rating']:.2f} ★" if pd.notna(k["avg_rating"]) else "–"
    cards = [
        ("Total Reviews", f"{k['total']:,}", "in current selection", NAVY),
        ("Average Rating", avg, "out of 5 stars", BLUE),
        ("Positive Sentiment", f"{k['positive']:.1f}%", f"{k['n_positive']:,} reviews", POS_COLOR),
        ("Negative Sentiment", f"{k['negative']:.1f}%", f"{k['n_negative']:,} reviews", NEG_COLOR),
        ("Neutral Sentiment", f"{k['neutral']:.1f}%", f"{k['n_neutral']:,} reviews", "#9AA5B8"),
    ]
    for col, (label, value, sub, accent) in zip(st.columns(5), cards):
        col.markdown(kpi_card(label, value, sub, accent), unsafe_allow_html=True)


def render_takeaways(df: pd.DataFrame):
    st.markdown('<div class="insight-title">Key Takeaways</div>', unsafe_allow_html=True)
    cards = compute_takeaways(df)
    for col, (label, head, sub, accent) in zip(st.columns(4), cards):
        col.markdown(
            f'<div class="insight-card" style="--accent:{accent}"><div class="insight-label">{html.escape(label)}</div>'
            f'<div class="insight-head">{html.escape(str(head))}</div>'
            f'<div class="insight-sub">{html.escape(sub)}</div></div>',
            unsafe_allow_html=True,
        )


def render_product_performance(df: pd.DataFrame):
    section("Product Performance", "How much customers talk about each product, and how highly they rate it")
    c1, c2 = st.columns(2, gap="large")
    with c1:
        chart_header("Reviews by Product", "Number of customer reviews per product")
        show_chart(chart_reviews_by_product(df))
    with c2:
        chart_header("Average Rating by Product", "Mean star rating on a 0–5 scale")
        show_chart(chart_rating_by_product(df))


def render_rating_distribution(df: pd.DataFrame):
    section("Customer Rating Distribution", "How each product's reviews are spread across 1 to 5 stars")
    chart_header("Rating Distribution by Product", "Share of each product's reviews by star rating")
    show_chart(chart_rating_distribution(df))


def render_sentiment(df: pd.DataFrame):
    section("Customer Sentiment", "Overall perception of each product and where reviews are posted")
    top_sites = df["Site"].value_counts().head(8).index.tolist()
    c1, c2 = st.columns(2, gap="large")
    with c1:
        chart_header("Sentiment by Product", "Share of positive, neutral and negative reviews")
        show_chart(chart_sentiment_share(df, "Product", ordered_products(df)))
    with c2:
        chart_header("Sentiment by Site", "Top sites by review volume (n = reviews)")
        show_chart(chart_sentiment_share(df, "Site", top_sites))


def render_themes(df: pd.DataFrame):
    section("What Customers Talk About", "The topics customers mention most, and how each product compares")
    pen, _, _, cnt = product_theme_matrices(df)
    zmax = max(float(np.nanmax(pen.values)), 10.0) if pen.notna().any().any() else 100.0
    c1, c2 = st.columns(2, gap="large")
    with c1:
        chart_header("Customer Theme Importance", "Number of reviews mentioning each theme")
        show_chart(chart_theme_importance(df))
    with c2:
        chart_header("Themes by Product", "% of each product's reviews mentioning the theme")
        show_chart(heatmap(pen, cnt, BLUE_SCALE, zmax=zmax))


def render_theme_sentiment(df: pd.DataFrame):
    section("Customer Sentiment by Theme", "Whether customers speak positively or negatively about each theme")
    _, pos, neg, cnt = product_theme_matrices(df)
    c1, c2 = st.columns(2, gap="large")
    with c1:
        chart_header("Positive Sentiment by Theme", "% of theme-related reviews that are positive")
        show_chart(heatmap(pos, cnt, GREEN_SCALE))
    with c2:
        chart_header("Negative Sentiment by Theme", "% of theme-related reviews that are negative")
        show_chart(heatmap(neg, cnt, RED_SCALE))


def render_trend(df: pd.DataFrame):
    section("Review Activity Over Time", "Whether review volume and satisfaction are rising or falling")
    if not df["Posted On"].notna().any():
        st.info("No review dates available for the current selection.")
        return
    view = st.radio("View by", ["Monthly", "Quarterly"], horizontal=True, key="trend_view")
    unit = "month" if view == "Monthly" else "quarter"
    c1, c2 = st.columns(2, gap="large")
    with c1:
        chart_header("Review Volume", f"Number of reviews by {unit}")
        show_chart(chart_review_volume(df, view))
    with c2:
        chart_header("Average Rating Over Time", f"Mean star rating by {unit} (periods with few reviews can be volatile)")
        show_chart(chart_rating_trend(df, view))


def render_feedback_tables(df: pd.DataFrame):
    section("Key Customer Feedback", f"Themes ranked by sentiment (themes with at least {MIN_THEME_REVIEWS} mentions)")
    stats = theme_stats(df)
    c1, c2 = st.columns(2, gap="large")
    with c1:
        chart_header("Customer Strengths", "Top themes by share of positive reviews")
        table = rank_themes(stats, "Positive %").dropna(subset=["Positive %"]).head(3)[["Theme", "Reviews", "Positive %"]]
        show_table(table, {"Reviews": "{:,}", "Positive %": "{:.1f}%"}, {"Positive %": ("#BFE5DD", 0, 100)})
    with c2:
        chart_header("Customer Pain Points", "Top themes by share of negative reviews")
        table = rank_themes(stats, "Negative %").dropna(subset=["Negative %"]).head(3)[["Theme", "Reviews", "Negative %"]]
        show_table(table, {"Reviews": "{:,}", "Negative %": "{:.1f}%"}, {"Negative %": ("#F6CDCF", 0, 100)})


def render_summary_table(df: pd.DataFrame):
    section("Product Summary", "All key measures in one view, ranked by average rating")
    show_table(
        product_summary(df),
        {"Reviews": "{:,}", "Average Rating": "{:.2f}", "Positive %": "{:.1f}%",
         "Neutral %": "{:.1f}%", "Negative %": "{:.1f}%"},
        {"Average Rating": ("#C9DDF3", 1, 5), "Positive %": ("#BFE5DD", 0, 100), "Negative %": ("#F6CDCF", 0, 100)},
    )


# ============================================================
# 9. MAIN APPLICATION
# ============================================================
def main():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    render_header()

    source, mtime, uploaded = find_data_source()
    if source is None and uploaded is not None:
        source = uploaded.getvalue()
    if source is None:
        st.info(f"Place **{DATA_FILE}** next to app.py, or upload it in the sidebar.")
        st.stop()

    try:
        df_all = clean_data(load_raw(source, mtime))
    except Exception as exc:
        st.error(f"Could not read the data file: {exc}")
        st.stop()

    if df_all.empty:
        st.warning("No reviews were found for the five required products.")
        st.stop()

    defaults = render_sidebar(df_all)
    df = apply_filters(df_all, defaults)
    render_sidebar_footer(defaults, len(df), len(df_all))

    if df.empty:
        st.warning("No reviews match the current filters. Try widening your selection or click Reset Filters.")
        st.stop()

    render_kpis(df)
    render_takeaways(df)
    render_product_performance(df)
    render_rating_distribution(df)
    render_sentiment(df)
    render_themes(df)
    render_theme_sentiment(df)
    render_trend(df)
    render_feedback_tables(df)
    render_summary_table(df)

    st.caption(f"Source: {DATA_FILE} · Sheet: {SHEET_NAME}")


if __name__ == "__main__":
    main()
