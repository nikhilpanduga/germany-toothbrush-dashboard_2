# app.py
# Germany Toothbrush Customer Review Dashboard
# Run with:  streamlit run app.py
# Requires:  pip install streamlit pandas plotly numpy openpyxl

# ============================================================
# 1. IMPORTS
# ============================================================
import io
import re
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# 2. CONFIGURATION
# ============================================================
st.set_page_config(
    page_title="Germany Toothbrush Review Dashboard",
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

# Colour palette
NAVY = "#14325C"
BLUE = "#2F6DB5"
LIGHT_BLUE = "#9CC3EC"
GREY = "#8A94A6"
TEXT = "#2B3445"
GRID = "#E8EDF4"
POS_COLOR = "#2E9E8A"
NEU_COLOR = "#B9C2D0"
NEG_COLOR = "#D9646A"
SENT_COLORS = {"positive": POS_COLOR, "neutral": NEU_COLOR, "negative": NEG_COLOR}
PRODUCT_COLORS = {
    "curaprox 5460": "#14325C",
    "Dr. Best Clean Pro Zwischenzahn": "#2F6DB5",
    "elmex expert precision": "#6FA8DC",
    "elmex InterX": "#8A94A6",
    "meridol base": "#E39B4B",
}
BLUE_SCALE = [[0, "#EEF4FB"], [1, "#5B9BD5"]]
GREEN_SCALE = [[0, "#EAF6F3"], [1, "#5DBBA8"]]
RED_SCALE = [[0, "#FBEEEE"], [1, "#E58A8F"]]

CHART_HEIGHT = 380

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background-color: #F6F8FB; }
.block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1400px; }
footer, #MainMenu { visibility: hidden; }

section[data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #E8EDF4; }

.dash-header { padding: 6px 0 18px 0; }
.dash-title { font-size: 2rem; font-weight: 700; color: #14325C; margin: 0; letter-spacing: -0.5px; }
.dash-subtitle { font-size: 1rem; color: #6B7689; margin-top: 4px; }

.kpi-card {
    background: #FFFFFF; border-radius: 16px; padding: 20px 22px;
    box-shadow: 0 1px 3px rgba(20, 50, 92, 0.06); border: 1px solid #EDF1F7;
    border-top: 4px solid var(--accent);
}
.kpi-label { font-size: 0.78rem; font-weight: 600; color: #7A8498; text-transform: uppercase; letter-spacing: 0.6px; }
.kpi-value { font-size: 2.2rem; font-weight: 700; color: #14325C; line-height: 1.2; margin-top: 6px; }

.section-title {
    font-size: 1.3rem; font-weight: 600; color: #14325C;
    margin: 40px 0 4px 0; padding-bottom: 8px; border-bottom: 2px solid #E3EAF3;
}
.chart-title { font-size: 0.95rem; font-weight: 600; color: #2B3445; margin: 14px 0 0 0; }

div[data-testid="stPlotlyChart"] {
    background: #FFFFFF; border-radius: 14px; padding: 8px;
    border: 1px solid #EDF1F7; box-shadow: 0 1px 3px rgba(20, 50, 92, 0.05);
}
div[data-testid="stDataFrame"] { border-radius: 12px; overflow: hidden; border: 1px solid #EDF1F7; }
</style>
"""


# ============================================================
# 3. DATA LOADING
# ============================================================
@st.cache_data(show_spinner="Loading data...")
def load_raw(source) -> pd.DataFrame:
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    return pd.read_excel(source, sheet_name=SHEET_NAME)


def find_data_source():
    """Look for the Excel file next to app.py, then in the working directory."""
    for folder in (Path(__file__).parent, Path.cwd()):
        path = folder / DATA_FILE
        if path.exists():
            return str(path), None
    return None, st.sidebar.file_uploader("Upload the Excel data file", type=["xlsx"])


# ============================================================
# 4. DATA CLEANING
# ============================================================
def clean_text_key(value) -> str:
    return re.sub(r"\s+", " ", str(value)).strip()


def map_product(row_text: str) -> str | None:
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

    # Clean column names (trim, collapse spaces, fix case against expected names)
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
    df = df.drop_duplicates(subset=dedupe_cols).reset_index(drop=True)
    return df


# ============================================================
# 5. FILTER FUNCTIONS
# ============================================================
def default_filters(df: pd.DataFrame) -> dict:
    dates = df["Posted On"].dropna()
    min_d = dates.min().date() if not dates.empty else None
    max_d = dates.max().date() if not dates.empty else None
    return {
        "f_product": ["All Products"],
        "f_country": [],
        "f_site": [],
        "f_sentiment": ["Positive", "Neutral", "Negative"],
        "f_stars": (1, 5),
        "f_dates": (min_d, max_d) if min_d else (),
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
    st.sidebar.multiselect(
        "Country", sorted(df["Country"].unique()), key="f_country", placeholder="All countries"
    )
    st.sidebar.multiselect(
        "Site", sorted(df["Site"].unique()), key="f_site", placeholder="All sites"
    )
    st.sidebar.multiselect("Sentiment", ["Positive", "Neutral", "Negative"], key="f_sentiment")
    st.sidebar.slider("Star Rating", 1, 5, key="f_stars")

    if defaults["f_dates"]:
        st.sidebar.date_input(
            "Review Date",
            key="f_dates",
            min_value=defaults["f_dates"][0],
            max_value=defaults["f_dates"][1],
        )
    st.sidebar.button("Reset Filters", on_click=reset_filters, args=(defaults,), use_container_width=True)

    return {"defaults": defaults}


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
    if len(sentiments) < 3:
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
# 6. KPI & AGGREGATION FUNCTIONS
# ============================================================
def sentiment_share(df: pd.DataFrame, label: str) -> float:
    valid = df["Sentiment"].notna().sum()
    return (df["Sentiment"] == label).sum() / valid * 100 if valid else 0.0


def compute_kpis(df: pd.DataFrame) -> dict:
    return {
        "total": len(df),
        "avg_rating": df["Stars"].mean() if df["Stars"].notna().any() else np.nan,
        "positive": sentiment_share(df, "positive"),
        "negative": sentiment_share(df, "negative"),
        "neutral": sentiment_share(df, "neutral"),
    }


def ordered_products(df: pd.DataFrame) -> list:
    return [p for p in PRODUCTS if p in set(df["Product"])]


def theme_penetration(df: pd.DataFrame) -> pd.DataFrame:
    """Reviews mentioning theme / total filtered reviews x 100."""
    total = len(df)
    rows = []
    for theme, col in THEMES.items():
        mentions = (df[col] == "Yes").sum()
        rows.append({"Theme": theme, "Reviews": mentions, "Pct": mentions / total * 100 if total else 0})
    return pd.DataFrame(rows).sort_values("Pct", ascending=False)


def product_theme_matrix(df: pd.DataFrame, mode: str = "penetration") -> pd.DataFrame:
    """Product x Theme matrix. mode: penetration | positive | negative."""
    products = ordered_products(df)
    matrix = pd.DataFrame(index=products, columns=list(THEMES), dtype=float)
    for product in products:
        sub = df[df["Product"] == product]
        for theme, col in THEMES.items():
            mentioned = sub[sub[col] == "Yes"]
            if mode == "penetration":
                matrix.loc[product, theme] = len(mentioned) / len(sub) * 100 if len(sub) else np.nan
            else:
                label = "positive" if mode == "positive" else "negative"
                matrix.loc[product, theme] = (
                    (mentioned["Sentiment"] == label).sum() / len(mentioned) * 100 if len(mentioned) else np.nan
                )
    return matrix


def theme_sentiment_table(df: pd.DataFrame, label: str) -> pd.DataFrame:
    rows = []
    for theme, col in THEMES.items():
        mentioned = df[df[col] == "Yes"]
        n = len(mentioned)
        pct = (mentioned["Sentiment"] == label).sum() / n * 100 if n else np.nan
        rows.append({"Theme": theme, "Reviews": n, f"{label.title()} %": pct})
    out = pd.DataFrame(rows).dropna()
    return out.sort_values([f"{label.title()} %", "Reviews"], ascending=False).head(3).reset_index(drop=True)


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


# ============================================================
# 7. CHART FUNCTIONS
# ============================================================
def style_fig(fig: go.Figure, height: int = CHART_HEIGHT, legend: bool = False) -> go.Figure:
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=20, t=20, b=10),
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(family="Inter, sans-serif", color=TEXT, size=12),
        showlegend=legend,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0, title=None),
        hoverlabel=dict(bgcolor="white", font_size=12),
    )
    fig.update_xaxes(gridcolor=GRID, zeroline=False, showline=False)
    fig.update_yaxes(gridcolor=GRID, zeroline=False, showline=False)
    return fig


def bar_by_product(df_agg: pd.DataFrame, value_col: str, xtitle: str, fmt: str, x_range=None) -> go.Figure:
    data = df_agg.sort_values(value_col, ascending=True)
    fig = go.Figure(go.Bar(
        x=data[value_col], y=data["Product"], orientation="h",
        marker_color=[PRODUCT_COLORS.get(p, BLUE) for p in data["Product"]],
        text=data[value_col], texttemplate=fmt, textposition="outside", cliponaxis=False,
        hovertemplate="%{y}<br>" + xtitle + ": %{x}<extra></extra>",
    ))
    fig.update_xaxes(title=xtitle, range=x_range)
    fig.update_yaxes(title=None)
    return style_fig(fig)


def chart_reviews_by_product(df: pd.DataFrame) -> go.Figure:
    agg = df.groupby("Product").size().reset_index(name="Reviews")
    fig = bar_by_product(agg, "Reviews", "Number of Reviews", "%{text:,}")
    fig.update_xaxes(range=[0, agg["Reviews"].max() * 1.18])
    return fig


def chart_rating_by_product(df: pd.DataFrame) -> go.Figure:
    agg = df.groupby("Product")["Stars"].mean().reset_index(name="Rating")
    return bar_by_product(agg, "Rating", "Average Rating", "%{text:.2f}", x_range=[0, 5])


def chart_sentiment_by_product(df: pd.DataFrame) -> go.Figure:
    products = ordered_products(df)[::-1]
    fig = go.Figure()
    for sentiment in SENTIMENTS:
        values = []
        for p in products:
            sub = df[df["Product"] == p]
            values.append(sentiment_share(sub, sentiment))
        fig.add_trace(go.Bar(
            name=sentiment.title(), x=values, y=products, orientation="h",
            marker_color=SENT_COLORS[sentiment],
            text=values, texttemplate="%{text:.0f}%", textposition="inside",
            textfont=dict(color="white" if sentiment != "neutral" else TEXT),
            hovertemplate="%{y}<br>" + sentiment.title() + ": %{x:.1f}%<extra></extra>",
        ))
    fig.update_layout(barmode="stack")
    fig.update_xaxes(title="Share of Reviews (%)", range=[0, 100])
    fig.update_yaxes(title=None)
    return style_fig(fig, legend=True)


def chart_overall_sentiment(df: pd.DataFrame) -> go.Figure:
    counts = df["Sentiment"].value_counts().reindex(SENTIMENTS).fillna(0)
    fig = go.Figure(go.Pie(
        labels=[s.title() for s in counts.index], values=counts.values, hole=0.62,
        marker=dict(colors=[SENT_COLORS[s] for s in counts.index], line=dict(color="white", width=2)),
        textinfo="percent", textfont=dict(size=13, color="white"), sort=False,
        hovertemplate="%{label}: %{value:,} reviews (%{percent})<extra></extra>",
    ))
    fig.add_annotation(
        text=f"<b>{int(counts.sum()):,}</b><br><span style='font-size:12px;color:{GREY}'>reviews</span>",
        showarrow=False, font=dict(size=22, color=NAVY),
    )
    return style_fig(fig, legend=True)


def chart_theme_importance(df: pd.DataFrame) -> go.Figure:
    data = theme_penetration(df).sort_values("Pct", ascending=True)
    fig = go.Figure(go.Bar(
        x=data["Pct"], y=data["Theme"], orientation="h", marker_color=BLUE,
        text=data["Pct"], texttemplate="%{text:.0f}%", textposition="outside", cliponaxis=False,
        hovertemplate="%{y}: %{x:.1f}%<extra></extra>",
    ))
    fig.update_xaxes(title="Reviews Mentioning Theme (%)", range=[0, max(data["Pct"].max() * 1.2, 10)])
    fig.update_yaxes(title=None)
    return style_fig(fig)


def heatmap(matrix: pd.DataFrame, colorscale, zmax=None) -> go.Figure:
    fig = go.Figure(go.Heatmap(
        z=matrix.values, x=matrix.columns.tolist(), y=matrix.index.tolist(),
        colorscale=colorscale, zmin=0, zmax=zmax if zmax else 100, showscale=False,
        text=matrix.values, texttemplate="%{text:.0f}%", textfont=dict(size=13, color=NAVY),
        xgap=3, ygap=3, hoverongaps=False,
        hovertemplate="%{y}<br>%{x}: %{z:.1f}%<extra></extra>",
    ))
    fig.update_yaxes(autorange="reversed", title=None, showgrid=False)
    fig.update_xaxes(title=None, showgrid=False, side="top", tickangle=0)
    fig.update_layout(margin=dict(l=10, r=10, t=60, b=10))
    return style_fig(fig, height=CHART_HEIGHT).update_layout(margin=dict(l=10, r=10, t=60, b=10))


def chart_rating_distribution(df: pd.DataFrame) -> go.Figure:
    fig = go.Figure()
    stars = df.dropna(subset=["Stars"]).copy()
    stars["Star"] = stars["Stars"].round().clip(1, 5).astype(int)
    for product in ordered_products(stars):
        sub = stars[stars["Product"] == product]
        pct = sub["Star"].value_counts(normalize=True).reindex(range(1, 6)).fillna(0) * 100
        fig.add_trace(go.Bar(
            name=product, x=[f"{i} Star" + ("s" if i > 1 else "") for i in range(1, 6)], y=pct.values,
            marker_color=PRODUCT_COLORS.get(product, BLUE),
            text=pct.values, texttemplate="%{text:.0f}%", textposition="outside", cliponaxis=False,
            hovertemplate=product + "<br>%{x}: %{y:.1f}%<extra></extra>",
        ))
    fig.update_layout(barmode="group", bargap=0.2)
    fig.update_yaxes(title="Share of Reviews (%)", range=[0, 100])
    fig.update_xaxes(title=None)
    return style_fig(fig, legend=True)


def chart_review_trend(df: pd.DataFrame) -> go.Figure:
    dated = df.dropna(subset=["Posted On"]).copy()
    dated["Month"] = dated["Posted On"].dt.to_period("M").dt.to_timestamp()
    full_range = pd.date_range(dated["Month"].min(), dated["Month"].max(), freq="MS")
    fig = go.Figure()
    for product in ordered_products(dated):
        counts = (dated[dated["Product"] == product].groupby("Month").size()
                  .reindex(full_range, fill_value=0))
        fig.add_trace(go.Scatter(
            name=product, x=counts.index, y=counts.values, mode="lines+markers",
            line=dict(color=PRODUCT_COLORS.get(product, BLUE), width=2.5), marker=dict(size=5),
            hovertemplate=product + "<br>%{x|%b %Y}: %{y} reviews<extra></extra>",
        ))
    fig.update_xaxes(title="Month", tickformat="%b %Y")
    fig.update_yaxes(title="Number of Reviews", rangemode="tozero")
    return style_fig(fig, legend=True)


# ============================================================
# 8. DASHBOARD LAYOUT HELPERS
# ============================================================
def show_chart(fig: go.Figure):
    try:
        st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})
    except TypeError:  # older Streamlit versions
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})


def show_table(df: pd.DataFrame, column_config: dict | None = None):
    kwargs = dict(hide_index=True, column_config=column_config or {})
    try:
        st.dataframe(df, width="stretch", **kwargs)
    except TypeError:
        st.dataframe(df, use_container_width=True, **kwargs)


def section(title: str):
    st.markdown(f'<div class="section-title">{title}</div>', unsafe_allow_html=True)


def chart_title(title: str):
    st.markdown(f'<div class="chart-title">{title}</div>', unsafe_allow_html=True)


def kpi_card(label: str, value: str, accent: str) -> str:
    return (
        f'<div class="kpi-card" style="--accent:{accent}">'
        f'<div class="kpi-label">{label}</div><div class="kpi-value">{value}</div></div>'
    )


def render_header():
    st.markdown(
        """
        <div class="dash-header">
            <p class="dash-title">Germany Toothbrush Customer Review Dashboard</p>
            <p class="dash-subtitle">Customer perception, ratings, sentiment and key product themes</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_kpis(df: pd.DataFrame):
    k = compute_kpis(df)
    avg = f"{k['avg_rating']:.2f} ★" if pd.notna(k["avg_rating"]) else "–"
    cards = [
        ("Total Reviews", f"{k['total']:,}", NAVY),
        ("Average Rating", avg, BLUE),
        ("Positive Sentiment", f"{k['positive']:.1f}%", POS_COLOR),
        ("Negative Sentiment", f"{k['negative']:.1f}%", NEG_COLOR),
        ("Neutral Sentiment", f"{k['neutral']:.1f}%", NEU_COLOR),
    ]
    for col, (label, value, accent) in zip(st.columns(5), cards):
        col.markdown(kpi_card(label, value, accent), unsafe_allow_html=True)


def render_product_performance(df: pd.DataFrame):
    section("Product Performance")
    c1, c2 = st.columns(2)
    with c1:
        chart_title("Reviews by Product")
        show_chart(chart_reviews_by_product(df))
    with c2:
        chart_title("Average Rating by Product")
        show_chart(chart_rating_by_product(df))


def render_sentiment(df: pd.DataFrame):
    section("Customer Sentiment")
    c1, c2 = st.columns([3, 2])
    with c1:
        chart_title("Sentiment by Product")
        show_chart(chart_sentiment_by_product(df))
    with c2:
        chart_title("Overall Sentiment")
        show_chart(chart_overall_sentiment(df))


def render_themes(df: pd.DataFrame):
    section("What Customers Talk About")
    c1, c2 = st.columns([2, 3])
    with c1:
        chart_title("Customer Theme Importance")
        show_chart(chart_theme_importance(df))
    with c2:
        chart_title("Product vs Customer Themes (% of reviews mentioning theme)")
        matrix = product_theme_matrix(df, "penetration")
        show_chart(heatmap(matrix, BLUE_SCALE, zmax=max(float(np.nanmax(matrix.values)), 10)))


def render_theme_sentiment(df: pd.DataFrame):
    section("Customer Sentiment by Theme")
    c1, c2 = st.columns(2)
    with c1:
        chart_title("Positive Sentiment by Theme")
        show_chart(heatmap(product_theme_matrix(df, "positive"), GREEN_SCALE))
    with c2:
        chart_title("Negative Sentiment by Theme")
        show_chart(heatmap(product_theme_matrix(df, "negative"), RED_SCALE))


def render_rating_distribution(df: pd.DataFrame):
    section("Customer Rating Distribution")
    show_chart(chart_rating_distribution(df))


def render_trend(df: pd.DataFrame):
    section("Review Activity Over Time")
    chart_title("Number of Reviews by Month")
    if df["Posted On"].notna().any():
        show_chart(chart_review_trend(df))
    else:
        st.info("No review dates available for the current selection.")


def render_feedback_tables(df: pd.DataFrame):
    section("Key Customer Feedback")
    c1, c2 = st.columns(2)
    pct_cfg = lambda name: {name: st.column_config.NumberColumn(name, format="%.1f%%")}
    with c1:
        chart_title("Customer Strengths")
        show_table(theme_sentiment_table(df, "positive"), pct_cfg("Positive %"))
    with c2:
        chart_title("Customer Pain Points")
        show_table(theme_sentiment_table(df, "negative"), pct_cfg("Negative %"))


def render_summary_table(df: pd.DataFrame):
    section("Product Summary")
    summary = product_summary(df)
    config = {
        "Average Rating": st.column_config.ProgressColumn(
            "Average Rating", min_value=0, max_value=5, format="%.2f"),
        "Positive %": st.column_config.ProgressColumn(
            "Positive %", min_value=0, max_value=100, format="%.1f%%"),
        "Neutral %": st.column_config.NumberColumn("Neutral %", format="%.1f%%"),
        "Negative %": st.column_config.ProgressColumn(
            "Negative %", min_value=0, max_value=100, format="%.1f%%"),
    }
    show_table(summary, config)


# ============================================================
# 9. MAIN APPLICATION
# ============================================================
def main():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    render_header()

    source, uploaded = find_data_source()
    if source is None and uploaded is not None:
        source = uploaded.getvalue()
    if source is None:
        st.info(f"Place **{DATA_FILE}** next to app.py, or upload it in the sidebar.")
        st.stop()

    try:
        df_all = clean_data(load_raw(source))
    except Exception as exc:
        st.error(f"Could not read the data file: {exc}")
        st.stop()

    if df_all.empty:
        st.warning("No reviews were found for the five required products.")
        st.stop()

    ctx = render_sidebar(df_all)
    df = apply_filters(df_all, ctx["defaults"])

    if df.empty:
        st.warning("No reviews match the current filters. Try widening your selection or click Reset Filters.")
        st.stop()

    render_kpis(df)
    render_product_performance(df)
    render_sentiment(df)
    render_themes(df)
    render_theme_sentiment(df)
    render_rating_distribution(df)
    render_trend(df)
    render_feedback_tables(df)
    render_summary_table(df)


if __name__ == "__main__":
    main()
