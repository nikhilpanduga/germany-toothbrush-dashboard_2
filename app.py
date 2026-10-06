# ============================================================
# Germany Toothbrush Customer Review Dashboard
# Streamlit + Pandas + Plotly
# ============================================================

from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Germany Toothbrush Customer Review Dashboard",
    page_icon="🪥",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown(
    """
    <style>

    /* Main page */
    .main {
        background-color: #F8FAFC;
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        max-width: 1500px;
    }

    /* Header */
    .dashboard-title {
        font-size: 30px;
        font-weight: 700;
        color: #173B63;
        margin-bottom: 4px;
    }

    .dashboard-subtitle {
        font-size: 15px;
        color: #64748B;
        margin-bottom: 25px;
    }

    /* Section headings */
    .section-title {
        font-size: 21px;
        font-weight: 700;
        color: #173B63;
        margin-top: 25px;
        margin-bottom: 12px;
    }

    /* KPI cards */
    div[data-testid="stMetric"] {
        background-color: white;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 15px 16px;
        min-height: 105px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.04);
    }

    div[data-testid="stMetricLabel"] {
        color: #64748B;
        font-size: 13px;
        font-weight: 500;
    }

    div[data-testid="stMetricValue"] {
        color: #173B63;
        font-size: 25px;
        font-weight: 700;
    }

    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: #F8FAFC;
    }

    /* Tables */
    .dataframe {
        font-size: 13px;
    }

    /* Reduce excessive spacing */
    div[data-testid="stVerticalBlock"] > div {
        gap: 0.5rem;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# CONSTANTS
# ============================================================

DATA_FILE = Path(__file__).parent / "Germany_five toothbrush data.xlsx"
SHEET_NAME = "Germany"

PRODUCTS = [
    "curaprox 5460",
    "Dr. Best Clean Pro Zwischenzahn",
    "elmex expert precision",
    "elmex InterX",
    "meridol base",
]

THEMES = [
    "Cleaning Performance",
    "Brushing Comfort",
    "Design & Ease of Use",
    "Quality & Durability",
    "Features & Technology",
    "Price & Value",
]

THEME_COLUMNS = {
    "Cleaning Performance": "Cleaning Performance_Present",
    "Brushing Comfort": "Brushing Comfort_Present",
    "Design & Ease of Use": "Design & Ease of Use_Present",
    "Quality & Durability": "Quality & Durability_Present",
    "Features & Technology": "Features & Technology_Present",
    "Price & Value": "Price & Value_Present",
}

SENTIMENT_ORDER = [
    "positive",
    "neutral",
    "negative"
]

SENTIMENT_LABELS = {
    "positive": "Positive",
    "neutral": "Neutral",
    "negative": "Negative",
}


# ============================================================
# DATA LOADING
# ============================================================

@st.cache_data
def load_data(file_path):
    df = pd.read_excel(
        file_path,
        sheet_name=SHEET_NAME
    )

    return df


# ============================================================
# DATA CLEANING
# ============================================================

def standardize_product(value):
    """
    Standardize product names into the five dashboard products.
    """

    if pd.isna(value):
        return np.nan

    text = str(value).strip().lower()

    if "curaprox" in text and "5460" in text:
        return "curaprox 5460"

    if "dr. best" in text and "clean pro" in text:
        return "Dr. Best Clean Pro Zwischenzahn"

    if "elmex" in text and "expert" in text and "precision" in text:
        return "elmex expert precision"

    if "elmex" in text and "interx" in text:
        return "elmex InterX"

    if "meridol" in text and "base" in text:
        return "meridol base"

    return np.nan


def normalize_sentiment(value):
    """
    Normalize sentiment values.
    """

    if pd.isna(value):
        return np.nan

    text = str(value).strip().lower()

    if text in ["positive", "pos"]:
        return "positive"

    if text in ["neutral", "neu"]:
        return "neutral"

    if text in ["negative", "neg"]:
        return "negative"

    return np.nan


def normalize_theme_value(value):
    """
    Standardize theme presence values.

    Missing values remain missing.
    """

    if pd.isna(value):
        return np.nan

    text = str(value).strip().lower()

    if text in [
        "yes",
        "y",
        "true",
        "1",
        "present"
    ]:
        return "Yes"

    if text in [
        "no",
        "n",
        "false",
        "0",
        "not present"
    ]:
        return "No"

    return np.nan


def clean_data(df):
    df = df.copy()

    # --------------------------------------------------------
    # Remove duplicate rows
    # --------------------------------------------------------

    df = df.drop_duplicates().reset_index(drop=True)

    # --------------------------------------------------------
    # Standardize product
    # --------------------------------------------------------

    if "Product Name" in df.columns:
        df["Product Name"] = df["Product Name"].apply(
            standardize_product
        )

    # --------------------------------------------------------
    # Keep only the five requested products
    # --------------------------------------------------------

    df = df[
        df["Product Name"].isin(PRODUCTS)
    ].copy()

    # --------------------------------------------------------
    # Sentiment
    # --------------------------------------------------------

    if "Sentiment" in df.columns:
        df["Sentiment"] = df["Sentiment"].apply(
            normalize_sentiment
        )

    # --------------------------------------------------------
    # Stars
    # --------------------------------------------------------

    if "Stars" in df.columns:
        df["Stars"] = pd.to_numeric(
            df["Stars"],
            errors="coerce"
        )

        df["Stars"] = df["Stars"].clip(
            lower=1,
            upper=5
        )

    # --------------------------------------------------------
    # Posted On
    # --------------------------------------------------------

    if "Posted On" in df.columns:
        df["Posted On"] = pd.to_datetime(
            df["Posted On"],
            errors="coerce"
        )

    # --------------------------------------------------------
    # Theme columns
    # --------------------------------------------------------

    for theme, column in THEME_COLUMNS.items():

        if column in df.columns:

            df[column] = df[column].apply(
                normalize_theme_value
            )

    return df.reset_index(drop=True)


# ============================================================
# LOAD AND CLEAN DATA
# ============================================================

if not DATA_FILE.exists():

    st.error(
        f"Data file not found:\n\n{DATA_FILE}"
    )

    st.stop()


raw_df = load_data(DATA_FILE)

df = clean_data(raw_df)


# ============================================================
# HEADER
# ============================================================

st.markdown(
    '<div class="dashboard-title">'
    'Germany Toothbrush Customer Review Dashboard'
    '</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="dashboard-subtitle">'
    'Customer perception, ratings, sentiment and key product themes'
    '</div>',
    unsafe_allow_html=True
)


# ============================================================
# SIDEBAR FILTERS
# ============================================================

st.sidebar.markdown("## Filters")


# ------------------------------------------------------------
# Product filter
# ------------------------------------------------------------

product_options = ["All Products"] + PRODUCTS

selected_products = st.sidebar.multiselect(
    "Product",
    options=product_options,
    default=["All Products"]
)


# ------------------------------------------------------------
# Brand filter
# ------------------------------------------------------------

brand_values = sorted(
    df["Brand"]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

selected_brands = st.sidebar.multiselect(
    "Brand",
    options=brand_values,
    default=brand_values
)


# ------------------------------------------------------------
# Site filter
# ------------------------------------------------------------

site_values = sorted(
    df["Site"]
    .dropna()
    .astype(str)
    .unique()
    .tolist()
)

selected_sites = st.sidebar.multiselect(
    "Site",
    options=site_values,
    default=site_values
)


# ------------------------------------------------------------
# Sentiment filter
# ------------------------------------------------------------

selected_sentiments = st.sidebar.multiselect(
    "Sentiment",
    options=SENTIMENT_ORDER,
    default=SENTIMENT_ORDER,
    format_func=lambda x: SENTIMENT_LABELS[x]
)


# ------------------------------------------------------------
# Star rating
# ------------------------------------------------------------

selected_stars = st.sidebar.slider(
    "Star Rating",
    min_value=1,
    max_value=5,
    value=(1, 5),
    step=1
)


# ------------------------------------------------------------
# Review date
# ------------------------------------------------------------

valid_dates = df["Posted On"].dropna()

if len(valid_dates) > 0:

    min_date = valid_dates.min().date()
    max_date = valid_dates.max().date()

    selected_dates = st.sidebar.date_input(
        "Review Date",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )

else:

    selected_dates = None


# ------------------------------------------------------------
# Reset button
# ------------------------------------------------------------

if st.sidebar.button(
    "Reset Filters",
    use_container_width=True
):

    st.rerun()


# ============================================================
# APPLY FILTERS
# ============================================================

filtered_df = df.copy()


# Product
if (
    selected_products
    and "All Products" not in selected_products
):

    filtered_df = filtered_df[
        filtered_df["Product Name"].isin(
            selected_products
        )
    ]


# Brand
if selected_brands:

    filtered_df = filtered_df[
        filtered_df["Brand"].astype(str).isin(
            selected_brands
        )
    ]


# Site
if selected_sites:

    filtered_df = filtered_df[
        filtered_df["Site"].astype(str).isin(
            selected_sites
        )
    ]


# Sentiment
if selected_sentiments:

    filtered_df = filtered_df[
        filtered_df["Sentiment"].isin(
            selected_sentiments
        )
    ]


# Stars
filtered_df = filtered_df[
    filtered_df["Stars"].between(
        selected_stars[0],
        selected_stars[1],
        inclusive="both"
    )
]


# Date
if selected_dates is not None:

    if isinstance(selected_dates, tuple) and len(selected_dates) == 2:

        start_date = pd.Timestamp(
            selected_dates[0]
        )

        end_date = pd.Timestamp(
            selected_dates[1]
        ) + pd.Timedelta(days=1)

        filtered_df = filtered_df[
            (filtered_df["Posted On"] >= start_date)
            &
            (filtered_df["Posted On"] < end_date)
        ]


# ============================================================
# EMPTY DATA CHECK
# ============================================================

if filtered_df.empty:

    st.warning(
        "No reviews match the selected filters. "
        "Please adjust the filters."
    )

    st.stop()


# ============================================================
# COMMON CHART STYLING
# ============================================================

def style_chart(
    fig,
    title,
    height=420,
    show_legend=True
):
    """
    Common Plotly layout.

    The extra top margin and separated legend prevent
    chart titles and legends from overlapping.
    """

    fig.update_layout(

        title=dict(
            text=title,
            x=0,
            xanchor="left",
            y=0.98,
            yanchor="top",
            font=dict(
                size=18,
                color="#173B63"
            )
        ),

        margin=dict(
            l=30,
            r=30,
            t=115,
            b=65
        ),

        height=height,

        paper_bgcolor="white",
        plot_bgcolor="white",

        font=dict(
            family="Arial",
            color="#334155"
        ),

        hoverlabel=dict(
            bgcolor="white"
        )
    )

    if show_legend:

        fig.update_layout(
            legend=dict(
                orientation="h",
                y=0.88,
                yanchor="top",
                x=0,
                xanchor="left",
                font=dict(size=11)
            )
        )

    return fig


def format_axes(fig):

    fig.update_xaxes(
        showgrid=True,
        gridcolor="#E2E8F0",
        zeroline=False
    )

    fig.update_yaxes(
        showgrid=False,
        zeroline=False
    )

    return fig


# ============================================================
# KPI CALCULATIONS
# ============================================================

total_reviews = len(filtered_df)

avg_rating = (
    filtered_df["Stars"].mean()
    if total_reviews > 0
    else 0
)

positive_pct = (
    filtered_df["Sentiment"].eq("positive").mean() * 100
    if total_reviews > 0
    else 0
)

negative_pct = (
    filtered_df["Sentiment"].eq("negative").mean() * 100
    if total_reviews > 0
    else 0
)

neutral_pct = (
    filtered_df["Sentiment"].eq("neutral").mean() * 100
    if total_reviews > 0
    else 0
)

five_star_pct = (
    filtered_df["Stars"].eq(5).mean() * 100
    if total_reviews > 0
    else 0
)

product_count = (
    filtered_df["Product Name"].nunique()
)

month_count = (
    filtered_df["Posted On"]
    .dropna()
    .dt.to_period("M")
    .nunique()
)

avg_reviews_month = (
    total_reviews / month_count
    if month_count > 0
    else 0
)

if total_reviews > 0:

    most_reviewed_product = (
        filtered_df["Product Name"]
        .value_counts()
        .idxmax()
    )

else:

    most_reviewed_product = "N/A"


# ============================================================
# KPI CARDS
# ============================================================

st.markdown(
    '<div class="section-title">Key Performance Indicators</div>',
    unsafe_allow_html=True
)


# First row
k1, k2, k3, k4 = st.columns(4)


with k1:

    st.metric(
        "Total Reviews",
        f"{total_reviews:,}"
    )


with k2:

    st.metric(
        "Average Rating",
        f"{avg_rating:.2f} / 5"
    )


with k3:

    st.metric(
        "Positive Sentiment",
        f"{positive_pct:.1f}%"
    )


with k4:

    st.metric(
        "Negative Sentiment",
        f"{negative_pct:.1f}%"
    )


# Second row
k5, k6, k7, k8 = st.columns(4)


with k5:

    st.metric(
        "Neutral Sentiment",
        f"{neutral_pct:.1f}%"
    )


with k6:

    st.metric(
        "5-Star Reviews",
        f"{five_star_pct:.1f}%"
    )


with k7:

    st.metric(
        "Avg Reviews / Month",
        f"{avg_reviews_month:,.0f}"
    )


with k8:

    st.metric(
        "Products",
        f"{product_count}"
    )


# ============================================================
# PRODUCT PERFORMANCE
# ============================================================

st.markdown(
    '<div class="section-title">Product Performance</div>',
    unsafe_allow_html=True
)


product_summary = (
    filtered_df
    .groupby("Product Name")
    .agg(
        Reviews=("Product Name", "size"),
        Average_Rating=("Stars", "mean")
    )
    .reset_index()
)

product_summary["Product"] = pd.Categorical(
    product_summary["Product Name"],
    categories=PRODUCTS,
    ordered=True
)

product_summary = product_summary.sort_values(
    "Reviews",
    ascending=True
)


col1, col2 = st.columns(2)


# ------------------------------------------------------------
# Reviews by Product
# ------------------------------------------------------------

with col1:

    fig_reviews = px.bar(
        product_summary,
        x="Reviews",
        y="Product Name",
        orientation="h",
        text="Reviews"
    )

    fig_reviews.update_traces(
        textposition="outside",
        hovertemplate=(
            "<b>%{y}</b><br>"
            "Reviews: %{x:,}<extra></extra>"
        )
    )

    fig_reviews = style_chart(
        fig_reviews,
        "Reviews by Product",
        show_legend=False
    )

    fig_reviews.update_xaxes(
        title="Number of Reviews"
    )

    fig_reviews.update_yaxes(
        title=""
    )

    fig_reviews = format_axes(fig_reviews)

    st.plotly_chart(
        fig_reviews,
        use_container_width=True
    )


# ------------------------------------------------------------
# Average Rating by Product
# ------------------------------------------------------------

rating_product = product_summary.sort_values(
    "Average_Rating",
    ascending=True
)

with col2:

    fig_rating = px.bar(
        rating_product,
        x="Average_Rating",
        y="Product Name",
        orientation="h",
        text=rating_product["Average_Rating"].round(2)
    )

    fig_rating.update_traces(
        textposition="outside",
        hovertemplate=(
            "<b>%{y}</b><br>"
            "Average Rating: %{x:.2f}<extra></extra>"
        )
    )

    fig_rating.update_xaxes(
        range=[0, 5],
        title="Average Rating",
        dtick=1
    )

    fig_rating.update_yaxes(
        title=""
    )

    fig_rating = style_chart(
        fig_rating,
        "Average Rating by Product",
        show_legend=False
    )

    fig_rating = format_axes(fig_rating)

    st.plotly_chart(
        fig_rating,
        use_container_width=True
    )


# ============================================================
# CUSTOMER SENTIMENT
# ============================================================

st.markdown(
    '<div class="section-title">Customer Sentiment</div>',
    unsafe_allow_html=True
)


col1, col2 = st.columns(2)


# ------------------------------------------------------------
# Sentiment by Product
# ------------------------------------------------------------

with col1:

    sentiment_product = (
        filtered_df
        .groupby(
            ["Product Name", "Sentiment"]
        )
        .size()
        .reset_index(name="Reviews")
    )

    sentiment_totals = (
        filtered_df
        .groupby("Product Name")
        .size()
        .reset_index(name="Total")
    )

    sentiment_product = sentiment_product.merge(
        sentiment_totals,
        on="Product Name",
        how="left"
    )

    sentiment_product["Percentage"] = (
        sentiment_product["Reviews"]
        / sentiment_product["Total"]
        * 100
    )

    sentiment_product["Sentiment"] = pd.Categorical(
        sentiment_product["Sentiment"],
        categories=SENTIMENT_ORDER,
        ordered=True
    )

    fig_sentiment_product = px.bar(
        sentiment_product,
        x="Percentage",
        y="Product Name",
        color="Sentiment",
        orientation="h",
        category_orders={
            "Sentiment": SENTIMENT_ORDER
        },
        labels={
            "Percentage": "Share of Reviews (%)",
            "Product Name": ""
        },
        text=None
    )

    fig_sentiment_product.update_layout(
        barmode="relative"
    )

    fig_sentiment_product.update_xaxes(
        range=[0, 100],
        ticksuffix="%"
    )

    fig_sentiment_product.update_traces(
        hovertemplate=(
            "<b>%{y}</b><br>"
            "Sentiment: %{fullData.name}<br>"
            "Share: %{x:.1f}%<extra></extra>"
        )
    )

    fig_sentiment_product = style_chart(
        fig_sentiment_product,
        "Sentiment by Product"
    )

    fig_sentiment_product = format_axes(
        fig_sentiment_product
    )

    st.plotly_chart(
        fig_sentiment_product,
        use_container_width=True
    )


# ------------------------------------------------------------
# Overall Sentiment Donut
# ------------------------------------------------------------

with col2:

    overall_sentiment = (
        filtered_df["Sentiment"]
        .value_counts()
        .reindex(
            SENTIMENT_ORDER,
            fill_value=0
        )
        .reset_index()
    )

    overall_sentiment.columns = [
        "Sentiment",
        "Reviews"
    ]

    overall_sentiment["Label"] = (
        overall_sentiment["Sentiment"]
        .map(SENTIMENT_LABELS)
    )

    fig_donut = go.Figure(
        data=[
            go.Pie(
                labels=overall_sentiment["Label"],
                values=overall_sentiment["Reviews"],
                hole=0.58,
                textinfo="percent",
                hovertemplate=(
                    "%{label}<br>"
                    "Reviews: %{value:,}<br>"
                    "Share: %{percent}"
                    "<extra></extra>"
                )
            )
        ]
    )

    fig_donut.update_layout(
        title=dict(
            text="Overall Customer Sentiment",
            x=0,
            xanchor="left",
            y=0.98,
            font=dict(
                size=18,
                color="#173B63"
            )
        ),
        height=420,
        margin=dict(
            l=20,
            r=20,
            t=80,
            b=30
        ),
        paper_bgcolor="white",
        plot_bgcolor="white",
        legend=dict(
            orientation="h",
            y=-0.05,
            x=0.5,
            xanchor="center"
        )
    )

    st.plotly_chart(
        fig_donut,
        use_container_width=True
    )


# ============================================================
# CUSTOMER THEMES
# ============================================================

st.markdown(
    '<div class="section-title">Customer Themes</div>',
    unsafe_allow_html=True
)


# ------------------------------------------------------------
# Theme importance
# ------------------------------------------------------------

theme_rows = []

for theme in THEMES:

    column = THEME_COLUMNS[theme]

    if column not in filtered_df.columns:
        continue

    mentioned = (
        filtered_df[column] == "Yes"
    ).sum()

    penetration = (
        mentioned / total_reviews * 100
        if total_reviews > 0
        else 0
    )

    theme_rows.append(
        {
            "Theme": theme,
            "Reviews": mentioned,
            "Percentage": penetration
        }
    )


theme_importance = pd.DataFrame(
    theme_rows
)

theme_importance = theme_importance.sort_values(
    "Percentage",
    ascending=True
)


col1, col2 = st.columns(2)


# ------------------------------------------------------------
# Customer Theme Importance
# ------------------------------------------------------------

with col1:

    fig_theme = px.bar(
        theme_importance,
        x="Percentage",
        y="Theme",
        orientation="h",
        text=theme_importance["Percentage"].round(1)
    )

    fig_theme.update_traces(
        texttemplate="%{text}%",
        textposition="outside",
        hovertemplate=(
            "<b>%{y}</b><br>"
            "Reviews: %{customdata[0]:,}<br>"
            "Share: %{x:.1f}%<extra></extra>"
        ),
        customdata=theme_importance[
            ["Reviews"]
        ]
    )

    fig_theme.update_xaxes(
        title="Theme Mention Rate",
        ticksuffix="%"
    )

    fig_theme.update_yaxes(
        title=""
    )

    fig_theme = style_chart(
        fig_theme,
        "Customer Theme Importance",
        show_legend=False
    )

    fig_theme = format_axes(
        fig_theme
    )

    st.plotly_chart(
        fig_theme,
        use_container_width=True
    )


# ------------------------------------------------------------
# Product × Theme Heatmap
# ------------------------------------------------------------

with col2:

    heatmap_data = []

    selected_product_list = [
        p for p in PRODUCTS
        if p in filtered_df["Product Name"].unique()
    ]

    for product in selected_product_list:

        product_df = filtered_df[
            filtered_df["Product Name"] == product
        ]

        product_total = len(product_df)

        for theme in THEMES:

            column = THEME_COLUMNS[theme]

            if column not in product_df.columns:
                percentage = 0

            else:

                mentions = (
                    product_df[column] == "Yes"
                ).sum()

                percentage = (
                    mentions / product_total * 100
                    if product_total > 0
                    else 0
                )

            heatmap_data.append(
                {
                    "Product": product,
                    "Theme": theme,
                    "Percentage": percentage
                }
            )

    heatmap_df = pd.DataFrame(
        heatmap_data
    )

    heatmap_pivot = heatmap_df.pivot(
        index="Product",
        columns="Theme",
        values="Percentage"
    )

    heatmap_pivot = heatmap_pivot.reindex(
        columns=THEMES
    )

    fig_heatmap = px.imshow(
        heatmap_pivot,
        text_auto=".1f",
        aspect="auto",
        labels={
            "x": "Theme",
            "y": "Product",
            "color": "%"
        }
    )

    fig_heatmap.update_traces(
        hovertemplate=(
            "Product: %{y}<br>"
            "Theme: %{x}<br>"
            "Mention Rate: %{z:.1f}%"
            "<extra></extra>"
        )
    )

    fig_heatmap = style_chart(
        fig_heatmap,
        "Product × Theme Importance",
        show_legend=False
    )

    fig_heatmap.update_xaxes(
        tickangle=-25
    )

    st.plotly_chart(
        fig_heatmap,
        use_container_width=True
    )


# ============================================================
# THEME SENTIMENT
# ============================================================

st.markdown(
    '<div class="section-title">Theme Sentiment</div>',
    unsafe_allow_html=True
)


positive_theme_data = []
negative_theme_data = []


for theme in THEMES:

    column = THEME_COLUMNS[theme]

    if column not in filtered_df.columns:
        continue

    theme_reviews = filtered_df[
        filtered_df[column] == "Yes"
    ]

    theme_total = len(theme_reviews)

    if theme_total == 0:

        positive_percentage = 0
        negative_percentage = 0

    else:

        positive_percentage = (
            theme_reviews["Sentiment"]
            .eq("positive")
            .mean()
            * 100
        )

        negative_percentage = (
            theme_reviews["Sentiment"]
            .eq("negative")
            .mean()
            * 100
        )

    positive_theme_data.append(
        {
            "Theme": theme,
            "Positive %": positive_percentage
        }
    )

    negative_theme_data.append(
        {
            "Theme": theme,
            "Negative %": negative_percentage
        }
    )


positive_theme_df = pd.DataFrame(
    positive_theme_data
)

negative_theme_df = pd.DataFrame(
    negative_theme_data
)


col1, col2 = st.columns(2)


# ------------------------------------------------------------
# Positive Theme Sentiment
# ------------------------------------------------------------

with col1:

    positive_heatmap = (
        positive_theme_df
        .set_index("Theme")
        [["Positive %"]]
    )

    fig_positive_heatmap = px.imshow(
        positive_heatmap,
        text_auto=".1f",
        aspect="auto",
        labels={
            "x": "",
            "y": "Theme",
            "color": "%"
        }
    )

    fig_positive_heatmap.update_traces(
        hovertemplate=(
            "Theme: %{y}<br>"
            "Positive Sentiment: %{z:.1f}%"
            "<extra></extra>"
        )
    )

    fig_positive_heatmap = style_chart(
        fig_positive_heatmap,
        "Positive Sentiment by Theme",
        show_legend=False
    )

    st.plotly_chart(
        fig_positive_heatmap,
        use_container_width=True
    )


# ------------------------------------------------------------
# Negative Theme Sentiment
# ------------------------------------------------------------

with col2:

    negative_heatmap = (
        negative_theme_df
        .set_index("Theme")
        [["Negative %"]]
    )

    fig_negative_heatmap = px.imshow(
        negative_heatmap,
        text_auto=".1f",
        aspect="auto",
        labels={
            "x": "",
            "y": "Theme",
            "color": "%"
        }
    )

    fig_negative_heatmap.update_traces(
        hovertemplate=(
            "Theme: %{y}<br>"
            "Negative Sentiment: %{z:.1f}%"
            "<extra></extra>"
        )
    )

    fig_negative_heatmap = style_chart(
        fig_negative_heatmap,
        "Negative Sentiment by Theme",
        show_legend=False
    )

    st.plotly_chart(
        fig_negative_heatmap,
        use_container_width=True
    )


# ============================================================
# RATING DISTRIBUTION
# ============================================================

st.markdown(
    '<div class="section-title">Rating Distribution</div>',
    unsafe_allow_html=True
)


rating_data = (
    filtered_df
    .groupby(
        ["Product Name", "Stars"]
    )
    .size()
    .reset_index(name="Reviews")
)


rating_totals = (
    filtered_df
    .groupby("Product Name")
    .size()
    .reset_index(name="Total")
)


rating_data = rating_data.merge(
    rating_totals,
    on="Product Name",
    how="left"
)


rating_data["Percentage"] = (
    rating_data["Reviews"]
    / rating_data["Total"]
    * 100
)


rating_data["Stars"] = rating_data["Stars"].astype(int)


fig_rating_distribution = px.bar(
    rating_data,
    x="Stars",
    y="Percentage",
    color="Product Name",
    barmode="group",
    text=None,
    labels={
        "Stars": "Star Rating",
        "Percentage": "Percentage of Reviews (%)",
        "Product Name": ""
    }
)


fig_rating_distribution.update_xaxes(
    tickmode="linear",
    dtick=1,
    range=[0.5, 5.5]
)


fig_rating_distribution.update_yaxes(
    ticksuffix="%"
)


fig_rating_distribution.update_traces(
    hovertemplate=(
        "<b>%{fullData.name}</b><br>"
        "Rating: %{x} stars<br>"
        "Share: %{y:.1f}%<extra></extra>"
    )
)


fig_rating_distribution = style_chart(
    fig_rating_distribution,
    "Customer Rating Distribution",
    height=460
)


fig_rating_distribution = format_axes(
    fig_rating_distribution
)


st.plotly_chart(
    fig_rating_distribution,
    use_container_width=True
)


# ============================================================
# REVIEW TREND
# ============================================================

st.markdown(
    '<div class="section-title">Review Trend</div>',
    unsafe_allow_html=True
)


monthly_reviews = (
    filtered_df
    .dropna(subset=["Posted On"])
    .assign(
        Month=lambda x:
        x["Posted On"].dt.to_period("M").dt.to_timestamp()
    )
    .groupby(
        ["Month", "Product Name"]
    )
    .size()
    .reset_index(name="Reviews")
)


fig_monthly = px.line(
    monthly_reviews,
    x="Month",
    y="Reviews",
    color="Product Name",
    markers=True,
    labels={
        "Month": "Month",
        "Reviews": "Number of Reviews",
        "Product Name": ""
    }
)


fig_monthly.update_traces(
    line=dict(width=2),
    marker=dict(size=5),
    hovertemplate=(
        "<b>%{fullData.name}</b><br>"
        "%{x|%b %Y}<br>"
        "Reviews: %{y:,}<extra></extra>"
    )
)


fig_monthly.update_xaxes(
    dtick="M3",
    tickformat="%b %Y"
)


fig_monthly = style_chart(
    fig_monthly,
    "Monthly Review Volume",
    height=450
)


fig_monthly = format_axes(
    fig_monthly
)


st.plotly_chart(
    fig_monthly,
    use_container_width=True
)


# ============================================================
# KEY CUSTOMER FEEDBACK
# ============================================================

st.markdown(
    '<div class="section-title">Key Customer Feedback</div>',
    unsafe_allow_html=True
)


feedback_rows = []


for theme in THEMES:

    column = THEME_COLUMNS[theme]

    if column not in filtered_df.columns:
        continue

    theme_reviews = filtered_df[
        filtered_df[column] == "Yes"
    ]

    theme_count = len(theme_reviews)

    if theme_count == 0:

        positive_percentage = 0
        negative_percentage = 0

    else:

        positive_percentage = (
            theme_reviews["Sentiment"]
            .eq("positive")
            .mean()
            * 100
        )

        negative_percentage = (
            theme_reviews["Sentiment"]
            .eq("negative")
            .mean()
            * 100
        )

    feedback_rows.append(
        {
            "Theme": theme,
            "Reviews": theme_count,
            "Positive %": positive_percentage,
            "Negative %": negative_percentage
        }
    )


feedback_df = pd.DataFrame(
    feedback_rows
)


col1, col2 = st.columns(2)


# ------------------------------------------------------------
# Customer Strengths
# ------------------------------------------------------------

with col1:

    st.markdown(
        "### Customer Strengths"
    )

    strengths = (
        feedback_df[
            feedback_df["Reviews"] > 0
        ]
        .sort_values(
            ["Positive %", "Reviews"],
            ascending=[False, False]
        )
        .head(5)
        .copy()
    )

    strengths["Positive %"] = (
        strengths["Positive %"]
        .round(1)
        .astype(str)
        + "%"
    )

    st.dataframe(
        strengths[
            [
                "Theme",
                "Reviews",
                "Positive %"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )


# ------------------------------------------------------------
# Customer Pain Points
# ------------------------------------------------------------

with col2:

    st.markdown(
        "### Customer Pain Points"
    )

    pain_points = (
        feedback_df[
            feedback_df["Reviews"] > 0
        ]
        .sort_values(
            ["Negative %", "Reviews"],
            ascending=[False, False]
        )
        .head(5)
        .copy()
    )

    pain_points["Negative %"] = (
        pain_points["Negative %"]
        .round(1)
        .astype(str)
        + "%"
    )

    st.dataframe(
        pain_points[
            [
                "Theme",
                "Reviews",
                "Negative %"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# PRODUCT SUMMARY
# ============================================================

st.markdown(
    '<div class="section-title">Product Summary</div>',
    unsafe_allow_html=True
)


summary_rows = []


for product in PRODUCTS:

    product_df = filtered_df[
        filtered_df["Product Name"] == product
    ]

    if product_df.empty:
        continue

    product_reviews = len(product_df)

    product_positive = (
        product_df["Sentiment"]
        .eq("positive")
        .mean()
        * 100
    )

    product_neutral = (
        product_df["Sentiment"]
        .eq("neutral")
        .mean()
        * 100
    )

    product_negative = (
        product_df["Sentiment"]
        .eq("negative")
        .mean()
        * 100
    )

    summary_rows.append(
        {
            "Product": product,
            "Reviews": product_reviews,
            "Average Rating": product_df["Stars"].mean(),
            "Positive %": product_positive,
            "Neutral %": product_neutral,
            "Negative %": product_negative
        }
    )


product_summary_table = pd.DataFrame(
    summary_rows
)


if not product_summary_table.empty:

    product_summary_table = (
        product_summary_table
        .sort_values(
            "Average Rating",
            ascending=False
        )
        .reset_index(drop=True)
    )

    display_summary = product_summary_table.copy()

    display_summary["Average Rating"] = (
        display_summary["Average Rating"]
        .round(2)
    )

    display_summary["Positive %"] = (
        display_summary["Positive %"]
        .round(1)
        .astype(str)
        + "%"
    )

    display_summary["Neutral %"] = (
        display_summary["Neutral %"]
        .round(1)
        .astype(str)
        + "%"
    )

    display_summary["Negative %"] = (
        display_summary["Negative %"]
        .round(1)
        .astype(str)
        + "%"
    )

    st.dataframe(
        display_summary,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div style="
        text-align:center;
        color:#94A3B8;
        font-size:12px;
        margin-top:35px;
        padding-top:15px;
        border-top:1px solid #E2E8F0;
    ">
        Germany Toothbrush Customer Review Dashboard
    </div>
    """,
    unsafe_allow_html=True
)
