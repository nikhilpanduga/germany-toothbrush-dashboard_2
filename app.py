from pathlib import Path
import warnings

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

warnings.filterwarnings("ignore")


# =========================================================
# PAGE CONFIGURATION
# =========================================================
st.set_page_config(
    page_title="Germany Toothbrush Customer Review Dashboard",
    page_icon="🪥",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# FILE CONFIGURATION
# =========================================================
DATA_FILE = Path(__file__).parent / "Germany_five toothbrush data.xlsx"
SHEET_NAME = "Germany"


# =========================================================
# PRODUCT ORDER
# =========================================================
PRODUCT_ORDER = [
    "curaprox 5460",
    "Dr. Best Clean Pro Zwischenzahn",
    "elmex expert precision",
    "elmex InterX",
    "meridol base",
]


# =========================================================
# CUSTOMER THEMES
# =========================================================
THEMES = [
    "Cleaning Performance",
    "Brushing Comfort",
    "Design & Ease of Use",
    "Quality & Durability",
    "Features & Technology",
    "Price & Value",
]


THEME_COLS = {
    theme: f"{theme}_Present"
    for theme in THEMES
}


# =========================================================
# COLORS
# =========================================================
COLORS = {
    "positive": "#2E7D32",
    "neutral": "#7A858D",
    "negative": "#B84A4A",
}


PRODUCT_COLORS = [
    "#2F5D8C",
    "#4F81BD",
    "#6793B5",
    "#7FA3BD",
    "#9AB7CA",
]


# =========================================================
# GLOBAL STYLING
# =========================================================
st.markdown(
    """
<style>

.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
    max-width: 1500px;
}

/* ==============================
   Dashboard Header
   ============================== */

.dashboard-header {
    background: linear-gradient(135deg, #173B5E, #2F5D8C);
    padding: 24px 30px;
    border-radius: 14px;
    margin-bottom: 22px;
    color: white;
}

.dashboard-header h1 {
    color: white;
    margin: 0 0 6px 0;
    padding: 0;
    font-size: 30px;
    font-weight: 700;
    line-height: 1.2;
}

.dashboard-header p {
    color: #E2E8F0;
    margin: 0;
    padding: 0;
    font-size: 15px;
    line-height: 1.5;
}


/* ==============================
   Section Titles
   ============================== */

.section-title {
    color: #173B5E;
    font-size: 21px;
    font-weight: 700;
    margin: 24px 0 10px 0;
    padding-bottom: 7px;
    border-bottom: 2px solid #D8E1E8;
}


/* ==============================
   KPI Cards
   ============================== */

[data-testid="stMetric"] {
    background: #FFFFFF;
    border: 1px solid #C8D4DE;
    border-radius: 12px;
    padding: 14px 16px;
}

[data-testid="stMetricLabel"] {
    color: #637586 !important;
    font-size: 13px !important;
}

[data-testid="stMetricValue"] {
    color: #173B5E !important;
    font-size: 28px !important;
    font-weight: 700 !important;
}


/* ==============================
   Sidebar
   ============================== */

div[data-testid="stSidebar"] {
    background-color: #F5F8FA;
}

.sidebar-note {
    color: #637586;
    font-size: 12px;
    line-height: 1.5;
}


/* ==============================
   Tables
   ============================== */

[data-testid="stDataFrame"] {
    border-radius: 10px;
}


/* ==============================
   Buttons
   ============================== */

.stButton > button {
    border-radius: 8px;
    font-weight: 600;
}


/* ==============================
   Selectboxes / Multiselect
   ============================== */

div[data-baseweb="select"] {
    border-radius: 8px;
}


/* ==============================
   Remove unnecessary top space
   ============================== */

div[data-testid="stVerticalBlock"] > div {
    gap: 0.5rem;
}

</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# DATA LOADING AND CLEANING
# =========================================================
@st.cache_data
def load_data():

    if not DATA_FILE.exists():
        raise FileNotFoundError(
            f"'{DATA_FILE.name}' was not found. "
            "Keep the Excel file in the same folder as app.py."
        )

    data = pd.read_excel(
        DATA_FILE,
        sheet_name=SHEET_NAME,
    )

    data.columns = [
        str(column).strip()
        for column in data.columns
    ]

    # -----------------------------------------------------
    # Required columns
    # -----------------------------------------------------
    required_columns = [
        "Product Name given",
        "Brand",
        "Site",
        "Posted On",
        "Stars",
        "Sentiment",
        *THEME_COLS.values(),
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in data.columns
    ]

    if missing_columns:
        raise ValueError(
            "Missing required columns:\n"
            + "\n".join(
                f"- {column}"
                for column in missing_columns
            )
        )

    # -----------------------------------------------------
    # Product standardization
    # -----------------------------------------------------
    product_map = {
        "curaprox 5460": "curaprox 5460",
        "Dr. Best Clean Pro Zwischenzahn":
            "Dr. Best Clean Pro Zwischenzahn",
        "elmex expert precision":
            "elmex expert precision",
        "elmex InterX":
            "elmex InterX",
        "meridol base":
            "meridol base",
    }

    data["Product"] = (
        data["Product Name given"]
        .astype(str)
        .str.strip()
        .replace(product_map)
    )

    # -----------------------------------------------------
    # Brand and Site
    # -----------------------------------------------------
    data["Brand"] = (
        data["Brand"]
        .astype(str)
        .str.strip()
    )

    data["Site"] = (
        data["Site"]
        .astype(str)
        .str.strip()
    )

    # -----------------------------------------------------
    # Sentiment standardization
    # -----------------------------------------------------
    sentiment_map = {
        "positive": "positive",
        "pos": "positive",
        "neutral": "neutral",
        "neu": "neutral",
        "negative": "negative",
        "neg": "negative",
    }

    data["Sentiment"] = (
        data["Sentiment"]
        .astype(str)
        .str.strip()
        .str.lower()
        .map(sentiment_map)
    )

    # -----------------------------------------------------
    # Theme standardization
    # -----------------------------------------------------
    yes_values = {
        "yes",
        "y",
        "true",
        "1",
    }

    no_values = {
        "no",
        "n",
        "false",
        "0",
    }

    def normalize_yes_no(value):

        if pd.isna(value):
            return np.nan

        value = (
            str(value)
            .strip()
            .lower()
        )

        if value in yes_values:
            return "Yes"

        if value in no_values:
            return "No"

        return np.nan

    for column in THEME_COLS.values():

        data[column] = data[column].apply(
            normalize_yes_no
        )

    # -----------------------------------------------------
    # Date
    # -----------------------------------------------------
    data["Posted On"] = pd.to_datetime(
        data["Posted On"],
        errors="coerce",
    )

    # -----------------------------------------------------
    # Stars
    # -----------------------------------------------------
    data["Stars"] = pd.to_numeric(
        data["Stars"],
        errors="coerce",
    )

    # -----------------------------------------------------
    # Keep only expected products
    # -----------------------------------------------------
    data = data[
        data["Product"].isin(PRODUCT_ORDER)
    ].copy()

    # -----------------------------------------------------
    # Product ordering
    # -----------------------------------------------------
    data["Product"] = pd.Categorical(
        data["Product"],
        categories=PRODUCT_ORDER,
        ordered=True,
    )

    # -----------------------------------------------------
    # Remove exact duplicate rows only
    # -----------------------------------------------------
    data = (
        data
        .drop_duplicates()
        .reset_index(drop=True)
    )

    return data


# =========================================================
# LOAD DATA
# =========================================================
try:

    df = load_data()

except Exception as exc:

    st.error(
        f"Unable to load the dashboard data.\n\n{exc}"
    )

    st.stop()


# =========================================================
# HELPER FUNCTIONS
# =========================================================
def pct(part, total):

    if total == 0:
        return 0

    return (part / total) * 100


def apply_filters(
    data,
    products,
    brands,
    sites,
    sentiments,
    star_range,
    date_range,
):

    filtered = data.copy()

    # Product
    if (
        products
        and "All Products" not in products
    ):

        filtered = filtered[
            filtered["Product"].isin(products)
        ]

    # Brand
    if brands:

        filtered = filtered[
            filtered["Brand"].isin(brands)
        ]

    # Site
    if sites:

        filtered = filtered[
            filtered["Site"].isin(sites)
        ]

    # Sentiment
    if sentiments:

        sentiment_values = [
            sentiment.lower()
            for sentiment in sentiments
        ]

        filtered = filtered[
            filtered["Sentiment"].isin(
                sentiment_values
            )
        ]

    # Star Rating
    if star_range:

        filtered = filtered[
            filtered["Stars"].between(
                star_range[0],
                star_range[1],
                inclusive="both",
            )
        ]

    # Review Date
    if (
        date_range
        and len(date_range) == 2
    ):

        start_date, end_date = date_range

        filtered = filtered[
            filtered["Posted On"].between(
                pd.Timestamp(start_date),
                (
                    pd.Timestamp(end_date)
                    + pd.Timedelta(days=1)
                    - pd.Timedelta(seconds=1)
                ),
            )
        ]

    return filtered


def base_layout(
    fig,
    height=390,
):

    fig.update_layout(

        template="plotly_white",

        height=height,

        margin=dict(
            l=45,
            r=25,
            t=60,
            b=50,
        ),

        font=dict(
            family="Arial",
            size=12,
            color="#34495A",
        ),

        title=dict(
            x=0.02,
            xanchor="left",
            font=dict(
                size=16,
                color="#173B5E",
            ),
        ),

        paper_bgcolor="white",

        plot_bgcolor="white",

        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="left",
            x=0,
        ),
    )

    fig.update_xaxes(
        showgrid=True,
        gridcolor="#D6E0E7",
        zeroline=False,
    )

    fig.update_yaxes(
        showgrid=False,
        zeroline=False,
    )

    return fig


# =========================================================
# DASHBOARD HEADER
# =========================================================
# IMPORTANT:
# HTML starts directly after the triple quote so Streamlit
# does not interpret it as a code block.
# =========================================================
st.markdown(
    """<div class="dashboard-header">
<h1>Germany Toothbrush Customer Review Dashboard</h1>
<p>Customer perception, ratings, sentiment and key product themes</p>
</div>""",
    unsafe_allow_html=True,
)


# =========================================================
# SIDEBAR FILTERS
# =========================================================
with st.sidebar:

    st.markdown("## Dashboard Filters")

    # -----------------------------------------------------
    # Product
    # -----------------------------------------------------
    product_options = [
        "All Products"
    ] + PRODUCT_ORDER

    selected_products = st.multiselect(
        "Product",
        product_options,
        default=["All Products"],
    )

    # -----------------------------------------------------
    # Brand
    # -----------------------------------------------------
    brand_options = sorted(
        df["Brand"]
        .dropna()
        .unique()
        .tolist()
    )

    selected_brands = st.multiselect(
        "Brand",
        brand_options,
    )

    # -----------------------------------------------------
    # Site
    # -----------------------------------------------------
    site_options = sorted(
        df["Site"]
        .dropna()
        .unique()
        .tolist()
    )

    selected_sites = st.multiselect(
        "Site",
        site_options,
    )

    # -----------------------------------------------------
    # Sentiment
    # -----------------------------------------------------
    selected_sentiments = st.multiselect(
        "Sentiment",
        [
            "Positive",
            "Neutral",
            "Negative",
        ],
    )

    # -----------------------------------------------------
    # Star Rating
    # -----------------------------------------------------
    selected_stars = st.slider(
        "Star Rating",
        min_value=1,
        max_value=5,
        value=(1, 5),
        step=1,
    )

    # -----------------------------------------------------
    # Review Date
    # -----------------------------------------------------
    valid_dates = (
        df["Posted On"]
        .dropna()
    )

    min_date = valid_dates.min().date()
    max_date = valid_dates.max().date()

    selected_dates = st.date_input(
        "Review Date",
        value=(
            min_date,
            max_date,
        ),
        min_value=min_date,
        max_value=max_date,
    )

    st.markdown("---")

    # -----------------------------------------------------
    # Reset Button
    # -----------------------------------------------------
    if st.button(
        "Reset Filters",
        use_container_width=True,
    ):

        st.session_state.clear()
        st.rerun()

    # -----------------------------------------------------
    # Sidebar Information
    # -----------------------------------------------------
    st.markdown(
        """<div class="sidebar-note">
<b>Data scope</b><br>
Germany customer reviews across five toothbrush products.
<br><br>
Missing theme values are kept as missing and are not automatically treated as "No".
</div>""",
        unsafe_allow_html=True,
    )


# =========================================================
# APPLY FILTERS
# =========================================================
filtered_df = apply_filters(
    df,
    selected_products,
    selected_brands,
    selected_sites,
    selected_sentiments,
    selected_stars,
    selected_dates,
)


# =========================================================
# KPI CALCULATIONS
# =========================================================
total_reviews = len(filtered_df)

avg_rating = (
    filtered_df["Stars"].mean()
    if total_reviews
    else np.nan
)

positive_pct = pct(
    (
        filtered_df["Sentiment"]
        == "positive"
    ).sum(),
    total_reviews,
)

negative_pct = pct(
    (
        filtered_df["Sentiment"]
        == "negative"
    ).sum(),
    total_reviews,
)


if (
    selected_products
    and "All Products"
    not in selected_products
):

    product_count = len(
        selected_products
    )

else:

    product_count = (
        filtered_df["Product"]
        .nunique()
    )


# =========================================================
# KPI CARDS
# =========================================================
k1, k2, k3, k4, k5 = st.columns(5)


k1.metric(
    "Total Reviews",
    f"{total_reviews:,}",
)


k2.metric(
    "Average Rating",
    (
        f"{avg_rating:.2f}"
        if not np.isnan(avg_rating)
        else "—"
    ),
)


k3.metric(
    "Positive Sentiment",
    f"{positive_pct:.1f}%",
)


k4.metric(
    "Negative Sentiment",
    f"{negative_pct:.1f}%",
)


k5.metric(
    "Products",
    f"{product_count}",
)


# =========================================================
# PRODUCT PERFORMANCE
# =========================================================
st.markdown(
    '<div class="section-title">Product Performance</div>',
    unsafe_allow_html=True,
)

left, right = st.columns(2)


# ---------------------------------------------------------
# Reviews by Product
# ---------------------------------------------------------
with left:

    review_counts = (
        filtered_df
        .groupby(
            "Product",
            observed=True,
        )
        .size()
        .reindex(PRODUCT_ORDER)
        .fillna(0)
        .astype(int)
        .reset_index(
            name="Reviews"
        )
    )

    fig = px.bar(
        review_counts,
        x="Reviews",
        y="Product",
        orientation="h",
        text="Reviews",
        title="Reviews by Product",
        labels={
            "Reviews": "Number of Reviews",
            "Product": "",
        },
    )

    fig.update_traces(
        marker_color="#4F81BD",
        textposition="outside",
    )

    fig.update_layout(
        yaxis=dict(
            categoryorder="array",
            categoryarray=PRODUCT_ORDER,
        )
    )

    st.plotly_chart(
        base_layout(fig),
        use_container_width=True,
        config={
            "displayModeBar": False
        },
    )


# ---------------------------------------------------------
# Average Rating by Product
# ---------------------------------------------------------
with right:

    rating_summary = (
        filtered_df
        .groupby(
            "Product",
            observed=True,
        )["Stars"]
        .mean()
        .reindex(PRODUCT_ORDER)
        .dropna()
        .reset_index(
            name="Average Rating"
        )
    )

    fig = px.bar(
        rating_summary,
        x="Average Rating",
        y="Product",
        orientation="h",
        text=rating_summary[
            "Average Rating"
        ].round(2),
        title="Average Rating by Product",
        labels={
            "Average Rating": "Average Rating",
            "Product": "",
        },
    )

    fig.update_traces(
        marker_color="#2F5D8C",
        textposition="outside",
    )

    fig.update_xaxes(
        range=[0, 5]
    )

    fig.update_layout(
        yaxis=dict(
            categoryorder="array",
            categoryarray=PRODUCT_ORDER,
        )
    )

    st.plotly_chart(
        base_layout(fig),
        use_container_width=True,
        config={
            "displayModeBar": False
        },
    )


# =========================================================
# CUSTOMER SENTIMENT
# =========================================================
st.markdown(
    '<div class="section-title">Customer Sentiment</div>',
    unsafe_allow_html=True,
)

left, right = st.columns(2)


# ---------------------------------------------------------
# Sentiment by Product
# ---------------------------------------------------------
with left:

    sentiment_rows = []

    for product in PRODUCT_ORDER:

        temp = filtered_df[
            filtered_df["Product"]
            == product
        ]

        product_total = len(temp)

        if product_total == 0:
            continue

        for sentiment in [
            "positive",
            "neutral",
            "negative",
        ]:

            sentiment_rows.append(
                {
                    "Product": product,
                    "Sentiment": sentiment.title(),
                    "Percentage": pct(
                        (
                            temp["Sentiment"]
                            == sentiment
                        ).sum(),
                        product_total,
                    ),
                }
            )

    sentiment_product = pd.DataFrame(
        sentiment_rows
    )

    if sentiment_product.empty:

        st.info(
            "No data available for the selected filters."
        )

    else:

        fig = px.bar(
            sentiment_product,
            x="Percentage",
            y="Product",
            color="Sentiment",
            orientation="h",
            text=(
                sentiment_product[
                    "Percentage"
                ]
                .round(0)
                .astype(int)
                .astype(str)
                + "%"
            ),
            title="Sentiment by Product",
            labels={
                "Percentage": "Share of Reviews (%)",
                "Product": "",
            },
            color_discrete_map={
                "Positive": COLORS["positive"],
                "Neutral": COLORS["neutral"],
                "Negative": COLORS["negative"],
            },
        )

        fig.update_traces(
            textposition="inside"
        )

        fig.update_layout(
            barmode="stack",
            xaxis=dict(
                range=[0, 100]
            ),
            yaxis=dict(
                categoryorder="array",
                categoryarray=PRODUCT_ORDER,
            ),
        )

        st.plotly_chart(
            base_layout(fig),
            use_container_width=True,
            config={
                "displayModeBar": False
            },
        )


# ---------------------------------------------------------
# Overall Sentiment
# ---------------------------------------------------------
with right:

    sentiment_total = (
        filtered_df["Sentiment"]
        .value_counts()
        .reindex(
            [
                "positive",
                "neutral",
                "negative",
            ]
        )
        .fillna(0)
        .reset_index()
    )

    sentiment_total.columns = [
        "Sentiment",
        "Reviews",
    ]

    sentiment_total["Sentiment"] = (
        sentiment_total["Sentiment"]
        .str.title()
    )

    fig = px.pie(
        sentiment_total,
        names="Sentiment",
        values="Reviews",
        hole=0.58,
        title="Overall Sentiment",
        color="Sentiment",
        color_discrete_map={
            "Positive": COLORS["positive"],
            "Neutral": COLORS["neutral"],
            "Negative": COLORS["negative"],
        },
    )

    fig.update_traces(
        textinfo="percent"
    )

    st.plotly_chart(
        base_layout(fig),
        use_container_width=True,
        config={
            "displayModeBar": False
        },
    )


# =========================================================
# CUSTOMER THEMES
# =========================================================
st.markdown(
    '<div class="section-title">What Customers Talk About</div>',
    unsafe_allow_html=True,
)

theme_rows = []

for theme, column in THEME_COLS.items():

    theme_count = (
        filtered_df[column]
        == "Yes"
    ).sum()

    theme_rows.append(
        {
            "Theme": theme,
            "Percentage": pct(
                theme_count,
                total_reviews,
            ),
        }
    )


theme_summary = (
    pd.DataFrame(theme_rows)
    .sort_values(
        "Percentage",
        ascending=True,
    )
)


fig = px.bar(
    theme_summary,
    x="Percentage",
    y="Theme",
    orientation="h",
    text=(
        theme_summary["Percentage"]
        .round(1)
        .astype(str)
        + "%"
    ),
    title="Customer Theme Importance",
    labels={
        "Percentage": "Reviews Mentioning Theme (%)",
        "Theme": "",
    },
)

fig.update_traces(
    marker_color="#4F81BD",
    textposition="outside",
)

st.plotly_chart(
    base_layout(fig, 410),
    use_container_width=True,
    config={
        "displayModeBar": False
    },
)


# =========================================================
# PRODUCT VS CUSTOMER THEMES
# =========================================================
theme_matrix = pd.DataFrame(
    index=PRODUCT_ORDER,
    columns=THEMES,
    dtype=float,
)


for product in PRODUCT_ORDER:

    product_data = filtered_df[
        filtered_df["Product"]
        == product
    ]

    for theme, column in THEME_COLS.items():

        theme_matrix.loc[
            product,
            theme,
        ] = (
            pct(
                (
                    product_data[column]
                    == "Yes"
                ).sum(),
                len(product_data),
            )
            if len(product_data)
            else np.nan
        )


fig = px.imshow(
    theme_matrix,
    text_auto=".1f",
    aspect="auto",
    color_continuous_scale=[
        "#F8FAFC",
        "#B4C7DC",
        "#2F5D8C",
    ],
    title="Product vs Customer Themes",
    labels={
        "color": "Theme Mentions (%)"
    },
)

fig.update_traces(
    hovertemplate=(
        "Product: %{y}"
        "<br>Theme: %{x}"
        "<br>Mention: %{z:.1f}%"
        "<extra></extra>"
    )
)

st.plotly_chart(
    base_layout(fig, 410),
    use_container_width=True,
    config={
        "displayModeBar": False
    },
)


# =========================================================
# CUSTOMER SENTIMENT BY THEME
# =========================================================
st.markdown(
    '<div class="section-title">Customer Sentiment by Theme</div>',
    unsafe_allow_html=True,
)

left, right = st.columns(2)


for container, sentiment, title in [

    (
        left,
        "positive",
        "Positive Sentiment by Theme",
    ),

    (
        right,
        "negative",
        "Negative Sentiment by Theme",
    ),

]:

    matrix = pd.DataFrame(
        index=PRODUCT_ORDER,
        columns=THEMES,
        dtype=float,
    )

    for product in PRODUCT_ORDER:

        product_data = filtered_df[
            filtered_df["Product"]
            == product
        ]

        for theme, column in THEME_COLS.items():

            theme_data = product_data[
                product_data[column]
                == "Yes"
            ]

            theme_total = len(
                theme_data
            )

            matrix.loc[
                product,
                theme,
            ] = (
                pct(
                    (
                        theme_data["Sentiment"]
                        == sentiment
                    ).sum(),
                    theme_total,
                )
                if theme_total
                else np.nan
            )

    fig = px.imshow(
        matrix,
        text_auto=".1f",
        aspect="auto",
        color_continuous_scale=[
            "#F8FAFC",
            "#B4C7DC",
            "#2F5D8C",
        ],
        title=title,
        labels={
            "color": "Sentiment (%)"
        },
    )

    fig.update_traces(
        hovertemplate=(
            "Product: %{y}"
            "<br>Theme: %{x}"
            "<br>Sentiment: %{z:.1f}%"
            "<extra></extra>"
        )
    )

    with container:

        st.plotly_chart(
            base_layout(fig, 410),
            use_container_width=True,
            config={
                "displayModeBar": False
            },
        )


# =========================================================
# CUSTOMER RATING DISTRIBUTION
# =========================================================
st.markdown(
    '<div class="section-title">Customer Rating Distribution</div>',
    unsafe_allow_html=True,
)

rating_rows = []


for product in PRODUCT_ORDER:

    temp = filtered_df[
        filtered_df["Product"]
        == product
    ]

    product_total = len(temp)

    if product_total == 0:
        continue

    for star in range(1, 6):

        rating_rows.append(
            {
                "Product": product,
                "Star": str(star),
                "Percentage": pct(
                    (
                        temp["Stars"]
                        == star
                    ).sum(),
                    product_total,
                ),
            }
        )


rating_distribution = pd.DataFrame(
    rating_rows
)


if rating_distribution.empty:

    st.info(
        "No rating data available for the selected filters."
    )

else:

    fig = px.bar(
        rating_distribution,
        x="Star",
        y="Percentage",
        color="Product",
        barmode="group",
        text=(
            rating_distribution[
                "Percentage"
            ]
            .round(0)
            .astype(int)
            .astype(str)
            + "%"
        ),
        title="Customer Rating Distribution",
        labels={
            "Percentage": "Reviews (%)",
            "Star": "Star Rating",
        },
        category_orders={
            "Star": [
                "1",
                "2",
                "3",
                "4",
                "5",
            ]
        },
        color_discrete_sequence=PRODUCT_COLORS,
    )

    fig.update_traces(
        textposition="outside",
        textfont_size=9,
    )

    st.plotly_chart(
        base_layout(fig, 400),
        use_container_width=True,
        config={
            "displayModeBar": False
        },
    )


# =========================================================
# REVIEW ACTIVITY OVER TIME
# =========================================================
st.markdown(
    '<div class="section-title">Review Activity Over Time</div>',
    unsafe_allow_html=True,
)

trend_data = (
    filtered_df
    .dropna(
        subset=["Posted On"]
    )
    .copy()
)


if trend_data.empty:

    st.info(
        "No date information available for the selected filters."
    )

else:

    monthly = (
        trend_data.assign(
            Month=(
                trend_data["Posted On"]
                .dt.to_period("M")
                .dt.to_timestamp()
            )
        )
        .groupby(
            [
                "Month",
                "Product",
            ],
            observed=True,
        )
        .size()
        .reset_index(
            name="Reviews"
        )
    )

    fig = px.line(
        monthly,
        x="Month",
        y="Reviews",
        color="Product",
        markers=True,
        title="Monthly Review Volume",
        labels={
            "Reviews": "Number of Reviews",
            "Month": "Month",
        },
        color_discrete_sequence=PRODUCT_COLORS,
    )

    fig.update_layout(
        hovermode="x unified"
    )

    st.plotly_chart(
        base_layout(fig, 400),
        use_container_width=True,
        config={
            "displayModeBar": False
        },
    )


# =========================================================
# KEY CUSTOMER FEEDBACK
# =========================================================
st.markdown(
    '<div class="section-title">Key Customer Feedback</div>',
    unsafe_allow_html=True,
)

strength_rows = []
pain_rows = []


for theme, column in THEME_COLS.items():

    theme_data = filtered_df[
        filtered_df[column]
        == "Yes"
    ]

    theme_total = len(
        theme_data
    )

    if theme_total == 0:
        continue

    positive = (
        theme_data["Sentiment"]
        == "positive"
    ).sum()

    negative = (
        theme_data["Sentiment"]
        == "negative"
    ).sum()

    strength_rows.append(
        {
            "Theme": theme,
            "Reviews": theme_total,
            "Positive %": round(
                pct(
                    positive,
                    theme_total,
                ),
                1,
            ),
        }
    )

    pain_rows.append(
        {
            "Theme": theme,
            "Reviews": theme_total,
            "Negative %": round(
                pct(
                    negative,
                    theme_total,
                ),
                1,
            ),
        }
    )


strengths = pd.DataFrame(
    strength_rows
)


if not strengths.empty:

    strengths = (
        strengths
        .sort_values(
            [
                "Positive %",
                "Reviews",
            ],
            ascending=[
                False,
                False,
            ],
        )
        .head(6)
    )


pain_points = pd.DataFrame(
    pain_rows
)


if not pain_points.empty:

    pain_points = (
        pain_points
        .sort_values(
            [
                "Negative %",
                "Reviews",
            ],
            ascending=[
                False,
                False,
            ],
        )
        .head(6)
    )


left, right = st.columns(2)


# ---------------------------------------------------------
# Customer Strengths
# ---------------------------------------------------------
with left:

    st.markdown(
        "### Customer Strengths"
    )

    st.dataframe(
        strengths,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Positive %":
                st.column_config.ProgressColumn(
                    "Positive %",
                    min_value=0,
                    max_value=100,
                    format="%.1f%%",
                )
        },
    )


# ---------------------------------------------------------
# Customer Pain Points
# ---------------------------------------------------------
with right:

    st.markdown(
        "### Customer Pain Points"
    )

    st.dataframe(
        pain_points,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Negative %":
                st.column_config.ProgressColumn(
                    "Negative %",
                    min_value=0,
                    max_value=100,
                    format="%.1f%%",
                )
        },
    )


# =========================================================
# PRODUCT SUMMARY
# =========================================================
st.markdown(
    '<div class="section-title">Product Summary</div>',
    unsafe_allow_html=True,
)

summary_rows = []


for product in PRODUCT_ORDER:

    temp = filtered_df[
        filtered_df["Product"]
        == product
    ]

    product_total = len(temp)

    if product_total == 0:

        summary_rows.append(
            {
                "Product": product,
                "Reviews": 0,
                "Average Rating": np.nan,
                "Positive %": np.nan,
                "Neutral %": np.nan,
                "Negative %": np.nan,
            }
        )

        continue

    summary_rows.append(
        {
            "Product": product,
            "Reviews": product_total,

            "Average Rating": round(
                temp["Stars"].mean(),
                2,
            ),

            "Positive %": round(
                pct(
                    (
                        temp["Sentiment"]
                        == "positive"
                    ).sum(),
                    product_total,
                ),
                1,
            ),

            "Neutral %": round(
                pct(
                    (
                        temp["Sentiment"]
                        == "neutral"
                    ).sum(),
                    product_total,
                ),
                1,
            ),

            "Negative %": round(
                pct(
                    (
                        temp["Sentiment"]
                        == "negative"
                    ).sum(),
                    product_total,
                ),
                1,
            ),
        }
    )


product_summary = pd.DataFrame(
    summary_rows
).sort_values(
    "Average Rating",
    ascending=False,
    na_position="last",
)


st.dataframe(
    product_summary,
    use_container_width=True,
    hide_index=True,
    column_config={

        "Average Rating":
            st.column_config.ProgressColumn(
                "Average Rating",
                min_value=0,
                max_value=5,
                format="%.2f",
            ),

        "Positive %":
            st.column_config.ProgressColumn(
                "Positive %",
                min_value=0,
                max_value=100,
                format="%.1f%%",
            ),

        "Neutral %":
            st.column_config.ProgressColumn(
                "Neutral %",
                min_value=0,
                max_value=100,
                format="%.1f%%",
            ),

        "Negative %":
            st.column_config.ProgressColumn(
                "Negative %",
                min_value=0,
                max_value=100,
                format="%.1f%%",
            ),
    },
)


# =========================================================
# INTERNAL REVIEW COUNT VALIDATION
# =========================================================
# This does NOT display a validation card.
# It only displays an error if the counts do not reconcile.
# =========================================================

summary_count_total = int(
    product_summary["Reviews"].sum()
)

if summary_count_total != total_reviews:

    st.error(
        "Review count reconciliation failed: "
        f"Product Summary = {summary_count_total:,}, "
        f"Total Reviews = {total_reviews:,}."
    )


# =========================================================
# FOOTER
# =========================================================
st.caption(
    f"Showing {total_reviews:,} filtered reviews "
    "from the Germany toothbrush review dataset."
)
