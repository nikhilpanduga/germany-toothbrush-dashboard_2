from pathlib import Path
import warnings

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

warnings.filterwarnings("ignore")


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Germany Toothbrush Customer Review Dashboard",
    page_icon="🪥",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# FILE / DATA CONFIG
# ============================================================

DATA_FILE = (
    Path(__file__).parent
    / "Germany_five toothbrush data.xlsx"
)

SHEET_NAME = "Germany"


PRODUCT_ORDER = [
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


THEME_COLS = {
    theme: f"{theme}_Present"
    for theme in THEMES
}


# ============================================================
# COLOR PALETTE
# ============================================================

# Stronger contrast than the previous version.
# Still professional and not excessively bright.

NAVY = "#173B5E"
DARK_BLUE = "#234F73"
BLUE = "#2F5D8C"
MEDIUM_BLUE = "#4F81BD"
LIGHT_BLUE = "#6F98BD"
PALE_BLUE = "#8EABC2"

POSITIVE = "#3F7D4A"
NEUTRAL = "#7A858D"
NEGATIVE = "#B84A4A"

BACKGROUND = "#EEF2F5"
CARD = "#FFFFFF"
BORDER = "#C8D4DE"
GRID = "#D6E0E7"
TEXT = "#34495A"
MUTED = "#637586"


PRODUCT_COLORS = [
    "#2F5D8C",
    "#4F81BD",
    "#6793B5",
    "#7FA3BD",
    "#9AB7CA",
]


# ============================================================
# CSS
# ============================================================

st.markdown(
    f"""
    <style>

    /* ======================================================
       PAGE
    ====================================================== */

    .stApp {{
        background-color: {BACKGROUND};
    }}

    .block-container {{
        max-width: 1500px;
        padding-top: 1.15rem;
        padding-bottom: 2rem;
    }}


    /* ======================================================
       HEADER
    ====================================================== */

    .dashboard-header {{
        background: linear-gradient(
            135deg,
            {NAVY} 0%,
            {BLUE} 100%
        );

        padding: 23px 28px;

        border-radius: 10px;

        margin-bottom: 20px;

        box-shadow:
            0 3px 7px rgba(23, 59, 94, 0.15);
    }}

    .dashboard-title {{
        color: #FFFFFF;

        font-size: 29px;

        font-weight: 700;

        line-height: 1.2;

        letter-spacing: -0.2px;
    }}

    .dashboard-subtitle {{
        color: #DCE7F0;

        font-size: 14px;

        margin-top: 6px;
    }}


    /* ======================================================
       SECTION HEADINGS
    ====================================================== */

    .section-title {{
        color: {NAVY};

        font-size: 20px;

        font-weight: 700;

        margin-top: 25px;

        margin-bottom: 10px;

        padding-bottom: 7px;

        border-bottom:
            2px solid {BORDER};
    }}


    /* ======================================================
       KPI CARDS
    ====================================================== */

    [data-testid="stMetric"] {{
        background-color: {CARD};

        border:
            1px solid {BORDER};

        border-radius: 9px;

        padding: 13px 16px;

        min-height: 96px;

        box-shadow:
            0 2px 4px
            rgba(35, 79, 115, 0.07);
    }}

    [data-testid="stMetricLabel"] {{
        color: {MUTED} !important;

        font-size: 12px !important;

        font-weight: 600 !important;
    }}

    [data-testid="stMetricValue"] {{
        color: {NAVY} !important;

        font-size: 27px !important;

        font-weight: 700 !important;
    }}


    /* ======================================================
       DATA VALIDATION CARD
       ====================================================== */

    .validation-card {{
        background: #E1E8EE;

        border:
            1px solid #B9C8D5;

        border-radius: 7px;

        padding: 8px 13px;

        margin-top: 10px;

        margin-bottom: 4px;

        color: {NAVY};

        font-size: 12px;
    }}

    .validation-number {{
        font-weight: 700;

        color: {NAVY};
    }}


    /* ======================================================
       SIDEBAR
       ====================================================== */

    section[data-testid="stSidebar"] {{
        background-color: #E1E7EC;

        border-right:
            1px solid #C5D0DA;
    }}

    section[data-testid="stSidebar"] h2 {{
        color: {NAVY};
    }}

    .sidebar-note {{
        color: #536878;

        font-size: 12px;

        line-height: 1.55;
    }}


    /* ======================================================
       BUTTONS
       ====================================================== */

    .stButton > button {{
        background-color: #FFFFFF;

        color: {NAVY};

        border:
            1px solid #AABBC9;

        border-radius: 6px;

        font-weight: 600;
    }}

    .stButton > button:hover {{
        background-color: #E7EEF3;

        color: {NAVY};

        border-color: #6D8AA2;
    }}


    /* ======================================================
       TABLE
       ====================================================== */

    [data-testid="stDataFrame"] {{
        border:
            1px solid {BORDER};

        border-radius: 8px;
    }}


    /* ======================================================
       SUB HEADINGS
       ====================================================== */

    h3 {{
        color: {DARK_BLUE} !important;
    }}

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# LOAD DATA
# ============================================================

@st.cache_data
def load_data():

    if not DATA_FILE.exists():

        raise FileNotFoundError(
            f"Could not find '{DATA_FILE.name}'. "
            "Place the Excel file in the same folder as app.py."
        )


    data = pd.read_excel(
        DATA_FILE,
        sheet_name=SHEET_NAME,
    )


    # --------------------------------------------------------
    # Clean column names
    # --------------------------------------------------------

    data.columns = [
        str(column).strip()
        for column in data.columns
    ]


    # --------------------------------------------------------
    # Required columns
    # --------------------------------------------------------

    required_columns = [
        "Product Name given",
        "Brand",
        "Site",
        "Posted On",
        "Stars",
        "Sentiment",
    ]

    required_columns.extend(
        THEME_COLS.values()
    )


    missing = [
        column
        for column in required_columns
        if column not in data.columns
    ]


    if missing:

        raise ValueError(
            "These required columns are missing:\n"
            + "\n".join(
                f"- {column}"
                for column in missing
            )
        )


    # --------------------------------------------------------
    # IMPORTANT:
    # Use exact product values from Excel.
    # Do NOT use contains() or partial matching.
    # --------------------------------------------------------

    data["Product"] = (
        data["Product Name given"]
        .astype("string")
        .str.strip()
    )


    # --------------------------------------------------------
    # Brand
    # --------------------------------------------------------

    data["Brand"] = (
        data["Brand"]
        .astype("string")
        .str.strip()
    )


    # --------------------------------------------------------
    # Site
    # --------------------------------------------------------

    data["Site"] = (
        data["Site"]
        .astype("string")
        .str.strip()
    )


    # --------------------------------------------------------
    # Sentiment
    # --------------------------------------------------------

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
        .astype("string")
        .str.strip()
        .str.lower()
        .map(sentiment_map)
    )


    # --------------------------------------------------------
    # Stars
    # --------------------------------------------------------

    data["Stars"] = pd.to_numeric(
        data["Stars"],
        errors="coerce",
    )


    # --------------------------------------------------------
    # Posted On
    # --------------------------------------------------------

    data["Posted On"] = pd.to_datetime(
        data["Posted On"],
        errors="coerce",
    )


    # --------------------------------------------------------
    # Theme normalization
    # --------------------------------------------------------

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


    def normalize_theme(value):

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

        data[column] = (
            data[column]
            .apply(normalize_theme)
        )


    # --------------------------------------------------------
    # KEEP EXACTLY FIVE PRODUCTS
    # --------------------------------------------------------

    data = data[
        data["Product"].isin(
            PRODUCT_ORDER
        )
    ].copy()


    # --------------------------------------------------------
    # REMOVE EXACT DUPLICATE ROWS ONLY
    # --------------------------------------------------------

    data = (
        data
        .drop_duplicates()
        .reset_index(drop=True)
    )


    # --------------------------------------------------------
    # Product ordering
    # --------------------------------------------------------

    data["Product"] = pd.Categorical(
        data["Product"],
        categories=PRODUCT_ORDER,
        ordered=True,
    )


    return data


# ============================================================
# LOAD
# ============================================================

try:

    df = load_data()

except Exception as error:

    st.error(
        f"Unable to load the dashboard data:\n\n{error}"
    )

    st.stop()


# ============================================================
# SOURCE DATA COUNT
# ============================================================

SOURCE_REVIEW_COUNT = len(df)


# ============================================================
# HELPER
# ============================================================

def percentage(
    numerator,
    denominator,
):

    if denominator == 0:

        return 0.0

    return (
        numerator
        / denominator
        * 100
    )


# ============================================================
# FILTER FUNCTION
# ============================================================

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


    # --------------------------------------------------------
    # PRODUCT
    #
    # Specific product selections take priority over
    # "All Products".
    # --------------------------------------------------------

    selected_specific_products = [
        product
        for product in products
        if product != "All Products"
    ]


    if selected_specific_products:

        filtered = filtered[
            filtered["Product"].isin(
                selected_specific_products
            )
        ]


    # --------------------------------------------------------
    # BRAND
    # --------------------------------------------------------

    if brands:

        filtered = filtered[
            filtered["Brand"].isin(
                brands
            )
        ]


    # --------------------------------------------------------
    # SITE
    # --------------------------------------------------------

    if sites:

        filtered = filtered[
            filtered["Site"].isin(
                sites
            )
        ]


    # --------------------------------------------------------
    # SENTIMENT
    # --------------------------------------------------------

    if sentiments:

        sentiment_values = [
            value.lower()
            for value in sentiments
        ]


        filtered = filtered[
            filtered["Sentiment"].isin(
                sentiment_values
            )
        ]


    # --------------------------------------------------------
    # STAR RATING
    # --------------------------------------------------------

    if star_range:

        filtered = filtered[
            filtered["Stars"].between(
                star_range[0],
                star_range[1],
                inclusive="both",
            )
        ]


    # --------------------------------------------------------
    # DATE
    # --------------------------------------------------------

    if (
        date_range
        and len(date_range) == 2
    ):

        start_date = pd.Timestamp(
            date_range[0]
        )


        end_date = (
            pd.Timestamp(
                date_range[1]
            )
            + pd.Timedelta(days=1)
            - pd.Timedelta(seconds=1)
        )


        filtered = filtered[
            filtered["Posted On"].between(
                start_date,
                end_date,
            )
        ]


    return filtered


# ============================================================
# PLOTLY STANDARD STYLE
# ============================================================

def style_chart(
    fig,
    height=390,
    show_legend=True,
):

    fig.update_layout(

        template="plotly_white",

        height=height,

        paper_bgcolor="#FFFFFF",

        plot_bgcolor="#FFFFFF",

        margin=dict(
            l=45,
            r=25,
            t=105,
            b=50,
        ),

        font=dict(
            family="Arial",
            size=12,
            color=TEXT,
        ),

        title=dict(

            x=0.02,

            xanchor="left",

            y=0.98,

            yanchor="top",

            font=dict(
                size=16,
                color=NAVY,
            ),
        ),
    )


    if show_legend:

        fig.update_layout(

            legend=dict(

                orientation="h",

                yanchor="top",

                y=0.87,

                xanchor="left",

                x=0,

                bgcolor=(
                    "rgba(255,255,255,0)"
                ),

                font=dict(
                    size=10,
                    color=MUTED,
                ),
            )
        )


    fig.update_xaxes(

        showgrid=True,

        gridcolor=GRID,

        zeroline=False,

        linecolor="#B8C6D1",
    )


    fig.update_yaxes(

        showgrid=False,

        zeroline=False,

        linecolor="#B8C6D1",
    )


    return fig


# ============================================================
# HEADER
# ============================================================

st.markdown(
    """
    <div class="dashboard-header">

        <div class="dashboard-title">
            Germany Toothbrush Customer Review Dashboard
        </div>

        <div class="dashboard-subtitle">
            Customer perception, ratings, sentiment and key product themes
        </div>

    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        "## Dashboard Filters"
    )


    # --------------------------------------------------------
    # Product
    # --------------------------------------------------------

    selected_products = st.multiselect(

        "Product",

        [
            "All Products"
        ] + PRODUCT_ORDER,

        default=[
            "All Products"
        ],
    )


    # --------------------------------------------------------
    # Brand
    # --------------------------------------------------------

    selected_brands = st.multiselect(

        "Brand",

        sorted(
            df["Brand"]
            .dropna()
            .unique()
            .tolist()
        ),
    )


    # --------------------------------------------------------
    # Site
    # --------------------------------------------------------

    selected_sites = st.multiselect(

        "Site",

        sorted(
            df["Site"]
            .dropna()
            .unique()
            .tolist()
        ),
    )


    # --------------------------------------------------------
    # Sentiment
    # --------------------------------------------------------

    selected_sentiments = st.multiselect(

        "Sentiment",

        [
            "Positive",
            "Neutral",
            "Negative",
        ],
    )


    # --------------------------------------------------------
    # Stars
    # --------------------------------------------------------

    selected_stars = st.slider(

        "Star Rating",

        min_value=1,

        max_value=5,

        value=(1, 5),

        step=1,
    )


    # --------------------------------------------------------
    # Dates
    # --------------------------------------------------------

    valid_dates = (
        df["Posted On"]
        .dropna()
    )


    min_date = (
        valid_dates
        .min()
        .date()
    )


    max_date = (
        valid_dates
        .max()
        .date()
    )


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


    # --------------------------------------------------------
    # RESET
    # --------------------------------------------------------

    if st.button(
        "Reset Filters",
        use_container_width=True,
    ):

        st.session_state.clear()

        st.rerun()


    # --------------------------------------------------------
    # Sidebar note
    # --------------------------------------------------------

    st.markdown(
        """
        <div class="sidebar-note">

        <b>Data scope</b><br>

        Germany customer reviews across five
        toothbrush products.

        <br><br>

        Theme fields with missing values remain
        missing and are not automatically treated
        as "No".

        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# FILTER DATA
# ============================================================

filtered_df = apply_filters(

    df,

    selected_products,

    selected_brands,

    selected_sites,

    selected_sentiments,

    selected_stars,

    selected_dates,
)


# ============================================================
# EMPTY RESULT
# ============================================================

if filtered_df.empty:

    st.warning(
        "No reviews match the selected filters."
    )

    st.stop()


# ============================================================
# REVIEW COUNT VALIDATION
# ============================================================

FILTERED_REVIEW_COUNT = len(
    filtered_df
)


FILTER_PERCENT = percentage(
    FILTERED_REVIEW_COUNT,
    SOURCE_REVIEW_COUNT,
)


# ============================================================
# KPI CALCULATIONS
# ============================================================

average_rating = (
    filtered_df["Stars"].mean()
)


positive_count = (
    filtered_df["Sentiment"]
    == "positive"
).sum()


negative_count = (
    filtered_df["Sentiment"]
    == "negative"
).sum()


five_star_count = (
    filtered_df["Stars"]
    == 5
).sum()


low_rating_count = (
    filtered_df["Stars"]
    .isin([1, 2])
).sum()


positive_percentage = percentage(
    positive_count,
    FILTERED_REVIEW_COUNT,
)


negative_percentage = percentage(
    negative_count,
    FILTERED_REVIEW_COUNT,
)


five_star_percentage = percentage(
    five_star_count,
    FILTERED_REVIEW_COUNT,
)


low_rating_percentage = percentage(
    low_rating_count,
    FILTERED_REVIEW_COUNT,
)


# ============================================================
# KPI SECTION
# ============================================================

st.markdown(
    """
    <div class="section-title">
        Key Performance Indicators
    </div>
    """,
    unsafe_allow_html=True,
)


k1, k2, k3, k4, k5, k6 = st.columns(6)


with k1:

    st.metric(
        "Total Reviews",
        f"{FILTERED_REVIEW_COUNT:,}",
    )


with k2:

    st.metric(
        "Average Rating",
        f"{average_rating:.2f} / 5",
    )


with k3:

    st.metric(
        "Positive Sentiment",
        f"{positive_percentage:.1f}%",
    )


with k4:

    st.metric(
        "Negative Sentiment",
        f"{negative_percentage:.1f}%",
    )


with k5:

    st.metric(
        "5-Star Reviews",
        f"{five_star_percentage:.1f}%",
    )


with k6:

    st.metric(
        "1–2 Star Reviews",
        f"{low_rating_percentage:.1f}%",
    )


# ============================================================
# DATA VALIDATION LINE
# ============================================================

st.markdown(
    f"""
    <div class="validation-card">

        <span class="validation-number">
            {FILTERED_REVIEW_COUNT:,}
        </span>

        filtered reviews out of

        <span class="validation-number">
            {SOURCE_REVIEW_COUNT:,}
        </span>

        source reviews

        &nbsp;•&nbsp;

        <span class="validation-number">
            {FILTER_PERCENT:.1f}%
        </span>

        of source reviews currently visible

    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# PRODUCT PERFORMANCE
# ============================================================

st.markdown(
    """
    <div class="section-title">
        Product Performance
    </div>
    """,
    unsafe_allow_html=True,
)


left, right = st.columns(2)


# ============================================================
# REVIEWS BY PRODUCT
# ============================================================

with left:

    product_review_counts = (

        filtered_df

        .groupby(
            "Product",
            observed=True,
        )

        .size()

        .reindex(
            PRODUCT_ORDER
        )

        .fillna(0)

        .astype(int)

        .reset_index(
            name="Reviews",
        )
    )


    product_review_counts = (
        product_review_counts[
            product_review_counts["Reviews"] > 0
        ]
    )


    fig = px.bar(

        product_review_counts,

        x="Reviews",

        y="Product",

        orientation="h",

        text="Reviews",

        title="Reviews by Product",

        labels={
            "Reviews":
                "Number of Reviews",

            "Product":
                "",
        },
    )


    fig.update_traces(

        marker_color=MEDIUM_BLUE,

        textposition="outside",

        textfont=dict(
            color=NAVY,
            size=11,
        ),
    )


    fig.update_layout(

        yaxis=dict(

            categoryorder="array",

            categoryarray=PRODUCT_ORDER,
        ),

        xaxis=dict(

            rangemode="tozero",
        ),
    )


    st.plotly_chart(

        style_chart(
            fig,
            show_legend=False,
        ),

        use_container_width=True,

        config={
            "displayModeBar": False,
        },
    )


# ============================================================
# AVERAGE RATING
# ============================================================

with right:

    product_rating = (

        filtered_df

        .groupby(
            "Product",
            observed=True,
        )["Stars"]

        .mean()

        .reindex(
            PRODUCT_ORDER
        )

        .dropna()

        .reset_index(
            name="Average Rating",
        )
    )


    fig = px.bar(

        product_rating,

        x="Average Rating",

        y="Product",

        orientation="h",

        text=(
            product_rating[
                "Average Rating"
            ]
            .round(2)
        ),

        title="Average Rating by Product",

        labels={
            "Average Rating":
                "Average Rating",

            "Product":
                "",
        },
    )


    fig.update_traces(

        marker_color=BLUE,

        textposition="outside",

        textfont=dict(
            color=NAVY,
            size=11,
        ),
    )


    fig.update_xaxes(
        range=[0, 5],
    )


    fig.update_layout(

        yaxis=dict(

            categoryorder="array",

            categoryarray=PRODUCT_ORDER,
        ),
    )


    st.plotly_chart(

        style_chart(
            fig,
            show_legend=False,
        ),

        use_container_width=True,

        config={
            "displayModeBar": False,
        },
    )


# ============================================================
# CUSTOMER SENTIMENT
# ============================================================

st.markdown(
    """
    <div class="section-title">
        Customer Sentiment
    </div>
    """,
    unsafe_allow_html=True,
)


left, right = st.columns(2)


# ============================================================
# SENTIMENT BY PRODUCT
# ============================================================

with left:

    sentiment_rows = []


    for product in PRODUCT_ORDER:

        product_df = filtered_df[
            filtered_df["Product"]
            == product
        ]


        product_total = len(
            product_df
        )


        if product_total == 0:

            continue


        for sentiment in [
            "positive",
            "neutral",
            "negative",
        ]:

            sentiment_rows.append(

                {
                    "Product":
                        product,

                    "Sentiment":
                        sentiment.title(),

                    "Percentage":
                        percentage(

                            (
                                product_df[
                                    "Sentiment"
                                ]
                                == sentiment
                            ).sum(),

                            product_total,
                        ),
                }
            )


    sentiment_product = pd.DataFrame(
        sentiment_rows
    )


    fig = px.bar(

        sentiment_product,

        x="Percentage",

        y="Product",

        color="Sentiment",

        orientation="h",

        title="Sentiment by Product",

        labels={
            "Percentage":
                "Share of Reviews (%)",

            "Product":
                "",
        },

        color_discrete_map={

            "Positive":
                POSITIVE,

            "Neutral":
                NEUTRAL,

            "Negative":
                NEGATIVE,
        },
    )


    fig.update_layout(

        barmode="stack",

        xaxis=dict(
            range=[0, 100],
        ),

        yaxis=dict(

            categoryorder="array",

            categoryarray=PRODUCT_ORDER,
        ),
    )


    st.plotly_chart(

        style_chart(fig),

        use_container_width=True,

        config={
            "displayModeBar": False,
        },
    )


# ============================================================
# OVERALL SENTIMENT
# ============================================================

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

            "Positive":
                POSITIVE,

            "Neutral":
                NEUTRAL,

            "Negative":
                NEGATIVE,
        },
    )


    fig.update_traces(
        textinfo="percent",
        textfont=dict(
            size=12,
        ),
    )


    st.plotly_chart(

        style_chart(
            fig,
            show_legend=False,
        ),

        use_container_width=True,

        config={
            "displayModeBar": False,
        },
    )


# ============================================================
# CUSTOMER THEMES
# ============================================================

st.markdown(
    """
    <div class="section-title">
        What Customers Talk About
    </div>
    """,
    unsafe_allow_html=True,
)


theme_rows = []


for theme, column in THEME_COLS.items():

    theme_mentions = (
        filtered_df[column]
        == "Yes"
    ).sum()


    theme_rows.append(

        {
            "Theme":
                theme,

            "Reviews":
                int(theme_mentions),

            "Percentage":
                percentage(
                    theme_mentions,
                    FILTERED_REVIEW_COUNT,
                ),
        }
    )


theme_summary = (

    pd.DataFrame(
        theme_rows
    )

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
        "Percentage":
            "Reviews Mentioning Theme (%)",

        "Theme":
            "",
    },
)


fig.update_traces(

    marker_color=MEDIUM_BLUE,

    textposition="outside",

    textfont=dict(
        color=NAVY,
        size=10,
    ),
)


st.plotly_chart(

    style_chart(
        fig,
        height=410,
        show_legend=False,
    ),

    use_container_width=True,

    config={
        "displayModeBar": False,
    },
)


# ============================================================
# PRODUCT × THEME HEATMAP
# ============================================================

theme_matrix = pd.DataFrame(

    index=PRODUCT_ORDER,

    columns=THEMES,

    dtype=float,
)


for product in PRODUCT_ORDER:

    product_df = filtered_df[
        filtered_df["Product"]
        == product
    ]


    for theme, column in THEME_COLS.items():

        mentions = (
            product_df[column]
            == "Yes"
        ).sum()


        theme_matrix.loc[
            product,
            theme,
        ] = (

            percentage(
                mentions,
                len(product_df),
            )

            if len(product_df) > 0

            else np.nan
        )


fig = px.imshow(

    theme_matrix,

    text_auto=".1f",

    aspect="auto",

    color_continuous_scale=[
        "#F0F3F5",
        "#91ABC0",
        "#2F5D8C",
    ],

    title="Product vs Customer Themes",

    labels={
        "color":
            "Theme Mentions (%)",
    },
)


fig.update_traces(

    hovertemplate=(

        "Product: %{y}<br>"

        "Theme: %{x}<br>"

        "Mention Rate: %{z:.1f}%"

        "<extra></extra>"
    ),
)


st.plotly_chart(

    style_chart(
        fig,
        height=410,
        show_legend=False,
    ),

    use_container_width=True,

    config={
        "displayModeBar": False,
    },
)


# ============================================================
# THEME SENTIMENT
# ============================================================

st.markdown(
    """
    <div class="section-title">
        Customer Sentiment by Theme
    </div>
    """,
    unsafe_allow_html=True,
)


left, right = st.columns(2)


for container, sentiment, chart_title in [

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

    sentiment_matrix = pd.DataFrame(

        index=PRODUCT_ORDER,

        columns=THEMES,

        dtype=float,
    )


    for product in PRODUCT_ORDER:

        product_df = filtered_df[
            filtered_df["Product"]
            == product
        ]


        for theme, column in THEME_COLS.items():

            theme_df = product_df[
                product_df[column]
                == "Yes"
            ]


            theme_total = len(
                theme_df
            )


            sentiment_matrix.loc[
                product,
                theme,
            ] = (

                percentage(

                    (
                        theme_df[
                            "Sentiment"
                        ]
                        == sentiment
                    ).sum(),

                    theme_total,
                )

                if theme_total > 0

                else np.nan
            )


    fig = px.imshow(

        sentiment_matrix,

        text_auto=".1f",

        aspect="auto",

        color_continuous_scale=[
            "#F0F3F5",
            "#91ABC0",
            "#2F5D8C",
        ],

        title=chart_title,

        labels={
            "color":
                "Sentiment (%)",
        },
    )


    fig.update_traces(

        hovertemplate=(

            "Product: %{y}<br>"

            "Theme: %{x}<br>"

            "Sentiment: %{z:.1f}%"

            "<extra></extra>"
        ),
    )


    with container:

        st.plotly_chart(

            style_chart(
                fig,
                height=410,
                show_legend=False,
            ),

            use_container_width=True,

            config={
                "displayModeBar": False,
            },
        )


# ============================================================
# RATING DISTRIBUTION
# ============================================================

st.markdown(
    """
    <div class="section-title">
        Customer Rating Distribution
    </div>
    """,
    unsafe_allow_html=True,
)


rating_rows = []


for product in PRODUCT_ORDER:

    product_df = filtered_df[
        filtered_df["Product"]
        == product
    ]


    product_total = len(
        product_df
    )


    if product_total == 0:

        continue


    for star in range(1, 6):

        star_count = (
            product_df["Stars"]
            == star
        ).sum()


        rating_rows.append(

            {
                "Product":
                    product,

                "Star":
                    str(star),

                "Percentage":
                    percentage(
                        star_count,
                        product_total,
                    ),
            }
        )


rating_distribution = pd.DataFrame(
    rating_rows
)


fig = px.bar(

    rating_distribution,

    x="Star",

    y="Percentage",

    color="Product",

    barmode="group",

    title="Customer Rating Distribution",

    labels={
        "Percentage":
            "Reviews (%)",

        "Star":
            "Star Rating",
    },

    category_orders={
        "Star": [
            "1",
            "2",
            "3",
            "4",
            "5",
        ],
    },

    color_discrete_sequence=
        PRODUCT_COLORS,
)


fig.update_traces(

    texttemplate="%{y:.0f}%",

    textposition="outside",
)


st.plotly_chart(

    style_chart(
        fig,
        height=420,
    ),

    use_container_width=True,

    config={
        "displayModeBar": False,
    },
)


# ============================================================
# REVIEW TREND
# ============================================================

st.markdown(
    """
    <div class="section-title">
        Review Activity Over Time
    </div>
    """,
    unsafe_allow_html=True,
)


trend_data = (
    filtered_df
    .dropna(
        subset=["Posted On"]
    )
    .copy()
)


monthly = (

    trend_data

    .assign(

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
        name="Reviews",
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
        "Reviews":
            "Number of Reviews",

        "Month":
            "Month",
    },

    color_discrete_sequence=
        PRODUCT_COLORS,
)


fig.update_layout(
    hovermode="x unified",
)


fig.update_traces(

    line=dict(
        width=2,
    ),

    marker=dict(
        size=5,
    ),
)


st.plotly_chart(

    style_chart(
        fig,
        height=420,
    ),

    use_container_width=True,

    config={
        "displayModeBar": False,
    },
)


# ============================================================
# KEY CUSTOMER FEEDBACK
# ============================================================

st.markdown(
    """
    <div class="section-title">
        Key Customer Feedback
    </div>
    """,
    unsafe_allow_html=True,
)


strength_rows = []

pain_rows = []


for theme, column in THEME_COLS.items():

    theme_df = filtered_df[
        filtered_df[column]
        == "Yes"
    ]


    theme_total = len(
        theme_df
    )


    if theme_total == 0:

        continue


    positive_theme_count = (
        theme_df["Sentiment"]
        == "positive"
    ).sum()


    negative_theme_count = (
        theme_df["Sentiment"]
        == "negative"
    ).sum()


    strength_rows.append(

        {
            "Theme":
                theme,

            "Reviews":
                theme_total,

            "Positive %":
                round(
                    percentage(
                        positive_theme_count,
                        theme_total,
                    ),
                    1,
                ),
        }
    )


    pain_rows.append(

        {
            "Theme":
                theme,

            "Reviews":
                theme_total,

            "Negative %":
                round(
                    percentage(
                        negative_theme_count,
                        theme_total,
                    ),
                    1,
                ),
        }
    )


strengths = (

    pd.DataFrame(
        strength_rows
    )

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


pain_points = (

    pd.DataFrame(
        pain_rows
    )

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
                ),
        },
    )


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
                ),
        },
    )


# ============================================================
# PRODUCT SUMMARY
# ============================================================

st.markdown(
    """
    <div class="section-title">
        Product Summary
    </div>
    """,
    unsafe_allow_html=True,
)


summary_rows = []


for product in PRODUCT_ORDER:

    product_df = filtered_df[
        filtered_df["Product"]
        == product
    ]


    product_total = len(
        product_df
    )


    if product_total == 0:

        continue


    summary_rows.append(

        {
            "Product":
                product,

            # IMPORTANT:
            # This is the exact same row count used
            # in the Reviews by Product chart.
            "Reviews":
                product_total,

            "Average Rating":
                round(
                    product_df["Stars"].mean(),
                    2,
                ),

            "Positive %":
                round(
                    percentage(
                        (
                            product_df[
                                "Sentiment"
                            ]
                            == "positive"
                        ).sum(),

                        product_total,
                    ),
                    1,
                ),

            "Neutral %":
                round(
                    percentage(
                        (
                            product_df[
                                "Sentiment"
                            ]
                            == "neutral"
                        ).sum(),

                        product_total,
                    ),
                    1,
                ),

            "Negative %":
                round(
                    percentage(
                        (
                            product_df[
                                "Sentiment"
                            ]
                            == "negative"
                        ).sum(),

                        product_total,
                    ),
                    1,
                ),
        }
    )


product_summary = (

    pd.DataFrame(
        summary_rows
    )

    .sort_values(
        "Average Rating",
        ascending=False,
    )
)


st.dataframe(

    product_summary,

    use_container_width=True,

    hide_index=True,
)


# ============================================================
# FINAL COUNT RECONCILIATION
# ============================================================

# This check confirms that Product Summary counts add up
# exactly to the Total Reviews KPI.

summary_count_total = int(
    product_summary["Reviews"].sum()
)


if summary_count_total != FILTERED_REVIEW_COUNT:

    st.error(
        "Review count reconciliation failed: "
        f"Product Summary = {summary_count_total:,}, "
        f"Total Reviews = {FILTERED_REVIEW_COUNT:,}."
    )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    f"""
    <div style="
        margin-top:18px;
        padding-top:10px;
        border-top:1px solid #CDD8E2;
        color:#637586;
        font-size:12px;
    ">
        Showing <b>{FILTERED_REVIEW_COUNT:,}</b>
        filtered reviews from
        <b>{SOURCE_REVIEW_COUNT:,}</b>
        source reviews.
    </div>
    """,
    unsafe_allow_html=True,
)
