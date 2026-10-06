from pathlib import Path
import warnings

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

warnings.filterwarnings("ignore")


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Germany Toothbrush Customer Review Dashboard",
    page_icon="🪥",
    layout="wide",
    initial_sidebar_state="expanded",
)


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


THEME_COLS = {
    theme: f"{theme}_Present"
    for theme in THEMES
}


# =========================================================
# COLORS
# =========================================================

# Muted professional colors
# Similar contrast to a business analytics dashboard.

SENTIMENT_COLORS = {
    "Positive": "#6E9278",
    "Neutral": "#8C969E",
    "Negative": "#B46F70",
}


PRODUCT_COLORS = [
    "#2F5D7C",
    "#477B9D",
    "#6492AF",
    "#7FA6BD",
    "#9AB8C9",
]


# =========================================================
# GLOBAL CSS
# =========================================================

st.markdown(
    """
    <style>

        /* -------------------------------------------------
           Overall Dashboard
        ------------------------------------------------- */

        .stApp {
            background: #F2F5F7;
        }

        .block-container {
            max-width: 1500px;
            padding-top: 1.25rem;
            padding-bottom: 2rem;
        }


        /* -------------------------------------------------
           Main Title
        ------------------------------------------------- */

        h1 {
            color: #213F56 !important;
            font-size: 28px !important;
            font-weight: 700 !important;
            margin-bottom: 0.1rem !important;
        }


        .dashboard-subtitle {
            color: #647687;
            font-size: 14px;
            margin-bottom: 1.2rem;
        }


        /* -------------------------------------------------
           Section Titles
        ------------------------------------------------- */

        .section-title {
            color: #294C66;
            font-size: 19px;
            font-weight: 700;
            margin: 24px 0 10px 0;
            padding-bottom: 7px;
            border-bottom: 1px solid #D3DDE5;
        }


        /* -------------------------------------------------
           KPI Cards
        ------------------------------------------------- */

        [data-testid="stMetric"] {
            background: #FFFFFF;
            border: 1px solid #D3DDE5;
            border-radius: 8px;
            padding: 13px 15px;
            min-height: 96px;
            box-shadow: 0 1px 2px rgba(31,55,72,.04);
        }


        [data-testid="stMetricLabel"] {
            color: #607384 !important;
            font-size: 12px !important;
            font-weight: 500 !important;
        }


        [data-testid="stMetricValue"] {
            color: #294C66 !important;
            font-size: 25px !important;
            font-weight: 700 !important;
        }


        /* -------------------------------------------------
           Sidebar
        ------------------------------------------------- */

        section[data-testid="stSidebar"] {
            background: #E9EEF2;
            border-right: 1px solid #D2DCE4;
        }


        section[data-testid="stSidebar"] h2 {
            color: #294C66;
        }


        /* -------------------------------------------------
           Buttons
        ------------------------------------------------- */

        .stButton > button {
            background: #FFFFFF;
            color: #294C66;
            border: 1px solid #B8C8D4;
            border-radius: 6px;
            font-weight: 600;
        }


        .stButton > button:hover {
            background: #EEF3F6;
            border-color: #7F9DB2;
            color: #213F56;
        }


        /* -------------------------------------------------
           Tables
        ------------------------------------------------- */

        [data-testid="stDataFrame"] {
            border: 1px solid #D5DEE5;
            border-radius: 8px;
        }


        /* -------------------------------------------------
           Sidebar Note
        ------------------------------------------------- */

        .sidebar-note {
            color: #607384;
            font-size: 12px;
            line-height: 1.5;
        }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# LOAD DATA
# =========================================================

@st.cache_data
def load_data():

    if not DATA_FILE.exists():

        raise FileNotFoundError(
            f"{DATA_FILE.name} was not found. "
            "Keep it in the same folder as app.py."
        )


    df = pd.read_excel(
        DATA_FILE,
        sheet_name=SHEET_NAME
    )


    df.columns = [
        str(column).strip()
        for column in df.columns
    ]


    # -----------------------------------------------------
    # Identify product column
    # -----------------------------------------------------

    if "Product Name" in df.columns:

        product_source = "Product Name"

    elif "Product Name given" in df.columns:

        product_source = "Product Name given"

    else:

        raise ValueError(
            "The Excel file must contain "
            "'Product Name' or 'Product Name given'."
        )


    # -----------------------------------------------------
    # Product standardization
    # -----------------------------------------------------

    def standardize_product(value):

        if pd.isna(value):
            return np.nan


        text = (
            str(value)
            .strip()
            .lower()
        )


        if (
            "curaprox" in text
            and "5460" in text
        ):

            return "curaprox 5460"


        if (
            "dr. best" in text
            and "clean pro" in text
        ):

            return "Dr. Best Clean Pro Zwischenzahn"


        if (
            "elmex" in text
            and "expert" in text
            and "precision" in text
        ):

            return "elmex expert precision"


        if (
            "elmex" in text
            and "interx" in text
        ):

            return "elmex InterX"


        if (
            "meridol" in text
            and "base" in text
        ):

            return "meridol base"


        return np.nan


    df["Product"] = (
        df[product_source]
        .apply(standardize_product)
    )


    # -----------------------------------------------------
    # Required columns
    # -----------------------------------------------------

    required_columns = [
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
        if column not in df.columns
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
    # Brand / Site
    # -----------------------------------------------------

    df["Brand"] = (
        df["Brand"]
        .astype(str)
        .str.strip()
    )


    df["Site"] = (
        df["Site"]
        .astype(str)
        .str.strip()
    )


    # -----------------------------------------------------
    # Sentiment
    # -----------------------------------------------------

    sentiment_map = {

        "positive": "positive",
        "pos": "positive",

        "neutral": "neutral",
        "neu": "neutral",

        "negative": "negative",
        "neg": "negative",
    }


    df["Sentiment"] = (
        df["Sentiment"]
        .astype(str)
        .str.strip()
        .str.lower()
        .map(sentiment_map)
    )


    # -----------------------------------------------------
    # Rating
    # -----------------------------------------------------

    df["Stars"] = pd.to_numeric(
        df["Stars"],
        errors="coerce"
    )


    # -----------------------------------------------------
    # Date
    # -----------------------------------------------------

    df["Posted On"] = pd.to_datetime(
        df["Posted On"],
        errors="coerce"
    )


    # -----------------------------------------------------
    # Theme normalization
    # -----------------------------------------------------

    def normalize_theme(value):

        if pd.isna(value):
            return np.nan


        value = (
            str(value)
            .strip()
            .lower()
        )


        if value in {
            "yes",
            "y",
            "true",
            "1",
        }:

            return "Yes"


        if value in {
            "no",
            "n",
            "false",
            "0",
        }:

            return "No"


        return np.nan


    for column in THEME_COLS.values():

        df[column] = (
            df[column]
            .apply(normalize_theme)
        )


    # -----------------------------------------------------
    # Keep only requested products
    # -----------------------------------------------------

    df = df[
        df["Product"].isin(PRODUCTS)
    ].copy()


    # -----------------------------------------------------
    # Remove duplicate reviews
    # -----------------------------------------------------

    df = (
        df
        .drop_duplicates()
        .reset_index(drop=True)
    )


    # -----------------------------------------------------
    # Product ordering
    # -----------------------------------------------------

    df["Product"] = pd.Categorical(
        df["Product"],
        categories=PRODUCTS,
        ordered=True,
    )


    return df


# =========================================================
# LOAD DATA
# =========================================================

try:

    df = load_data()

except Exception as error:

    st.error(
        f"Unable to load the data.\n\n{error}"
    )

    st.stop()


# =========================================================
# HELPER FUNCTIONS
# =========================================================

def percentage(
    part,
    total
):

    if total == 0:
        return 0

    return (
        part
        / total
        * 100
    )


def filter_data(
    data,
    products,
    brands,
    sites,
    sentiments,
    stars,
    dates,
):

    filtered = data.copy()


    # -----------------------------------------------------
    # Product
    # -----------------------------------------------------

    if (
        products
        and "All Products" not in products
    ):

        filtered = filtered[
            filtered["Product"].isin(
                products
            )
        ]


    # -----------------------------------------------------
    # Brand
    # -----------------------------------------------------

    if brands:

        filtered = filtered[
            filtered["Brand"].isin(
                brands
            )
        ]


    # -----------------------------------------------------
    # Site
    # -----------------------------------------------------

    if sites:

        filtered = filtered[
            filtered["Site"].isin(
                sites
            )
        ]


    # -----------------------------------------------------
    # Sentiment
    # -----------------------------------------------------

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


    # -----------------------------------------------------
    # Star rating
    # -----------------------------------------------------

    filtered = filtered[
        filtered["Stars"].between(
            stars[0],
            stars[1],
            inclusive="both",
        )
    ]


    # -----------------------------------------------------
    # Date
    # -----------------------------------------------------

    if (
        dates
        and len(dates) == 2
    ):

        start_date = pd.Timestamp(
            dates[0]
        )


        end_date = (
            pd.Timestamp(
                dates[1]
            )
            + pd.Timedelta(days=1)
        )


        filtered = filtered[
            (
                filtered["Posted On"]
                >= start_date
            )
            &
            (
                filtered["Posted On"]
                < end_date
            )
        ]


    return filtered


# =========================================================
# COMMON PLOTLY STYLE
# =========================================================

def chart_layout(
    fig,
    height=390,
    legend=True,
):

    fig.update_layout(

        template="plotly_white",

        height=height,

        margin=dict(
            l=45,
            r=25,
            t=105,
            b=50,
        ),

        paper_bgcolor="#FFFFFF",

        plot_bgcolor="#FFFFFF",

        font=dict(
            family="Arial",
            size=12,
            color="#536879",
        ),

        # -------------------------------------------------
        # Title
        # -------------------------------------------------

        title=dict(
            x=0,
            xanchor="left",
            y=0.98,
            yanchor="top",
            font=dict(
                size=16,
                color="#35546D",
            ),
        ),
    )


    # -----------------------------------------------------
    # Legend
    # -----------------------------------------------------

    if legend:

        fig.update_layout(

            legend=dict(

                orientation="h",

                yanchor="top",

                y=0.87,

                xanchor="left",

                x=0,

                font=dict(
                    size=10,
                    color="#607384",
                ),

                bgcolor=(
                    "rgba(255,255,255,0)"
                ),
            )
        )


    # -----------------------------------------------------
    # X axis
    # -----------------------------------------------------

    fig.update_xaxes(

        showgrid=True,

        gridcolor="#D5DEE5",

        zeroline=False,

        linecolor="#BCC9D3",
    )


    # -----------------------------------------------------
    # Y axis
    # -----------------------------------------------------

    fig.update_yaxes(

        showgrid=False,

        zeroline=False,

        linecolor="#BCC9D3",
    )


    return fig


# =========================================================
# HEADER
# =========================================================

st.title(
    "Germany Toothbrush Customer Review Dashboard"
)


st.markdown(
    """
    <div class="dashboard-subtitle">
        Customer perception, ratings, sentiment and key product themes
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown(
        "## Dashboard Filters"
    )


    # -----------------------------------------------------
    # Product
    # -----------------------------------------------------

    selected_products = st.multiselect(

        "Product",

        [
            "All Products"
        ] + PRODUCTS,

        default=[
            "All Products"
        ],
    )


    # -----------------------------------------------------
    # Brand
    # -----------------------------------------------------

    selected_brands = st.multiselect(

        "Brand",

        sorted(
            df["Brand"]
            .dropna()
            .unique()
            .tolist()
        ),
    )


    # -----------------------------------------------------
    # Site
    # -----------------------------------------------------

    selected_sites = st.multiselect(

        "Site",

        sorted(
            df["Site"]
            .dropna()
            .unique()
            .tolist()
        ),
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
    # Star rating
    # -----------------------------------------------------

    selected_stars = st.slider(

        "Star Rating",

        min_value=1,

        max_value=5,

        value=(1, 5),

        step=1,
    )


    # -----------------------------------------------------
    # Date range
    # -----------------------------------------------------

    valid_dates = (
        df["Posted On"]
        .dropna()
    )


    if not valid_dates.empty:

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


    else:

        selected_dates = None


    st.markdown("---")


    # -----------------------------------------------------
    # Reset
    # -----------------------------------------------------

    if st.button(
        "Reset Filters",
        use_container_width=True,
    ):

        st.session_state.clear()

        st.rerun()


    # -----------------------------------------------------
    # Sidebar note
    # -----------------------------------------------------

    st.markdown(
        """
        <div class="sidebar-note">

        <b>Data scope</b><br>

        Germany customer reviews across five
        toothbrush products.

        <br><br>

        Missing theme values remain missing
        and are not treated as "No".

        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# APPLY FILTERS
# =========================================================

filtered = filter_data(

    df,

    selected_products,

    selected_brands,

    selected_sites,

    selected_sentiments,

    selected_stars,

    selected_dates,
)


# =========================================================
# EMPTY DATA CHECK
# =========================================================

if filtered.empty:

    st.warning(
        "No reviews match the selected filters. "
        "Please adjust the filters."
    )

    st.stop()


# =========================================================
# KPI CALCULATIONS
# =========================================================

total_reviews = len(
    filtered
)


avg_rating = (
    filtered["Stars"].mean()
)


positive_pct = percentage(

    (
        filtered["Sentiment"]
        == "positive"
    ).sum(),

    total_reviews,
)


negative_pct = percentage(

    (
        filtered["Sentiment"]
        == "negative"
    ).sum(),

    total_reviews,
)


five_star_pct = percentage(

    (
        filtered["Stars"]
        == 5
    ).sum(),

    total_reviews,
)


low_rating_pct = percentage(

    filtered["Stars"]
    .isin([1, 2])
    .sum(),

    total_reviews,
)


# =========================================================
# KPI SECTION
# =========================================================

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
        f"{total_reviews:,}",
    )


with k2:

    st.metric(
        "Average Rating",
        f"{avg_rating:.2f} / 5",
    )


with k3:

    st.metric(
        "Positive Sentiment",
        f"{positive_pct:.1f}%",
    )


with k4:

    st.metric(
        "Negative Sentiment",
        f"{negative_pct:.1f}%",
    )


with k5:

    st.metric(
        "5-Star Reviews",
        f"{five_star_pct:.1f}%",
    )


with k6:

    st.metric(
        "1–2 Star Reviews",
        f"{low_rating_pct:.1f}%",
    )


# =========================================================
# PRODUCT PERFORMANCE
# =========================================================

st.markdown(
    """
    <div class="section-title">
        Product Performance
    </div>
    """,
    unsafe_allow_html=True,
)


left, right = st.columns(2)


# ---------------------------------------------------------
# Reviews by Product
# ---------------------------------------------------------

with left:

    product_reviews = (

        filtered

        .groupby(
            "Product",
            observed=True,
        )

        .size()

        .reindex(
            PRODUCTS
        )

        .dropna()

        .reset_index(
            name="Reviews"
        )
    )


    fig = px.bar(

        product_reviews,

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

        marker_color="#477B9D",

        textposition="outside",
    )


    fig.update_layout(

        yaxis=dict(

            categoryorder="array",

            categoryarray=PRODUCTS,
        )
    )


    st.plotly_chart(

        chart_layout(
            fig,
            legend=False,
        ),

        use_container_width=True,

        config={
            "displayModeBar": False
        },
    )


# ---------------------------------------------------------
# Average Rating
# ---------------------------------------------------------

with right:

    product_rating = (

        filtered

        .groupby(
            "Product",
            observed=True,
        )["Stars"]

        .mean()

        .reindex(
            PRODUCTS
        )

        .dropna()

        .reset_index(
            name="Average Rating"
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

        marker_color="#2F5D7C",

        textposition="outside",
    )


    fig.update_xaxes(
        range=[0, 5]
    )


    fig.update_layout(

        yaxis=dict(

            categoryorder="array",

            categoryarray=PRODUCTS,
        )
    )


    st.plotly_chart(

        chart_layout(
            fig,
            legend=False,
        ),

        use_container_width=True,

        config={
            "displayModeBar": False
        },
    )


# =========================================================
# CUSTOMER SENTIMENT
# =========================================================

st.markdown(
    """
    <div class="section-title">
        Customer Sentiment
    </div>
    """,
    unsafe_allow_html=True,
)


left, right = st.columns(2)


# ---------------------------------------------------------
# Sentiment by Product
# ---------------------------------------------------------

with left:

    rows = []


    for product in PRODUCTS:

        temp = filtered[
            filtered["Product"]
            == product
        ]


        total = len(
            temp
        )


        if total == 0:
            continue


        for sentiment in [
            "positive",
            "neutral",
            "negative",
        ]:

            rows.append(

                {
                    "Product":
                        product,

                    "Sentiment":
                        sentiment.title(),

                    "Percentage":
                        percentage(

                            (
                                temp["Sentiment"]
                                == sentiment
                            ).sum(),

                            total,
                        ),
                }
            )


    sentiment_product = pd.DataFrame(
        rows
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

        color_discrete_map=
            SENTIMENT_COLORS,
    )


    fig.update_layout(

        barmode="stack",

        xaxis=dict(
            range=[0, 100]
        ),

        yaxis=dict(

            categoryorder="array",

            categoryarray=PRODUCTS,
        ),
    )


    st.plotly_chart(

        chart_layout(fig),

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

        filtered["Sentiment"]

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


    sentiment_total[
        "Sentiment"
    ] = (
        sentiment_total[
            "Sentiment"
        ]
        .str.title()
    )


    fig = px.pie(

        sentiment_total,

        names="Sentiment",

        values="Reviews",

        hole=0.58,

        title="Overall Sentiment",

        color="Sentiment",

        color_discrete_map=
            SENTIMENT_COLORS,
    )


    fig.update_traces(
        textinfo="percent",
    )


    st.plotly_chart(

        chart_layout(
            fig,
            legend=False,
        ),

        use_container_width=True,

        config={
            "displayModeBar": False
        },
    )


# =========================================================
# CUSTOMER THEMES
# =========================================================

st.markdown(
    """
    <div class="section-title">
        Customer Themes
    </div>
    """,
    unsafe_allow_html=True,
)


theme_rows = []


for theme, column in THEME_COLS.items():

    mentions = (
        filtered[column]
        == "Yes"
    ).sum()


    theme_rows.append(

        {
            "Theme":
                theme,

            "Reviews":
                mentions,

            "Percentage":
                percentage(
                    mentions,
                    total_reviews,
                ),
        }
    )


theme_df = (

    pd.DataFrame(
        theme_rows
    )

    .sort_values(
        "Percentage",
        ascending=True,
    )
)


fig = px.bar(

    theme_df,

    x="Percentage",

    y="Theme",

    orientation="h",

    text=(
        theme_df["Percentage"]
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

    marker_color="#477B9D",

    textposition="outside",
)


st.plotly_chart(

    chart_layout(
        fig,
        410,
        legend=False,
    ),

    use_container_width=True,

    config={
        "displayModeBar": False
    },
)


# =========================================================
# PRODUCT × THEME HEATMAP
# =========================================================

theme_matrix = pd.DataFrame(

    index=PRODUCTS,

    columns=THEMES,

    dtype=float,
)


for product in PRODUCTS:

    product_df = filtered[
        filtered["Product"]
        == product
    ]


    for theme, column in THEME_COLS.items():

        theme_matrix.loc[
            product,
            theme
        ] = (

            percentage(

                (
                    product_df[column]
                    == "Yes"
                ).sum(),

                len(product_df),
            )

            if len(product_df)

            else np.nan
        )


fig = px.imshow(

    theme_matrix,

    text_auto=".1f",

    aspect="auto",

    color_continuous_scale=[
        "#EEF2F5",
        "#A9BCC9",
        "#527892",
    ],

    title="Product × Customer Theme",

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
    )
)


st.plotly_chart(

    chart_layout(
        fig,
        410,
        legend=False,
    ),

    use_container_width=True,

    config={
        "displayModeBar": False
    },
)


# =========================================================
# THEME SENTIMENT
# =========================================================

st.markdown(
    """
    <div class="section-title">
        Theme Sentiment
    </div>
    """,
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

        index=PRODUCTS,

        columns=THEMES,

        dtype=float,
    )


    for product in PRODUCTS:

        product_df = filtered[
            filtered["Product"]
            == product
        ]


        for theme, column in THEME_COLS.items():

            theme_df = product_df[
                product_df[column]
                == "Yes"
            ]


            matrix.loc[
                product,
                theme
            ] = (

                percentage(

                    (
                        theme_df["Sentiment"]
                        == sentiment
                    ).sum(),

                    len(theme_df),
                )

                if len(theme_df)

                else np.nan
            )


    fig = px.imshow(

        matrix,

        text_auto=".1f",

        aspect="auto",

        color_continuous_scale=[
            "#EEF2F5",
            "#A9BCC9",
            "#527892",
        ],

        title=title,

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
        )
    )


    with container:

        st.plotly_chart(

            chart_layout(
                fig,
                410,
                legend=False,
            ),

            use_container_width=True,

            config={
                "displayModeBar": False
            },
        )


# =========================================================
# RATING DISTRIBUTION
# =========================================================

st.markdown(
    """
    <div class="section-title">
        Rating Distribution
    </div>
    """,
    unsafe_allow_html=True,
)


rating_rows = []


for product in PRODUCTS:

    temp = filtered[
        filtered["Product"]
        == product
    ]


    total = len(
        temp
    )


    if total == 0:
        continue


    for star in range(1, 6):

        rating_rows.append(

            {
                "Product":
                    product,

                "Star":
                    str(star),

                "Percentage":
                    percentage(

                        (
                            temp["Stars"]
                            == star
                        ).sum(),

                        total,
                    ),
            }
        )


rating_df = pd.DataFrame(
    rating_rows
)


fig = px.bar(

    rating_df,

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
        "Star":
            [
                "1",
                "2",
                "3",
                "4",
                "5",
            ]
    },

    color_discrete_sequence=
        PRODUCT_COLORS,
)


fig.update_traces(

    hovertemplate=(

        "<b>%{fullData.name}</b><br>"

        "Rating: %{x} stars<br>"

        "Share: %{y:.1f}%"

        "<extra></extra>"
    )
)


st.plotly_chart(

    chart_layout(
        fig,
        420,
    ),

    use_container_width=True,

    config={
        "displayModeBar": False
    },
)


# =========================================================
# REVIEW TREND
# =========================================================

st.markdown(
    """
    <div class="section-title">
        Review Activity Over Time
    </div>
    """,
    unsafe_allow_html=True,
)


trend = (
    filtered
    .dropna(
        subset=["Posted On"]
    )
    .copy()
)


monthly = (

    trend

    .assign(

        Month=(
            trend["Posted On"]
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
        "Month":
            "Month",

        "Reviews":
            "Number of Reviews",
    },

    color_discrete_sequence=
        PRODUCT_COLORS,
)


fig.update_layout(
    hovermode="x unified"
)


fig.update_traces(

    line=dict(
        width=2
    ),

    marker=dict(
        size=5
    ),
)


st.plotly_chart(

    chart_layout(
        fig,
        420,
    ),

    use_container_width=True,

    config={
        "displayModeBar": False
    },
)


# =========================================================
# KEY CUSTOMER FEEDBACK
# =========================================================

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

    theme_df = filtered[
        filtered[column]
        == "Yes"
    ]


    total_theme = len(
        theme_df
    )


    if total_theme == 0:
        continue


    strength_rows.append(

        {
            "Theme":
                theme,

            "Reviews":
                total_theme,

            "Positive %":
                round(

                    percentage(

                        (
                            theme_df[
                                "Sentiment"
                            ]
                            == "positive"
                        ).sum(),

                        total_theme,
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
                total_theme,

            "Negative %":
                round(

                    percentage(

                        (
                            theme_df[
                                "Sentiment"
                            ]
                            == "negative"
                        ).sum(),

                        total_theme,
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

    .head(5)
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

    .head(5)
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


# =========================================================
# PRODUCT SUMMARY
# =========================================================

st.markdown(
    """
    <div class="section-title">
        Product Summary
    </div>
    """,
    unsafe_allow_html=True,
)


summary = []


for product in PRODUCTS:

    temp = filtered[
        filtered["Product"]
        == product
    ]


    total = len(
        temp
    )


    if total == 0:
        continue


    summary.append(

        {
            "Product":
                product,

            "Reviews":
                total,

            "Average Rating":
                round(
                    temp["Stars"].mean(),
                    2,
                ),

            "Positive %":
                round(

                    percentage(

                        (
                            temp["Sentiment"]
                            == "positive"
                        ).sum(),

                        total,
                    ),

                    1,
                ),

            "Neutral %":
                round(

                    percentage(

                        (
                            temp["Sentiment"]
                            == "neutral"
                        ).sum(),

                        total,
                    ),

                    1,
                ),

            "Negative %":
                round(

                    percentage(

                        (
                            temp["Sentiment"]
                            == "negative"
                        ).sum(),

                        total,
                    ),

                    1,
                ),
        }
    )


product_summary = (

    pd.DataFrame(
        summary
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


# =========================================================
# FOOTER
# =========================================================

st.caption(

    f"Showing {total_reviews:,} filtered reviews "
    "from the Germany toothbrush review dataset."
)
