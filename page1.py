import pandas as pd
import plotly.express as px
from dash import html, dcc, Input, Output, callback, dash_table
from pathlib import Path
import datetime
import pytz
import plotly.graph_objects as go


# Constants
data_dir = Path(__file__).resolve().parent  # Folder where files are uploaded


# Function to load the most recent file containing a specific keyword
def load_latest_file(keyword: str, ext=".csv") -> Path | None:
    files = list(data_dir.glob(f"*{keyword}*{ext}"))
    if not files:
        return None
    return max(files, key=lambda f: f.stat().st_mtime)

def load_latest_excel(keyword: str, ext=".xlsx") -> Path | None:
    files = list(data_dir.glob(f"*{keyword}*{ext}"))
    if not files:
        return None
    return max(files, key=lambda f: f.stat().st_mtime)

# Function to load and clean Henry Hub Excel data
def load_henry_hub() -> pd.DataFrame:
    hh_path = load_latest_file("HenryHub")
    if hh_path is None:
        return pd.DataFrame(columns=["Date", "Henry Hub"])
    df = pd.read_csv(hh_path)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["Henry Hub"] = pd.to_numeric(df["Close"], errors="coerce")
    return df[["Date", "Henry Hub"]].dropna()

# Function to load and clean JKM CSV data
def load_jkm() -> pd.DataFrame:
    jkm_path = load_latest_file("JKM")
    if jkm_path is None:
        return pd.DataFrame(columns=["Date", "JKM"])
    df = pd.read_csv(jkm_path)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["JKM"] = pd.to_numeric(df["Price"], errors="coerce")
    return df[["Date", "JKM"]].dropna()

# Function to load and clean TTF CSV data, converting to USD
def load_ttf() -> pd.DataFrame:
    ttf_path = load_latest_file("TTF_Daily")
    if ttf_path is None:
        return pd.DataFrame(columns=["Date", "TTF (USD)"])
    df = pd.read_csv(ttf_path)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    eur_usd_rate = 1.17
    df["TTF (USD)"] = pd.to_numeric(df["Price"], errors="coerce") * eur_usd_rate / 3.412
    return df[["Date", "TTF (USD)"]].dropna()

def parse_month_label(label: str):
    try:
        # Example input: "July '25"
        parts = label.strip().split()
        month_name = parts[0]
        year_suffix = parts[1].replace("'", "")  # remove apostrophe
        year_full = int("20" + year_suffix)      # e.g. '25 → 2025
        date_obj = pd.to_datetime(f"{month_name} {year_full}", format="%B %Y")
        return date_obj
    except Exception:
        return pd.NaT

def load_ttf_forward() -> pd.DataFrame:
    ttf_for_path = load_latest_excel("TTFCurve1")
    if ttf_for_path is None:
        return pd.DataFrame(columns=["Month", "TTF_Forward_Price"])
    df = pd.read_excel(ttf_for_path, sheet_name="TTF_Curve", header=None, engine="openpyxl")
    date_labels = df.iloc[1, 6:]
    months = date_labels.astype(str).apply(parse_month_label)
    prices = df.iloc[3, 6:]
    ttf_for_df = pd.DataFrame({
        "Month": months,
        "TTF_Forward_Price": prices.values
    })
    ttf_for_df["Date"] = ttf_for_df["Month"]
    return ttf_for_df.reset_index(drop=True)

def load_hh_forward() -> pd.DataFrame:
    ttf_for_path = load_latest_excel("TTFCurve1")
    if ttf_for_path is None:
        return pd.DataFrame(columns=["Month", "HH_Forward_Price"])
    df = pd.read_excel(ttf_for_path, sheet_name="NG_Curve", header=None, engine="openpyxl")
    date_labels = df.iloc[1, 6:]
    months = date_labels.astype(str).apply(parse_month_label)
    prices = df.iloc[3, 6:]
    hh_for_df = pd.DataFrame({
        "Month": months,
        "HH_Forward_Price": prices.values
    })
    hh_for_df["Date"] = hh_for_df["Month"]
    return hh_for_df.reset_index(drop=True)

# Function to merge all daily benchmark data into a wide-format DataFrame
def get_benchmark_prices_daily():
    hh = load_henry_hub()
    jkm = load_jkm()
    ttf = load_ttf()

    df = pd.merge(hh, jkm, on="Date", how="outer")
    df = pd.merge(df, ttf, on="Date", how="outer")
    return df.sort_values("Date").dropna().reset_index(drop=True)

def create_benchmark_price_chart(df):
    fig = px.line(
        df,
        x="Date",
        y=["Henry Hub", "JKM", "TTF (USD)"],
        title="Daily Natural Gas Price Benchmarks (USD/MMBtu)",
        markers=True
    )
    fig.update_layout(
        template="plotly_white",
        legend_title_text="Benchmark",
        yaxis_title= "Price"
    )
    return fig

def create_spot_price_table(df: pd.DataFrame):
    # Get the latest price from the merged benchmark dataframe
    latest = df.sort_values("Date", ascending=False).iloc[0]
    table_data = [
        {"Benchmark": "Henry Hub", "Spot Price": f"${latest['Henry Hub']:.2f}"},
        {"Benchmark": "TTF", "Spot Price": f"${latest['TTF (USD)']:.2f}"},
        {"Benchmark": "JKM", "Spot Price": f"${latest['JKM']:.2f}"}
    ]
    return dash_table.DataTable(
        data=table_data,
        columns=[
            {"name": "Benchmark", "id": "Benchmark"},
            {"name": "Spot Price", "id": "Spot Price"},
        ],
        style_cell={"textAlign": "left", "padding": "10px", "fontSize": 16},
        style_header={"fontWeight": "bold", "backgroundColor": "#f0f0f0"},
        style_table={"marginBottom": "30px"}
    )

def create_ttf_spread_table(df: pd.DataFrame):
    latest = df.sort_values("Date", ascending=False).iloc[0]
    hh = latest["Henry Hub"]
    ttf = latest["TTF (USD)"]
    spread = ttf - hh
    shipping = 0.70
    regas = 0.35
    liquefaction = 2.75
    spread_less_var = ttf - (hh * 1.15) - shipping - regas
    spread_less_all_in = ttf - (hh * 1.15) - shipping - regas - liquefaction
    return dash_table.DataTable(
        columns=[
            {"name": "TTF Spread", "id": "spread"},
            {"name": "Spread less Variable Costs", "id": "variable_cost"},
            {"name": "Spread less All-In Costs", "id": "all_in_cost"}
        ],
        data=[{
            "spread": f"${spread:.2f}",
            "variable_cost": f"${spread_less_var:.2f}",
            "all_in_cost": f"${spread_less_all_in:.2f}"
        }],
        style_cell={"textAlign": "left", "padding": "10px", "fontSize": 16},
        style_header={"fontWeight": "bold", "backgroundColor": "#f0f0f0"},
        style_table={"marginBottom": "30px", "width": "60%"}
    )

def create_jkm_spread_table(df: pd.DataFrame):
    latest = df.sort_values("Date", ascending=False).iloc[0]
    hh = latest["Henry Hub"]
    jkm = latest["JKM"]
    spread = jkm - hh
    shipping = 2.20
    regas = 0.50
    liquefaction = 2.75
    spread_less_var = jkm - (hh * 1.15) - shipping - regas
    spread_less_all_in = jkm - (hh * 1.15) - shipping - regas - liquefaction
    return dash_table.DataTable(
        columns=[
            {"name": "JKM Spread", "id": "spread"},
            {"name": "Spread Less Variable Costs", "id": "variable_cost"},
            {"name": "Spread Less All-In Costs", "id": "all_in_cost"}
        ],
        data=[{
            "spread": f"${spread:.2f}",
            "variable_cost": f"${spread_less_var:.2f}",
            "all_in_cost": f"${spread_less_all_in:.2f}"
        }],
        style_cell={"textAlign": "left", "padding": "10px", "fontSize": 16},
        style_header={"fontWeight": "bold", "backgroundColor": "#f0f0f0"},
        style_table={"marginBottom": "30px", "width": "60%"}
    )

def get_last_modified_time():
    files = []
    for keyword in ["HenryHub", "JKM", "TTF"]:
        for ext in [".csv", ".xlsx"]:
            f = load_latest_file(keyword, ext)
            if f:
                files.append(f)
    if not files:
        return "No files found"
    latest_file = max(files, key=lambda f: f.stat().st_mtime)
    utc_time = datetime.datetime.fromtimestamp(latest_file.stat().st_mtime, tz=datetime.UTC)
    eastern = pytz.timezone("America/New_York")
    local_time = utc_time.astimezone(eastern)
    return f"Last updated: {local_time.strftime('%B %d, %Y at %I:%M %p %Z')}"

def plot_ttf_vs_us_export_costs(df) -> go.Figure:
    fig = go.Figure()
    shipping = 0.70
    regas = 0.35
    liquefaction = 2.75
    df["Variable_Cost"] = (
            df["HH_Forward_Price"] * 1.15 + shipping + regas
    )

    df["All_In_Cost"]  = df["Variable_Cost"] + liquefaction
    print(df.head(20))
    # Line 1: TTF Forwards
    fig.add_trace(go.Scatter(
        x=df["Date"],
        y=df["TTF_Forward_Price"],
        mode="lines",
        name="TTF forwards",
        line=dict(color="blue", width=3)
    ))

    # Line 2: US all-in cost to Europe
    fig.add_trace(go.Scatter(
        x=df["Date"],
        y=df["All_In_Cost"],
        mode="lines",
        name="US all-in cost to Europe",
        line=dict(color="lightblue", dash="dash", width=2)
    ))

    # Line 3: US var cost to Europe
    fig.add_trace(go.Scatter(
        x=df["Date"],
        y=df["Variable_Cost"],
        mode="lines",
        name="US var cost to Europe",
        line=dict(color="red", dash="dash", width=2)
    ))

    # Formatting
    fig.update_layout(
        xaxis_title="Date",
        yaxis=dict(title="$/mmBtu", range=[0, None]),
        template="plotly_white",
        legend=dict(orientation="h", y=-0.2, x=0.05),
        margin=dict(l=60, r=40, t=40, b=60)
    )

    return fig

# Load to preview the result
benchmark_df = get_benchmark_prices_daily()
ttf_forward_df = load_ttf_forward()
hh_forward_df = load_hh_forward()
forward_curves = pd.merge(ttf_forward_df, hh_forward_df,
                          on="Date", how="inner")
print(forward_curves.head(20))
ttf_forward_chart = plot_ttf_vs_us_export_costs(forward_curves)
price_chart = create_benchmark_price_chart(benchmark_df)
time_stamp = get_last_modified_time()
price_table = create_spot_price_table(benchmark_df)
ttf_spread_table = create_ttf_spread_table(benchmark_df)
jkm_spread_table = create_jkm_spread_table(benchmark_df)

def get_sources(sources):
    return html.Div([
        html.Hr(),
        html.H4("Sources:", style={"marginTop": "20px"}),
        html.Ul([
            html.Li(html.A(label, href=link, target="_blank"))
            for label, link in sources
        ])
    ], style={"marginTop": "30px", "marginBottom": "20px"})

page1_sources = [
    ("Henry Hub Data", "https://markets.businessinsider.com/commodities/natural-gas-price"),
    ("JKM Data", "https://www.investing.com/commodities/lng-japan-korea-marker-platts-futures-historical-data"),
    ("TTF Data", "https://www.investing.com/commodities/dutch-ttf-gas-c1-futures-historical-data"),
]

layout = html.Div([
    html.H1("LNG Price Inputs", style={"textAlign": "center", "marginBottom": "10px"}),

    dcc.Interval(id="interval", interval=60 * 1000, n_intervals=0),

    dcc.Graph(id="benchmark-chart", figure=price_chart),

    html.Div(time_stamp, id="last-updated", style={
        "textAlign": "left",
        "marginLeft": "20px",
        "marginTop": "10px",
        "fontStyle": "italic"
    }),

    # Wrapper Div for tables
    html.Div([
        # Section title
        html.H2("Spot Prices and Spreads", style={"textAlign": "center", "marginBottom": "20px"}),

        # First Row: Price Table Centered
        html.Div([
            price_table
        ], style={"display": "flex", "justifyContent": "center", "marginBottom": "40px", "boxShadow": "0 2px 4px rgba(0, 0, 0, 0.1)", "borderRadius": "8px", "padding": "20px", "backgroundColor": "#fafafa"}),

        # Second Row: Spread Tables Side-by-Side
        html.Div([
            html.Div([
                html.H3("TTF vs. Henry Hub Spread Analysis", style={"textAlign": "center"}),
                html.Div(ttf_spread_table, style={"display": "flex", "justifyContent": "center"})
            ], style={"flex": "1", "margin": "0 20px", "boxShadow": "0 2px 4px rgba(0, 0, 0, 0.1)", "borderRadius": "8px", "padding": "20px", "backgroundColor": "#fafafa"}),

            html.Div([
                html.H3("JKM vs. Henry Hub Spread Analysis", style={"textAlign": "center"}),
                html.Div(jkm_spread_table, style={"display": "flex", "justifyContent": "center"})
            ], style={"flex": "1", "margin": "0 20px", "boxShadow": "0 2px 4px rgba(0, 0, 0, 0.1)", "borderRadius": "8px", "padding": "20px", "backgroundColor": "#fafafa"}),
        ], style={"display": "flex", "justifyContent": "center"})
    ]),
    html.Div([
        html.H2("TTF Forwards vs US LNG Export Costs", style={"textAlign": "center", "marginTop": "40px"}),
        dcc.Graph(figure=ttf_forward_chart)
    ], style={"margin": "40px 20px"}),
    get_sources(page1_sources)
])

# Dynamic chart update callback
@callback(
    Output("benchmark-chart", "figure"),
    Output("last-updated", "children"),
    Input("interval", "n_intervals")
)

def update_chart(n_intervals):
    df = get_benchmark_prices_daily()
    fig = create_benchmark_price_chart(df)
    timestamp = get_last_modified_time()
    return fig, timestamp

