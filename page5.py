import os
import requests
import pandas as pd
from dash import html, dcc
import plotly.graph_objects as go
from pathlib import Path

# update monthly under U.S. working natural gas in storage: https://www.eia.gov/outlooks/steo/data.php
storage_url = "https://www.eia.gov/outlooks/steo/xls/Fig27.xlsx"

# update monthly under: https://ec.europa.eu/eurostat/databrowser/view/nrg_stk_gasm__custom_16946737/default/table?lang=en
data_dir = Path(__file__).resolve().parent

def download_storage_excel(url, save_dir=None, filename="monthly_gas_storage.xlsx"):
    if save_dir is None:
        save_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(save_dir, exist_ok=True)

    full_path = os.path.join(save_dir, filename)

    response = requests.get(url)
    response.raise_for_status()

    with open(full_path, "wb") as f:
        f.write(response.content)

    return full_path

def load_latest_file(keyword: str, ext=".csv") -> Path | None:
    files = list(data_dir.glob(f"*{keyword}*{ext}"))
    if not files:
        return None
    return max(files, key=lambda f: f.stat().st_mtime)

def load_eu_storage() -> pd.DataFrame:
    eur_stor_path = load_latest_file("EUR", ext=".xlsx")
    df = pd.read_excel(eur_stor_path, sheet_name="Sheet 1", header=9, skiprows=[10, 11], engine="openpyxl")
    df.rename(columns={df.columns[0]: "Country"}, inplace=True)
    df = df[df["Country"].astype(str).str.match("^[A-Za-z -]+$")]

    # Set index and transpose
    df.set_index("Country", inplace=True)
    df_transposed = df.transpose()
    df_transposed.index.name = "Date"
    df_transposed.index = pd.to_datetime(df_transposed.index, format="%Y-%m", errors="coerce")
    df_transposed = df_transposed.dropna(axis=0, how="all")  # Drop rows where all values are NaN


    # Total monthly storage
    df_transposed = df_transposed.apply(pd.to_numeric, errors="coerce") * 0.0353147
    df_transposed["Total"] = df_transposed.sum(axis=1, skipna=True)
    df_transposed = df_transposed.dropna(subset=["Total"])

    # Build a 5-year rolling average + min/max by month
    monthly = df_transposed["Total"].groupby(df_transposed.index.month)

    avg = monthly.transform("mean")
    high = monthly.transform("max")
    low = monthly.transform("min")

    result = pd.DataFrame({
        "Date": df_transposed.index,
        "Total": df_transposed["Total"],
        "5-Year Avg": avg,
        "5-Year High": high,
        "5-Year Low": low,
    }).dropna().reset_index(drop=True)

    return result

def clean_storage_data(file_path):
    # Load from row 2 (zero-indexed row 1) as header
    df = pd.read_excel(
        file_path,
        sheet_name="27",
        header=27,  # Row 2 in Excel = header row (0-indexed)
        engine="openpyxl"
    )
    df = df.rename(columns={
        "Unnamed: 0": "Date",
        "Level": "Actual Storage",
        "Average": "5-Year Avg",
        "Low": "5-Year Low",
        "High": "5-Year High"
    })
    df = df[["Date", "Actual Storage", "5-Year Avg", "5-Year Low", "5-Year High"]].dropna()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Date"].dt.year >= 2016]
    return df

def create_storage_figure(df):
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=df["Date"],
        y=df["Actual Storage"],
        mode="lines",
        name="Actual Storage",
        line=dict(color="blue"),
        hoverinfo="skip",
        showlegend=True
    ))
    fig.add_trace(go.Scatter(
        x=pd.concat([df["Date"], df["Date"][::-1]]),
        y=pd.concat([df["5-Year High"], df["5-Year Low"][::-1]]),
        fill="toself",
        fillcolor="rgba(200,200,200,0.4)",
        line=dict(color="rgba(255,255,255,0)"),
        name="5-Year Range"
    ))
    fig.add_trace(go.Scatter(
        x=df["Date"],
        y=df["5-Year Avg"],
        mode="lines",
        name="5-Year Avg",
        line=dict(color="green", dash="dash")
    ))
    fig.update_layout(
        xaxis_title="Year",
        yaxis_title="Storage (Bcf)",
        template="plotly_white"
    )
    return fig

def create_eu_storage_chart(df):
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df["Date"], y=df["Total"],
        name="Actual Storage", mode="lines", line=dict(color="blue")
    ))
    # 5-Year Range Band (like U.S. method)
    fig.add_trace(go.Scatter(
        x=pd.concat([df["Date"], df["Date"][::-1]]),
        y=pd.concat([df["5-Year High"], df["5-Year Low"][::-1]]),
        fill="toself",
        fillcolor="rgba(200,200,200,0.4)",
        line=dict(color="rgba(255,255,255,0)"),
        hoverinfo="skip",
        showlegend=True,
        name="5-Year Range"
    ))
    # 5-Year Average
    fig.add_trace(go.Scatter(
        x=df["Date"], y=df["5-Year Avg"],
        name="5-Year Avg", mode="lines", line=dict(color="black", dash="dash")
    ))
    fig.update_layout(
        template="plotly_white",
        yaxis_title="Storage (Bcf)",
        xaxis_title="Year"
    )
    return fig

storage_file_path = download_storage_excel(storage_url)
storage_df = clean_storage_data(storage_file_path)
storage_figure = create_storage_figure(storage_df)
eu_storage_df = load_eu_storage()
eu_storage_fig = create_eu_storage_chart(eu_storage_df)

def get_sources(sources):
    return html.Div([
        html.Hr(),
        html.H4("Sources:", style={"marginTop": "20px"}),
        html.Ul([
            html.Li(html.A(label, href=link, target="_blank"))
            for label, link in sources
        ])
    ], style={"marginTop": "30px", "marginBottom": "20px"})

page5_sources = [
    ("US Nat Gas Storage", "https://www.eia.gov/outlooks/steo/data.php"),
    ("EU Nat Gas Storage", "https://ec.europa.eu/eurostat/databrowser/view/nrg_stk_gasm__custom_16946737/default/table?lang=en")
]

layout = html.Div([
    html.H1("Natural Gas Storage Levels", style={"textAlign": "center"}),
    html.H2("U.S. Natural Gas Storage Data"),
    dcc.Graph(figure=storage_figure),
    html.H2("European Natural Gas Storage Data"),
    dcc.Graph(figure=eu_storage_fig),
    get_sources(page5_sources)
])
