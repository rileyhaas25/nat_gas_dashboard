import requests
import os
import pandas as pd
import plotly.express as px
from dash import html, dcc
from datetime import timedelta
import plotly.graph_objects as go

# update weekly here: https://rigcount.bakerhughes.com/na-rig-count
rig_url = 'https://rigcount.bakerhughes.com/static-files/10ef884d-be25-4591-afdc-d9f3b7ab7c67'
# update monthly here under Monthly U.S. dry shale natural gas production by formation: https://www.eia.gov/outlooks/steo/data.php
production_url = "https://www.eia.gov/outlooks/steo/xls/Fig43.xlsx"
FOCUS_BASINS = ["Marcellus", "Haynesville", "Permian", "Eagle Ford", "Utica", "Woodford"]
BASIN_COLOR_MAP = {
    "Permian": "#636EFA",
    "Haynesville": "#EF553B",
    "Marcellus": "#00CC96",
    "Utica": "#AB63FA",
    "Eagle Ford": "#FFA15A",
    "Woodford": "#19D3F3"
}

def download_and_load_rig(url, save_dir=None, filename='baker_hughes_rig_count.xlsx'):
    if save_dir is None:
        save_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(save_dir, exist_ok=True)
    full_path = os.path.join(save_dir, filename)

    response = requests.get(url)
    response.raise_for_status()

    with open(full_path, 'wb') as f:
        f.write(response.content)

    return full_path

def download_and_load_production(url, save_dir=None, filename='dry_shale_gas_production_by_formation.xlsx'):
    if save_dir is None:
        save_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(save_dir, exist_ok=True)
    full_path = os.path.join(save_dir, filename)

    response = requests.get(url)
    response.raise_for_status()

    with open(full_path, 'wb') as f:
        f.write(response.content)

    return full_path

def clean_rig_count_data(file_path, sheet_name="NAM Weekly"):
    raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine="openpyxl")
    header_row = \
    raw_df[raw_df.apply(lambda row: row.astype(str).str.contains("Date", case=False, na=False)).any(axis=1)].index[0]
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, engine="openpyxl").dropna(how="all").dropna(
        axis=1, how="all")
    df.columns = df.columns.str.strip()
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)

    df = df[df["Country"] == "UNITED STATES"]
    df = df[df["DrillFor"] == "Gas"]
    df["US_PublishDate"] = pd.to_datetime(df["US_PublishDate"], errors="coerce")
    df = df[df["US_PublishDate"].dt.year >= 2016]
    woodford_aliases = ["Ardmore Woodford", "Arkoma Woodford", "Cana Woodford"]
    df["Basin"] = df["Basin"].replace(woodford_aliases, "Woodford")
    return df

def filter_columns(df):
    columns_to_drop = ['County', 'GOM', 'Location', 'State/Province']
    return df.drop(columns=columns_to_drop, errors="ignore")

def get_most_recent_date(df):
    latest_date = df["US_PublishDate"].max()
    return df[df["US_PublishDate"] == latest_date]

def clean_rig_count_yearly(file_path, sheet_name="NAM Yearly"):
    raw_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine="openpyxl")
    header_row = \
    raw_df[raw_df.apply(lambda row: row.astype(str).str.contains("Basin", case=False, na=False)).any(axis=1)].index[0]
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, engine="openpyxl").dropna(how="all").dropna(
        axis=1, how="all")
    df.columns = df.columns.str.strip()
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
    df = df[df["Country"] == "UNITED STATES"]
    df = df[df["DrillFor"] == "Gas"]
    woodford_aliases = ["Ardmore Woodford", "Arkoma Woodford", "Cana Woodford"]
    df["Basin"] = df["Basin"].replace(woodford_aliases, "Woodford")
    df = df[df["Basin"].isin(FOCUS_BASINS)]
    df = df[df["Year"] >= 2016]
    df = df.groupby(["Year", "Basin"], as_index=False)["Rig Count Value"].sum()
    return df

def prep_data_for_graph(df_all, df_latest):
    # Add Year and Month columns
    df_all["Year"] = df_all["US_PublishDate"].dt.year
    df_all["Month"] = df_all["US_PublishDate"].dt.to_period("M").dt.to_timestamp()

    # Filter to focus basins only
    df_all = df_all[df_all["Basin"].isin(FOCUS_BASINS)]
    df_latest = df_latest[df_latest["Basin"].isin(FOCUS_BASINS)]

    # === Monthly MoM Aggregation ===
    df_monthly = df_all.groupby(["Month", "Basin"], as_index=False)["Rig Count Value"].sum()
    df_monthly["MoM % Change"] = df_monthly.groupby("Basin")["Rig Count Value"].pct_change() * 100

    # Extract current month and corresponding MoM % change
    current_month = df_latest["US_PublishDate"].max().to_period("M").to_timestamp()
    mom_current = df_monthly[df_monthly["Month"] == current_month][["Basin", "MoM % Change"]]

    # === YoY Comparison ===
    current_date = df_latest["US_PublishDate"].iloc[0]
    prior_start = current_date - pd.DateOffset(years=1) - timedelta(days=3)
    prior_end = current_date - pd.DateOffset(years=1) + timedelta(days=3)

    df_prior = df_all[(df_all["US_PublishDate"] >= prior_start) & (df_all["US_PublishDate"] <= prior_end)]
    df_prior_grouped = df_prior.groupby("Basin")["Rig Count Value"].sum().reset_index()
    df_prior_grouped.rename(columns={"Rig Count Value": "Prior Year Rig Count"}, inplace=True)

    # Current week aggregation
    df_current_grouped = df_latest.groupby("Basin")["Rig Count Value"].sum().reset_index()

    # Merge YoY % Change
    df_current_grouped = df_current_grouped.merge(df_prior_grouped, on="Basin", how="left")
    df_current_grouped["YoY % Change"] = (
                                                 (df_current_grouped["Rig Count Value"] - df_current_grouped[
                                                     "Prior Year Rig Count"]) /
                                                 df_current_grouped["Prior Year Rig Count"]
                                         ) * 100

    # Merge MoM % Change
    df_current_grouped = df_current_grouped.merge(mom_current, on="Basin", how="left")

    return df_current_grouped

def clean_production_data(file_path, sheet_name="43"):
    # Step 1: Read from row 27 (index 27) as header
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=27, engine="openpyxl")

    df.rename(columns={df.columns[0]: "Date"}, inplace=True)
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Date"].dt.year >= 2016]

    columns_to_keep = ["Date"] + [col for col in df.columns if any(basin in str(col) for basin in FOCUS_BASINS)]

    return df[columns_to_keep]

# === Rig Count Processing ===
rig_file_path = download_and_load_rig(rig_url)
df_rig_all = clean_rig_count_data(rig_file_path)
df_rig_all = filter_columns(df_rig_all)
df_rig_latest = get_most_recent_date(df_rig_all)
df_rig_current_grouped = prep_data_for_graph(df_rig_all, df_rig_latest)
df_yearly = clean_rig_count_yearly(rig_file_path)

# === Production Data Processing ===
prod_file_path = download_and_load_production(production_url)
df_prod_raw = clean_production_data(prod_file_path)

# Reshape to long format and filter invalid values
df_prod_long = df_prod_raw.melt(id_vars="Date", var_name="Basin", value_name="Production (Bcf/d)")
df_prod_long = df_prod_long[df_prod_long["Production (Bcf/d)"] > 0]

# Filter out future values by basin-specific cutoff
last_valid_prod = df_prod_long.groupby("Basin")["Date"].max().reset_index()
last_valid_prod.columns = ["Basin", "LastValidDate"]

df_prod_trimmed = df_prod_long.merge(last_valid_prod, on="Basin")
df_prod_trimmed = df_prod_trimmed[df_prod_trimmed["Date"] <= df_prod_trimmed["LastValidDate"]]

def fig_prod_change(df):
    df = df.sort_values('Date')
    basin_cols = df.columns[1:]
    zero_total_rows = df[df[basin_cols].sum(axis=1) == 0].index
    df = df.drop(index=zero_total_rows).copy()
    latest_date = df['Date'].max()
    prior_date = latest_date - pd.DateOffset(years=1)
    latest_row = df.loc[df['Date'] == df['Date'][df['Date'].sub(latest_date).abs().idxmin()]]
    prior_row = df.loc[df['Date'] == df['Date'][df['Date'].sub(prior_date).abs().idxmin()]]
    change_df = (latest_row.iloc[0, 1:] - prior_row.iloc[0, 1:]).reset_index()
    change_df.columns = ['Basin', 'YoY Change']
    change_df["Label"] = change_df["YoY Change"].apply(lambda x: f"{x:.2f} BCF/d")
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=change_df["Basin"],
        y=change_df["YoY Change"],
        text=change_df["Label"],
        textposition="outside",
        cliponaxis=False,
        marker_color="royalblue"
    ))
    fig.update_layout(
        yaxis_title="Change in Production (BCF/d)",
        yaxis_range=[-0.5, None],
        margin=dict(t=80),
        xaxis_tickangle=-45,
        autosize=True,
        template="plotly_white"
    )
    return fig

def hist_area_chart(df):
    # Historical Area Chart
    fig_historical = px.area(
        df,
        x="Year",
        y="Rig Count Value",
        color="Basin",
        color_discrete_map=BASIN_COLOR_MAP,
        labels={"Rig Count Value": "Rig Count", "Year": "Year"},
        template="plotly_white"
    )
    return fig_historical

def current_week(df):
    # Current Week Bar + YoY Line + MoM Line
    fig_latest_combo = go.Figure()
    fig_latest_combo.add_trace(go.Bar(
        x=df["Basin"],
        y=df["Rig Count Value"],
        name="Current Week Rig Count",
        marker_color="steelblue"
    ))

    # YoY % Change
    fig_latest_combo.add_trace(go.Scatter(
        x=df["Basin"],
        y=df["YoY % Change"],
        name="YoY % Change",
        mode="lines+markers",
        yaxis="y2",
        marker=dict(color="firebrick")
    ))

    # MoM % Change
    fig_latest_combo.add_trace(go.Scatter(
        x=df["Basin"],
        y=df["MoM % Change"],
        name="MoM % Change",
        mode="lines+markers",
        yaxis="y2",
        marker=dict(color="orange")
    ))

    # Layout with Dual Axes
    fig_latest_combo.update_layout(
        xaxis_title="Basin",
        yaxis=dict(title="Rig Count", side="left"),
        yaxis2=dict(title="% Change (YoY & MoM)", overlaying="y", side="right"),
        legend=dict(x=0.01, y=1.15, xanchor="left", yanchor="bottom"),
        template="plotly_white"
    )

    return fig_latest_combo

def historical_production(df):
    # Historical Production Chart
    fig_production = px.area(
        df,
        x="Date",
        y="Production (Bcf/d)",
        color="Basin",
        color_discrete_map=BASIN_COLOR_MAP,
        labels={"Date": "Year", "Production (Bcf/d)": "Production (Bcf/d)"},
        template="plotly_white"
    )
    return fig_production

rig_historical = hist_area_chart(df_yearly)
rig_current_week = current_week(df_rig_current_grouped)
hist_prod_area = historical_production(df_prod_trimmed)
production_change_chart = fig_prod_change(df_prod_raw)

def get_sources(sources):
    return html.Div([
        html.Hr(),
        html.H4("Sources:", style={"marginTop": "20px"}),
        html.Ul([
            html.Li(html.A(label, href=link, target="_blank"))
            for label, link in sources
        ])
    ], style={"marginTop": "30px", "marginBottom": "20px"})

page3_sources = [
    ("Rig Count", "https://rigcount.bakerhughes.com/na-rig-count"),
    ("Nat Gas Production", "https://www.eia.gov/outlooks/steo/data.php")
]

layout = html.Div([
    html.H1("U.S. Natural Gas Rig Count & Production", style={"textAlign": "center", "marginBottom": "10px"}),

    html.Div([
        html.Div([
            html.H3("Historical Rig Counts by Basin"),
            dcc.Graph(figure=rig_historical, style={"height": "500px", "overflow": "hidden"})
        ], style={"width": "50%", "padding": "10px"}),

        html.Div([
            html.H3("Current Week Rig Count"),
            dcc.Graph(figure=rig_current_week, style={"height": "500px", "overflow": "hidden"})
        ], style={"width": "50%", "padding": "10px"}),
    ], style={"display": "flex", "flexDirection": "row"}),

    html.Div([
        html.Div([
            html.H3("Monthly Dry Shale Gas Production by Basin"),
            dcc.Graph(figure=hist_prod_area, style={"height": "500px", "overflow": "hidden"})
        ], style={"width": "50%", "padding": "10px"}),

        html.Div([
            html.H3("Year-over-Year Change in Dry Shale Gas Production by Basin"),
            dcc.Graph(figure=production_change_chart, style={"height": "500px", "overflow": "hidden"})
        ], style={"width": "50%", "padding": "10px"}),
    ], style={"display": "flex", "flexDirection": "row"}),
    get_sources(page3_sources)
])

