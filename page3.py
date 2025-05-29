import requests
import os
import pandas as pd
import plotly.express as px
from dash import html, dcc
from datetime import timedelta
import plotly.graph_objects as go

# update weekly here: https://rigcount.bakerhughes.com/na-rig-count
rig_url = 'https://rigcount.bakerhughes.com/static-files/46b0db4a-61f4-4c78-80bb-a61583b39f35'
# update monthly here under Monthly U.S. dry shale natural gas production by formation: https://www.eia.gov/outlooks/steo/data.php
production_url = "https://www.eia.gov/outlooks/steo/xls/Fig43.xlsx"
FOCUS_BASINS = ["Marcellus", "Haynesville", "Permian", "Eagle Ford", "Utica", "Woodford"]

def download_and_load_rig(url, save_dir=None, filename='baker_hughes_rig_count.xlsx'):
    if save_dir is None:
        save_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(save_dir, exist_ok=True)
    full_path = os.path.join(save_dir, filename)

    response = requests.get(url)
    response.raise_for_status()

    with open(full_path, 'wb') as f:
        f.write(response.content)

    print(f"Downloaded Excel file to: {full_path}")
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

    print(f"Downloaded Excel file to: {full_path}")
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

def prep_data_for_graph(df_all, df_latest):
    df_all["Year"] = df_all["US_PublishDate"].dt.year
    df_all["Month"] = df_all["US_PublishDate"].dt.to_period("M").dt.to_timestamp()

    df_all = df_all[df_all["Basin"].isin(FOCUS_BASINS)]
    df_latest = df_latest[df_latest["Basin"].isin(FOCUS_BASINS)]

    # === Monthly MoM Aggregation ===
    df_monthly = df_all.groupby(["Month", "Basin"], as_index=False)["Rig Count Value"].sum()
    df_monthly["MoM % Change"] = df_monthly.groupby("Basin")["Rig Count Value"].pct_change() * 100

    # Get MoM % Change for current month
    current_month = df_latest["US_PublishDate"].max().to_period("M").to_timestamp()
    mom_current = df_monthly[df_monthly["Month"] == current_month][["Basin", "MoM % Change"]]

    # === YoY Comparison ===
    current_date = df_latest["US_PublishDate"].iloc[0]
    prior_window = (current_date - pd.DateOffset(years=1) - timedelta(days=3),
                    current_date - pd.DateOffset(years=1) + timedelta(days=3))
    df_prior = df_all[(df_all["US_PublishDate"] >= prior_window[0]) & (df_all["US_PublishDate"] <= prior_window[1])]

    df_prior_grouped = df_prior.groupby("Basin")["Rig Count Value"].sum().reset_index().rename(
        columns={"Rig Count Value": "Prior Year Rig Count"})
    df_current_grouped = df_latest.groupby("Basin")["Rig Count Value"].sum().reset_index()

    # Merge with YoY
    df_current_grouped = df_current_grouped.merge(df_prior_grouped, on="Basin", how="left")
    df_current_grouped["YoY % Change"] = (
                                                 (df_current_grouped["Rig Count Value"] - df_current_grouped[
                                                     "Prior Year Rig Count"]) /
                                                 df_current_grouped["Prior Year Rig Count"]
                                         ) * 100

    # Merge with MoM
    df_current_grouped = df_current_grouped.merge(mom_current, on="Basin", how="left")

    # === Year-End Weekly Snapshot for Each Year (last rig count of each year) ===
    df_all = df_all.copy()
    df_all["Week"] = df_all["US_PublishDate"].dt.to_period("W").dt.to_timestamp()

    # Get last publish date per year
    last_week_dates = (
        df_all.groupby(["Year"])["US_PublishDate"]
        .max()
        .reset_index()
        .rename(columns={"US_PublishDate": "LastDate"})
    )

    # Merge to get only rows for the last week of each year
    df_year_end = df_all.merge(last_week_dates, left_on=["Year", "US_PublishDate"], right_on=["Year", "LastDate"])

    # Group by Basin and Year to get year-end snapshot
    df_grouped = (
        df_year_end.groupby(["Year", "Basin"], as_index=False)["Rig Count Value"]
        .sum()
    )

    return df_grouped, df_current_grouped

def clean_production_data(file_path, sheet_name="43"):
    # Step 1: Read from row 27 (index 27) as header
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=27, engine="openpyxl")

    # Step 2: Rename the first column
    df.rename(columns={df.columns[0]: "Date"}, inplace=True)

    # Step 3: Parse date column and filter from 2016+
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Date"].dt.year >= 2016]

    # Step 4: Match actual column names to focus basins
    columns_to_keep = ["Date"] + [col for col in df.columns if any(basin in str(col) for basin in FOCUS_BASINS)]

    df_filtered = df[columns_to_keep]

    return df_filtered

path = download_and_load_rig(rig_url)
df_all = clean_rig_count_data(path)
df_all = filter_columns(df_all)
df_latest = get_most_recent_date(df_all)
df_grouped, df_current_grouped = prep_data_for_graph(df_all, df_latest)
path_2 = download_and_load_production(production_url)
df_production = clean_production_data(path_2)
# Reshape to long format
df_melted = df_production.melt(id_vars="Date", var_name="Basin", value_name="Production (Bcf/d)")

# Filter out zeros and nulls
non_zero = df_melted[df_melted["Production (Bcf/d)"] > 0]

# Get latest date per basin with non-zero production
last_valid = non_zero.groupby("Basin")["Date"].max().reset_index()
last_valid.columns = ["Basin", "LastValidDate"]

# Merge back to filter out dates beyond last valid production
df_trimmed = df_melted.merge(last_valid, on="Basin")
df_trimmed = df_trimmed[df_trimmed["Date"] <= df_trimmed["LastValidDate"]]

# Historical Area Chart
fig_historical = px.area(
    df_grouped,
    x="Year",
    y="Rig Count Value",
    color="Basin",
    labels={"Rig Count Value": "Rig Count", "Year": "Year"},
    template="plotly_white"
)

# Current Week Bar + YoY Line + MoM Line
fig_latest_combo = go.Figure()
fig_latest_combo.add_trace(go.Bar(
    x=df_current_grouped["Basin"],
    y=df_current_grouped["Rig Count Value"],
    name="Current Week Rig Count",
    marker_color="steelblue"
))

# YoY % Change
fig_latest_combo.add_trace(go.Scatter(
    x=df_current_grouped["Basin"],
    y=df_current_grouped["YoY % Change"],
    name="YoY % Change",
    mode="lines+markers",
    yaxis="y2",
    marker=dict(color="firebrick")
))

# MoM % Change
fig_latest_combo.add_trace(go.Scatter(
    x=df_current_grouped["Basin"],
    y=df_current_grouped["MoM % Change"],
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

# Historical Production Chart
fig_production = px.area(
    df_trimmed,
    x="Date",
    y="Production (Bcf/d)",
    color="Basin",
    labels={"Date": "Year", "Production (Bcf/d)": "Production"},
    template="plotly_white"
)


layout = html.Div([
    html.H1("U.S. Natural Gas Rig Count & Production"),

    html.Div([
        html.Div([
            html.H3("Historical Rig Counts by Basin"),
            dcc.Graph(figure=fig_historical)
        ], style={"width": "50%", "padding": "10px"}),

        html.Div([
            html.H3("Current Week Rig Count"),
            dcc.Graph(figure=fig_latest_combo)
        ], style={"width": "50%", "padding": "10px"}),
    ], style={"display": "flex", "flexDirection": "row"}),

    html.Div([
        html.H3("Monthly Dry Shale Gas Production by Basin"),
        dcc.Graph(figure=fig_production)
    ], style={"width": "100%", "padding": "10px"})
])

