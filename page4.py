import requests
import os
import pandas as pd
import plotly.express as px
from dash import html, dcc

# update monthly here: https://www.energy.gov/fecm/articles/natural-gas-imports-and-exports-monthly-2025
url = 'https://www.energy.gov/sites/default/files/2025-08/1.%20U.S.%20Natural%20Gas%20Imports%20Exports%20and%20Re-Exports%20Summary%20%28Jan%202000%20-Jun%202025%29.xlsx'

def download_and_load_file(url, save_dir=None, filename='import_and_exports.xlsx'):
    if save_dir is None:
        save_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(save_dir, exist_ok=True)
    full_path = os.path.join(save_dir, filename)

    response = requests.get(url)
    response.raise_for_status()

    with open(full_path, 'wb') as f:
        f.write(response.content)

    return full_path

def clean_imp_exp_data(file_path, sheet_name="By Country Summary"):
    raw_df = pd.read_excel(file_path, sheet_name=sheet_name, engine="openpyxl")
    raw_df["Transaction Month"] = pd.to_datetime(raw_df["Transaction Month"], errors="coerce")
    raw_df["Year"] = raw_df["Transaction Month"].dt.year
    df_filtered = raw_df[raw_df["Year"] >= 2016].copy()
    country_region_map = {
        # North America
        "Canada": "North America",
        "Mexico": "North America",
        "United States": "North America",

        # Europe
        "United Kingdom": "Europe",
        "France": "Europe",
        "Germany": "Europe",
        "Spain": "Europe",
        "Italy": "Europe",
        "Netherlands": "Europe",
        "Belgium": "Europe",
        "Portugal": "Europe",
        "Greece": "Europe",
        "Turkey": "Europe",
        "Croatia": "Europe",
        "Poland": "Europe",
        "Finland": "Europe",
        "Lithuania": "Europe",
        "Malta": "Europe",

        # Asia
        "Japan": "Asia",
        "South Korea": "Asia",
        "China": "Asia",
        "India": "Asia",
        "Pakistan": "Asia",
        "Thailand": "Asia",
        "Singapore": "Asia",
        "Bangladesh": "Asia",
        "Taiwan": "Asia",
        "Malaysia": "Asia",
        "Indonesia": "Asia",
        "Philippines": "Asia",

        # Middle East
        "United Arab Emirates": "Middle East",
        "Qatar": "Middle East",
        "Saudi Arabia": "Middle East",
        "Kuwait": "Middle East",
        "Israel": "Middle East",
        "Oman": "Middle East",
        "Jordan": "Middle East",

        # Africa
        "South Africa": "Africa",
        "Egypt": "Africa",
        "Morocco": "Africa",
        "Algeria": "Africa",
        "Nigeria": "Africa",
        "Mauritania": "Africa",
        "Senegal": "Africa",

        # South America
        "Brazil": "South America",
        "Argentina": "South America",
        "Chile": "South America",
        "Colombia": "South America",
        "Peru": "South America",

        # Oceania
        "Australia": "Oceania",
        "New Zealand": "Oceania",
    }
    df_filtered.loc[:, "Region"] = df_filtered["Country"].map(country_region_map)
    df_filtered.loc[:, "Region"] = df_filtered["Region"].fillna("RoW")
    return df_filtered

def get_last_12_months_data(df):
    latest_date = df["Transaction Month"].max()
    start_date = latest_date - pd.DateOffset(months=12)
    return df[(df["Transaction Month"] > start_date) & (df["Transaction Month"] <= latest_date)].copy()

def plot_import_export_monthly(df):
    df_last12 = get_last_12_months_data(df)
    df_last12 = df_last12[df_last12["Activity"].isin(["Imports", "Exports"])]
    df_last12["Month"] = df_last12["Transaction Month"].dt.to_period("M").dt.to_timestamp()
    df_grouped = df_last12.groupby(["Month", "Activity"], as_index=False)["Volume (MMCF)"].sum()
    df_grouped["Month"] = pd.Categorical(df_grouped["Month"], ordered=True,
                                         categories=sorted(df_grouped["Month"].unique(), key=lambda x: pd.to_datetime(x)))

    fig = px.bar(
        df_grouped,
        x="Month",
        y="Volume (MMCF)",
        color="Activity",
        barmode="group"
    )
    fig.update_layout(
        template="plotly_white",
        xaxis_title="Month",
        xaxis=dict(
            tickmode='array',
            tickvals=df_grouped["Month"],
            ticktext=df_grouped["Month"].dt.strftime("%b %Y")
        ),
        yaxis_title="Volume (MMCF)",
        legend_title="Activity"
    )
    return fig

def plot_region_volume(df):
    """Bar chart of exports by region for the last 12 months."""
    df_last12 = get_last_12_months_data(df)

    # Filter to exports only
    df_exports = df_last12[df_last12["Activity"] == "Exports"]

    # Group by Region
    df_grouped = df_exports.groupby("Region", as_index=False)["Volume (MMCF)"].sum()

    # Sort regions by volume descending
    df_grouped.sort_values("Volume (MMCF)", ascending=False, inplace=True)

    def format_volume(val):
        if val >= 1_000_000:
            return f"{val / 1_000_000:.2f}M"
        elif val >= 1_000:
            return f"{val / 1_000:.2f}K"
        else:
            return f"{val:.2f}"

    df_grouped["Label"] = df_grouped["Volume (MMCF)"].apply(format_volume)

    fig = px.bar(
        df_grouped,
        x="Region",
        y="Volume (MMCF)",
        text="Label"
    )

    fig.update_traces(
        textposition="outside"
    )

    fig.update_layout(
        template="plotly_white",
        xaxis_title="Region",
        yaxis_title="Export Volume (MMCF)",
        showlegend=False,
        uniformtext_minsize = 8,
        uniformtext_mode = "hide",
        margin = dict(t=80)
    )

    fig.update_yaxes(range=[0, df_grouped["Volume (MMCF)"].max() * 1.15])

    return fig

def plot_us_exports_by_year(df):
    """Bar chart of total U.S. exports by year (2016–2025)."""
    df_exports = df[df["Activity"] == "Exports"].copy()
    df_grouped = df_exports[df_exports["Year"].between(2016, 2025)].groupby("Year", as_index=False)["Volume (MMCF)"].sum()

    fig = px.bar(
        df_grouped,
        x="Year",
        y="Volume (MMCF)",
        text_auto=".2s",
        labels={"Volume (MMCF)": "Export Volume (MMCF)"}
    )
    fig.update_layout(template="plotly_white")
    fig.update_layout(
        xaxis_title="Year",
        yaxis_title="Export Volume (MMCF)",
        xaxis=dict(
            tickmode="linear",  # Ensures all years show up
            tick0=2016,
            dtick=1
        )
    )
    return fig

def exports_eur_asia(df):
    df = df[df["Year"] >= 2021]
    df_exports = df[df["Activity"] == "Exports"]
    df_region = df_exports[df_exports["Region"].isin(["Europe", "Asia"])]
    df_grouped = df_region.groupby([df_region["Transaction Month"], "Region"], as_index=False)["Volume (MMCF)"].sum()

    fig = px.line(
        df_grouped,
        x="Transaction Month",
        y="Volume (MMCF)",
        color="Region"
    )
    fig.update_layout(template="plotly_white", xaxis_title="Date", yaxis_title="Export Volume (MMCF)", showlegend=True)

    return fig

file_path = download_and_load_file(url)
df = clean_imp_exp_data(file_path)

# Generate graphs
fig_monthly = plot_import_export_monthly(df)
fig_region = plot_region_volume(df)
fig_exports_yearly = plot_us_exports_by_year(df)
eur_vs_asia = exports_eur_asia(df)

def get_sources(sources):
    return html.Div([
        html.Hr(),
        html.H4("Sources:", style={"marginTop": "20px"}),
        html.Ul([
            html.Li(html.A(label, href=link, target="_blank"))
            for label, link in sources
        ])
    ], style={"marginTop": "30px", "marginBottom": "20px"})

page4_sources = [
    ("Imports & Exports", "https://www.energy.gov/fecm/articles/natural-gas-imports-and-exports-monthly-2025")
]

layout = html.Div([
    html.H1("U.S. Natural Gas Imports & Exports", style={"textAlign": "center"}),

    html.Div([
        html.H3("Annual U.S. Exports"),
        dcc.Graph(figure=fig_exports_yearly, style={"height": "500px", "overflow": "hidden"})
    ], style={"width": "100%", "padding": "10px"}),

    # Top row: Monthly & Regional graphs
    html.Div([
        html.Div([
            html.H3("Monthly Imports vs Exports (LTM)"),
            dcc.Graph(figure=fig_monthly, style={"height": "500px", "overflow": "hidden"})
        ], style={"width": "32%", "padding": "10px"}),

        html.Div([
            html.H3("U.S. Exports by Region"),
            dcc.Graph(figure=fig_region, style={"height": "500px", "overflow": "hidden"})
        ], style={"width": "32%", "padding": "10px"}),

        html.Div([
            html.H3("U.S. Exports to Europe vs Asia"),
            dcc.Graph(figure=eur_vs_asia, style={"height": "500px", "overflow": "hidden"})  # ← your new figure
        ], style={"width": "32%", "padding": "10px"}),
    ], style={"display": "flex", "flex-direction": "row", "justify-content": "space-between", "flex-wrap": "nowrap"}),

    get_sources(page4_sources)
])

