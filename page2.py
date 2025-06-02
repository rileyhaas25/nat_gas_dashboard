import os
import requests
import pandas as pd
from dash import html, dcc, dash_table, Input, Output
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path

# update quarterly here under Pipeline Projects: https://www.eia.gov/naturalgas/data.php#imports
pipeline_url = "https://www.eia.gov/naturalgas/pipelines/EIA-NaturalGasPipelineProjects_Apr2025.xlsx"

# update monthly under U.S. working natural gas in storage: https://www.eia.gov/outlooks/steo/data.php
storage_url = "https://www.eia.gov/outlooks/steo/xls/Fig27.xlsx"

# update monthly under: https://ec.europa.eu/eurostat/databrowser/view/nrg_stk_gasm__custom_16946737/default/table?lang=en
data_dir = Path(__file__).resolve().parent

def download_pipeline_excel(url, save_dir=None, filename="pipeline_projects.xlsx"):
    if save_dir is None:
        save_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(save_dir, exist_ok=True)

    full_path = os.path.join(save_dir, filename)

    response = requests.get(url)
    response.raise_for_status()

    with open(full_path, "wb") as f:
        f.write(response.content)

    print(f"Downloaded to: {full_path}")
    return full_path

def download_storage_excel(url, save_dir=None, filename="monthly_gas_storage.xlsx"):
    if save_dir is None:
        save_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(save_dir, exist_ok=True)

    full_path = os.path.join(save_dir, filename)

    response = requests.get(url)
    response.raise_for_status()

    with open(full_path, "wb") as f:
        f.write(response.content)

    print(f"Downloaded to: {full_path}")
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


    print("Processed Storage:")
    print(result.tail())
    return result

def clean_pipeline_data(file_path):
    # Load from row 2 (zero-indexed row 1) as header
    df = pd.read_excel(
        file_path,
        sheet_name="Natural Gas Pipeline Projects",
        header=1,  # Row 2 in Excel = header row (0-indexed)
        engine="openpyxl"
    )
    # Only keep the specified columns
    keep_cols = [
        "Last Updated Date", "Project Name", "Pipeline Operator Name", "Project Type",
        "Status", "Year In Service Date", "State(s)", "Additional Capacity (MMcf/d)", "Notes", "Demand Served"
    ]
    df = df[keep_cols]
    # Drop fully empty rows and reset index
    df.dropna(how="all", inplace=True)
    df.reset_index(drop=True, inplace=True)

    # Convert dates
    df["Last Updated Date"] = pd.to_datetime(df["Last Updated Date"], errors="coerce")
    df["Last Updated Date"] = df["Last Updated Date"].dt.strftime("%-m/%-d/%Y")
    df["Year In Service Date"] = pd.to_numeric(df["Year In Service Date"], errors="coerce").astype("Int64")

    # Filter to keep only rows where 'Demand Served' contains 'LNG' (case-insensitive)
    df = df[df["Demand Served"].astype(str).str.contains("LNG", case=False, na=False)]

    return df

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

def create_capacity_by_year_chart(df):
    cap_df = df.dropna(subset=["Year In Service Date", "Additional Capacity (MMcf/d)"])
    grouped = cap_df.groupby("Year In Service Date")["Additional Capacity (MMcf/d)"].sum().reset_index()

    fig = px.bar(
        grouped,
        x="Year In Service Date",
        y="Additional Capacity (MMcf/d)",
        labels={
            "Year In Service Date": "Year",
            "Additional Capacity (MMcf/d)": "Capacity (MMcf/d)"
        },
        title="Planned Pipeline Capacity Additions by Year",
        text_auto=".2s",
        template="plotly_white"
    )
    return fig

def create_storage_figure(df):
    fig = go.Figure()
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
    fig.add_trace(go.Scatter(
        x=df["Date"],
        y=df["Actual Storage"],
        mode="lines",
        name="Actual Storage",
        line=dict(color="blue")
    ))
    fig.add_trace(go.Scatter(
        x=df["Date"],
        y=df["5-Year Avg"],
        mode="lines",
        name="5-Year Avg",
        line=dict(color="green", dash="dash")
    ))
    fig.update_layout(
        title="U.S. Natural Gas Storage vs. 5-Year Range",
        xaxis_title="Year",
        yaxis_title="Storage (Bcf)",
        template="plotly_white",
        legend=dict(x=0.01, y=0.99)
    )
    return fig

def create_eu_storage_chart(df):
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df["Date"], y=df["Total"],
        name="European Storage", mode="lines", line=dict(color="blue")
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
        title="European Natural Gas Storage vs. 5-Year Range",
        template="plotly_white",
        yaxis_title="Storage (Bcf)",
        xaxis_title="Year"
    )
    return fig

pipeline_file_path = download_pipeline_excel(pipeline_url)
pipeline_df = clean_pipeline_data(pipeline_file_path)
storage_file_path = download_storage_excel(storage_url)
storage_df = clean_storage_data(storage_file_path)
storage_figure = create_storage_figure(storage_df)
capacity_chart = create_capacity_by_year_chart(pipeline_df)
eu_storage_df = load_eu_storage()
eu_storage_fig = create_eu_storage_chart(eu_storage_df)

# Create dropdown options
status_options = [{"label": s, "value": s} for s in sorted(pipeline_df["Status"].dropna().unique())]
state_options = [{"label": s, "value": s} for s in sorted(pipeline_df["State(s)"].dropna().unique())]
year_options = [{"label": str(y), "value": str(y)} for y in sorted(pipeline_df["Year In Service Date"].dropna().unique())]
type_options = [{"label": t, "value": t} for t in sorted(pipeline_df["Project Type"].dropna().unique())]



layout = html.Div([
    html.H1("US Natural Gas Pipeline Projects and Storage"),

    html.H2("Pipeline Capacity Additions"),
    dcc.Graph(figure=capacity_chart),

    html.H2("US Pipeline Projects"),
    html.Div([
        html.Label("Filter by Status:"),
        dcc.Dropdown(options=status_options, id="status-filter", multi=True),

        html.Label("Filter by State(s):"),
        dcc.Dropdown(options=state_options, id="state-filter", multi=True),

        html.Label("Filter by Year In Service:"),
        dcc.Dropdown(options=year_options, id="year-filter", multi=True),

        html.Label("Filter by Project Type:"),
        dcc.Dropdown(options=type_options, id="type-filter", multi=True),
    ], style={"columnCount": 2, "marginBottom": "20px"}),

    dash_table.DataTable(
        id="pipeline-table",
        columns=[{"name": col, "id": col} for col in pipeline_df.columns],
        data=pipeline_df.to_dict("records"),
        page_action="none",
        style_table={"overflowY": "auto", "maxHeight": "800px"},
        fixed_rows={"headers": True},
        style_cell={"textAlign": "left", "whiteSpace": "normal", "minWidth": "120px"},
        filter_action="native",
        sort_action="native"
    ),

    html.H2("U.S. Natural Gas Storage Data"),
    dcc.Graph(figure=storage_figure),

    html.H2("European Natural Gas Storage Data"),
    dcc.Graph(figure=eu_storage_fig)
])

def register_callbacks(app):
    @app.callback(
        Output("pipeline-table", "data"),
        Input("status-filter", "value"),
        Input("state-filter", "value"),
        Input("year-filter", "value"),
        Input("type-filter", "value")
    )
    def update_table(status, states, years, types):
        dff = pipeline_df.copy()
        if status:
            dff = dff[dff["Status"].isin(status)]
        if states:
            dff = dff[dff["State(s)"].isin(states)]
        if years:
            dff = dff[dff["Year In Service Date"].astype(str).isin(years)]
        if types:
            dff = dff[dff["Project Type"].isin(types)]
        return dff.to_dict("records")





