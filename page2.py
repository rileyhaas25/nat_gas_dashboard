import os
import requests
import pandas as pd
from dash import html, dcc, dash_table, Input, Output
import plotly.express as px

# update quarterly here under Pipeline Projects: https://www.eia.gov/naturalgas/data.php#imports
pipeline_url = "https://www.eia.gov/naturalgas/pipelines/EIA-NaturalGasPipelineProjects_Apr2025.xlsx"

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


pipeline_file_path = download_pipeline_excel(pipeline_url)
pipeline_df = clean_pipeline_data(pipeline_file_path)
capacity_chart = create_capacity_by_year_chart(pipeline_df)


# Create dropdown options
status_options = [{"label": s, "value": s} for s in sorted(pipeline_df["Status"].dropna().unique())]
state_options = [{"label": s, "value": s} for s in sorted(pipeline_df["State(s)"].dropna().unique())]
year_options = [{"label": str(y), "value": str(y)} for y in sorted(pipeline_df["Year In Service Date"].dropna().unique())]
type_options = [{"label": t, "value": t} for t in sorted(pipeline_df["Project Type"].dropna().unique())]

def get_sources(sources):
    return html.Div([
        html.Hr(),
        html.H4("Sources:", style={"marginTop": "20px"}),
        html.Ul([
            html.Li(html.A(label, href=link, target="_blank"))
            for label, link in sources
        ])
    ], style={"marginTop": "30px", "marginBottom": "20px"})

page2_sources = [
    ("Pipeline Projects", "https://www.eia.gov/naturalgas/data.php")
]

layout = html.Div([
    html.H1("U.S. Natural Gas Pipeline Projects", style={"textAlign": "center"}),

    html.H2("Pipeline Capacity Additions"),
    dcc.Graph(figure=capacity_chart),

    html.H2("U.S. Pipeline Projects"),
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
    get_sources(page2_sources)
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





