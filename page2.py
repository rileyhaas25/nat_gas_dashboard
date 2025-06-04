import pandas as pd
from dash import html, dcc, dash_table, Input, Output
import plotly.express as px
from pathlib import Path

data_dir = Path(__file__).resolve().parent

def load_latest_file(keyword: str, ext=".xlsx") -> Path | None:
    files = list(data_dir.glob(f"*{keyword}*{ext}"))
    if not files:
        return None
    return max(files, key=lambda f: f.stat().st_mtime)

def load_pipeline_data() -> pd.DataFrame:
    file_path = load_latest_file("LNG_Production")
    if file_path is None:
        raise FileNotFoundError("No LNG Production Excel file found in the data directory.")
    df = pd.read_excel(file_path)
    df["First Cargo"] = pd.to_datetime(df["First Cargo"], errors="coerce")
    df["Year"] = df["First Cargo"].dt.year
    df = df.drop(columns=["Last Updated"], errors="ignore")
    df = df.dropna(how="all").reset_index(drop=True)
    return df

def us_production_chart(df):
    df_us = df[
        (df["Country"] == "United States") &
        (df["Status"].isin(["Online", "Under Construction"]))
        ].copy()
    yearly_cumulative = (
        df_us.groupby("Year")["MTPA"]
        .sum()
        .sort_index()
        .cumsum()
        .reset_index()
        .rename(columns={"MTPA": "Cumulative MTPA"})
    )

    fig = px.bar(
        yearly_cumulative,
        x=yearly_cumulative["Year"].astype(int).astype(str),
        y="Cumulative MTPA",
        text="Cumulative MTPA",
        title="Total U.S. LNG Production by Year (Online & Under Construction)",
    )

    fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
    fig.update_layout(xaxis_type='category', xaxis_tickformat=',d')
    fig.update_layout(
        template="plotly_white",
        xaxis_title="Year",
        yaxis_title="Capacity (MTPA)",
        uniformtext_minsize=8,
        uniformtext_mode="hide",
    )

    max_y = yearly_cumulative["Cumulative MTPA"].max()
    fig.update_yaxes(range=[0, max_y * 1.1])

    return fig

def qatar_production_chart(df):
    df_qatar = df[
        (df["Country"] == "Qatar") &
        (df["Status"].isin(["Online", "Under Construction"]))
        ].copy()
    yearly_cumulative = (
        df_qatar.groupby("Year")["MTPA"]
        .sum()
        .sort_index()
        .cumsum()
        .reset_index()
        .rename(columns={"MTPA": "Cumulative MTPA"})
    )

    fig = px.bar(
        yearly_cumulative,
        x=yearly_cumulative["Year"].astype(int).astype(str),
        y="Cumulative MTPA",
        text="Cumulative MTPA",
        title="Total Qatar LNG Production by Year (Online & Under Construction)",
    )

    fig.update_traces(texttemplate="%{text:.1f}", textposition="outside")
    fig.update_layout(xaxis_type='category', xaxis_tickformat=',d')
    fig.update_layout(
        template="plotly_white",
        xaxis_title="Year",
        yaxis_title="Capacity (MTPA)",
        uniformtext_minsize=8,
        uniformtext_mode="hide",
    )
    max_y = yearly_cumulative["Cumulative MTPA"].max()
    fig.update_yaxes(range=[0, max_y * 1.1])

    return fig

pipeline_df = load_pipeline_data()
us_graph = us_production_chart(pipeline_df)
qatar_graph = qatar_production_chart(pipeline_df)


# Create dropdown options
status_options = [{"label": s, "value": s} for s in sorted(pipeline_df["Status"].dropna().unique())]
country_options = [{"label": s, "value": s} for s in sorted(pipeline_df["Country"].dropna().unique())]
year_options = [{"label": str(y), "value": str(y)} for y in sorted(pipeline_df["Year"].dropna().unique())]

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
    ("Pipeline Projects", "https://www.respectmyplanet.org/publications/international/rmps-international-lng-map-10th-anniversary-upgrade-with-report")
]

layout = html.Div([
    html.H1("LNG Projects & Capacity", style={"textAlign": "center"}),

    html.Div([
        html.Div([
            html.H2("U.S. LNG Production by Year"),
            dcc.Graph(figure=us_graph)
        ], style={"width": "50%", "padding": "10px"}),

        html.Div([
            html.H2("Qatar LNG Production by Year"),
            dcc.Graph(figure=qatar_graph)
        ], style={"width": "50%", "padding": "10px"})
    ], style={"display": "flex", "flexDirection": "row", "justifyContent": "space-between"}),

    html.H2("LNG Project Tracker"),
    html.Div([
        html.Div([
            html.Label("Filter by Status:"),
            dcc.Dropdown(options=status_options, id="status-filter", multi=True),
        ], style={"marginBottom": "20px"}),

        html.Div([
            html.Label("Filter by Year of First Cargo:"),
            dcc.Dropdown(options=year_options, id="year-filter", multi=True),
        ], style={"marginBottom": "20px"}),

        html.Div([
            html.Label("Filter by Country:"),
            dcc.Dropdown(options=country_options, id="country-filter", multi=True),
        ], style={"marginBottom": "20px"}),
    ], style={"width": "60%", "margin": "auto"}),

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
        Input("country-filter", "value"),
        Input("year-filter", "value")
    )
    def update_table(status, countries, years):
        dff = pipeline_df.copy()
        if status:
            dff = dff[dff["Status"].isin(status)]
        if countries:
            dff = dff[dff["Country"].isin(countries)]
        if years:
            dff = dff[dff["Year"].astype(str).isin(years)]
        return dff.to_dict("records")





