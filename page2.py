import pandas as pd
from dash import html, dcc, dash_table, Input, Output
import plotly.express as px
from pathlib import Path
import plotly.graph_objects as go

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

def extract_section(df, start_row, end_row, category):
    year_headers = df.iloc[5, 2:19].values
    section = df.iloc[start_row:end_row, :]
    section = section.reset_index(drop=True)

    countries = section.iloc[:, 1].values  # Column B has country names
    data = section.iloc[:, 2:19]             # Columns C onward have values
    data.columns = year_headers            # Set year headers
    data.insert(0, "Country", countries)
    data.insert(0, "Category", category)

    return data

def load_balance_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    file_path = load_latest_file("Global_LNG")
    sheet_name = "Global LNG Balance"
    if file_path is None:
        raise FileNotFoundError("No LNG Balance Excel file found in the data directory.")

    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine="openpyxl")

    africa = ["Nigeria", "Mozambique"]
    asia_pacific = ["Australia", "Malaysia", "Indonesia"]
    df_supply = extract_section(df, start_row=6, end_row=18, category="Supply")
    df_supply["Country"] = df_supply["Country"].replace(
        {c: "Africa" for c in africa} | {c: "Asia-Pacific" for c in asia_pacific}
    )
    df_supply = df_supply.dropna(subset=["Country"])
    df_demand = extract_section(df, start_row=20, end_row=33, category="Demand")
    df_demand = df_demand.dropna(subset=["Country"])
    return df_supply, df_demand

def clean_year_label(val):
    val_str = str(val)
    if val_str.endswith('E'):
        return val_str  # leave forecast years like '2030E' untouched
    try:
        return str(int(float(val)))  # convert 2020.0 → '2020'
    except:
        return val_str  # fallback just in case

def supply_area_chart(df):
    year_cols = [col for col in df.columns if str(col).startswith("20")]
    df_grouped = df.groupby("Country")[year_cols].sum().reset_index()
    df_long = df_grouped.melt(id_vars="Country", var_name="Year", value_name="MTPA")
    df_long["Year"] = df_long["Year"].apply(clean_year_label)
    fig = px.area(
        df_long,
        x="Year",
        y="MTPA",
        color="Country",
        labels={"MTPA": "Supply (MTPA)", "Country": "Region/Country"}
    )
    fig.update_layout(template="plotly_white", xaxis_title="Year", yaxis_title="Cumulative Supply", xaxis_type="category")

    return fig

def demand_area_chart(df):
    year_cols = [col for col in df.columns if isinstance(col, (int, float)) or str(col).endswith("E")]
    asia_row = df[df["Country"] == "Asia"].copy()
    china_row = df[df["Country"] == "Mainland China"].copy()
    asia_ex_china = asia_row.copy()
    for col in year_cols:
        asia_ex_china[col] = asia_row[col].values - china_row[col].values
    asia_ex_china["Country"] = "Asia (ex-China)"
    to_drop = ["Japan", "South Korea", "India", "Taiwan", "Pak-Ban", "SE Asia", "Asia"]
    df = df[~df["Country"].isin(to_drop)]
    df = pd.concat([df, asia_ex_china], ignore_index=True)
    df_long = df.melt(id_vars=["Country"], value_vars=year_cols,
                             var_name="Year", value_name="MTPA")

    # Convert year column to string/int for clean x-axis
    df_long["Year"] = df_long["Year"].apply(clean_year_label)

    fig = px.area(
        df_long,
        x="Year",
        y="MTPA",
        color="Country",
        labels={"MTPA": "Demand (MTPA)", "Country": "Region/Country"}
    )
    fig.update_layout(template="plotly_white", xaxis_title="Year", yaxis_title="Cumulative Demand", xaxis_type="category")

    return fig

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
    yearly_cumulative["Cumulative Bcf/d"] = yearly_cumulative["Cumulative MTPA"] * 0.131584156

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=yearly_cumulative["Year"].astype(int).astype(str),
        y=yearly_cumulative["Cumulative MTPA"],
        marker_color="royalblue",
        text=yearly_cumulative["Cumulative MTPA"],
        textposition="outside",
        texttemplate="%{text:.1f}",
        cliponaxis=True,
        name = "",
        yaxis="y"
    ))
    # Invisible bars on secondary y-axis (just to activate it)
    fig.add_trace(go.Bar(
        x=yearly_cumulative["Year"].astype(int).astype(str),
        y=yearly_cumulative["Cumulative Bcf/d"],
        marker_color="rgba(0,0,0,0)",  # Fully transparent
        name="",  # No legend entry
        showlegend=False,
        hoverinfo="skip",
        yaxis="y2"
    ))

    max_mtpa = yearly_cumulative["Cumulative MTPA"].max()
    max_bcf_d = yearly_cumulative["Cumulative Bcf/d"].max()

    fig.update_layout(
        xaxis=dict(title="Year"),
        yaxis=dict(
            title="Cumulative MTPA",
            side="left",
            range=[0, max_mtpa * 1.1],
            showgrid=False
        ),
        yaxis2=dict(
            title="Cumulative Bcf/d",
            side="right",
            overlaying="y",
            range=[0, max_bcf_d * 1.1],
            showgrid=False,
            showticklabels=True,
            showline=True,
            tickfont=dict(color="black")
        ),
        template="plotly_white",
        uniformtext_minsize=8,
        uniformtext_mode="hide",
        showlegend=False
    )

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
    yearly_cumulative["Cumulative Bcf/d"] = yearly_cumulative["Cumulative MTPA"] * 0.131584156

    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=yearly_cumulative["Year"].astype(int).astype(str),
        y=yearly_cumulative["Cumulative MTPA"],
        marker_color="royalblue",
        text=yearly_cumulative["Cumulative MTPA"],
        textposition="outside",
        texttemplate="%{text:.1f}",
        cliponaxis=True,
        name="",
        yaxis="y"
    ))
    # Invisible bars on secondary y-axis (just to activate it)
    fig.add_trace(go.Bar(
        x=yearly_cumulative["Year"].astype(int).astype(str),
        y=yearly_cumulative["Cumulative Bcf/d"],
        marker_color="rgba(0,0,0,0)",  # Fully transparent
        name="",  # No legend entry
        showlegend=False,
        hoverinfo="skip",
        yaxis="y2"
    ))

    max_mtpa = yearly_cumulative["Cumulative MTPA"].max()
    max_bcf_d = yearly_cumulative["Cumulative Bcf/d"].max()

    fig.update_layout(
        xaxis=dict(title="Year"),
        yaxis=dict(
            title="Cumulative MTPA",
            side="left",
            range=[0, max_mtpa * 1.1],
            showgrid=False
        ),
        yaxis2=dict(
            title="Cumulative Bcf/d",
            side="right",
            overlaying="y",
            range=[0, max_bcf_d * 1.1],
            showgrid=False,
            showticklabels=True,
            showline=True,
            tickfont=dict(color="black")
        ),
        template="plotly_white",
        uniformtext_minsize=8,
        uniformtext_mode="hide",
        showlegend=False
    )

    return fig

pipeline_df = load_pipeline_data()
us_graph = us_production_chart(pipeline_df)
qatar_graph = qatar_production_chart(pipeline_df)

df_supply, df_demand = load_balance_data()
lng_supply = supply_area_chart(df_supply)
lng_demand = demand_area_chart(df_demand)

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
    ("Pipeline Projects", "https://www.respectmyplanet.org/publications/international/rmps-international-lng-map-10th-anniversary-upgrade-with-report"),
    ("Global LNG Supply & Demand", "https://marquee.gs.com/content/research/authors/15b3f07d-5914-4e9c-80ad-26811164a1c5.html")
]

layout = html.Div([
    html.H1("LNG Projects & Capacity", style={"textAlign": "center"}),

    html.Div([
        html.Div([
            html.H3("U.S. LNG Production by Year (Online & Under Construction)"),
            dcc.Graph(figure=us_graph, style={"height": "500px", "overflow": "hidden"})
        ], style={"width": "50%", "padding": "10px"}),

        html.Div([
            html.H3("Qatar LNG Production by Year (Online & Under Construction)"),
            dcc.Graph(figure=qatar_graph, style={"height": "500px", "overflow": "hidden"})
        ], style={"width": "50%", "padding": "10px"})
    ], style={"display": "flex", "flexDirection": "row", "justifyContent": "space-between"}),

    html.H2("LNG Project Tracker", style={"textAlign": "center"}),
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
    html.Div([
            html.H2("Global LNG Supply & Demand", style={"textAlign": "center"}),

            html.Div([
                html.Div([
                    html.H3("Global LNG Supply by Country/Region"),
                    dcc.Graph(figure=lng_supply, style={"height": "500px", "overflow": "hidden"})
                ], style={"width": "50%", "padding": "10px"}),

                html.Div([
                    html.H3("Global LNG Demand by Region"),
                    dcc.Graph(figure=lng_demand, style={"height": "500px", "overflow": "hidden"})
                ], style={"width": "50%", "padding": "10px"})
            ], style={"display": "flex", "flexDirection": "row", "justifyContent": "space-between"})
    ]),
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





