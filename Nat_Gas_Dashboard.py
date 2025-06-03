from dash import Dash, dcc, html, Input, Output
import os
import page1
import page2
import page3
import page4
import page5

app = Dash(__name__, suppress_callback_exceptions=True)
server = app.server

app.layout = html.Div([
    dcc.Location(id='url', refresh=False),
    html.Div([
        dcc.Link('Pricing', href='/'),
        dcc.Link('LNG Related Pipeline Capacity', href='/pipelines'),
        dcc.Link('Rig Activity/Production', href='/rigs'),
        dcc.Link('Imports/Exports', href='/lng'),
        dcc.Link('Storage', href='/storage'),
    ], style={'margin': '20px', 'display': 'flex', 'gap': '20px'}),
    html.Div(id='page-content')
])

@app.callback(Output('page-content', 'children'), Input('url', 'pathname'))
def display_page(pathname):
    if pathname == '/':
        return page1.layout
    elif pathname == '/pipelines':
        return page2.layout
    elif pathname == '/rigs':
        return page3.layout
    elif pathname == '/lng':
        return page4.layout
    elif pathname == '/storage':
        return page5.layout
    else:
        return html.H1("404 - Page Not Found")

if hasattr(page2, "register_callbacks"):
    page2.register_callbacks(app)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)
