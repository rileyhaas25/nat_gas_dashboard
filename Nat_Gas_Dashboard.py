
from dash import Dash, dcc, html, Input, Output
import page3, page1, page2  # These are the converted pages from your original files

app = Dash(__name__, suppress_callback_exceptions=True)
server = app.server

app.layout = html.Div([
    dcc.Location(id='url', refresh=False),
    html.Div([
        dcc.Link('Imports/Exports', href='/'),
        dcc.Link('Pipelines/Storage', href='/pipelines'),
        dcc.Link('Rig Activity/Production', href='/rigs'),
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
    else:
        return html.H1("404 - Page Not Found")

if __name__ == '__main__':
    app.run(debug=True)
