//pip install dash dash-bootstrap-components pandas openpyxl

import dash
from dash import dcc, html, Output, Input, State
import dash_bootstrap_components as dbc
import base64
import io

app = dash.Dash(
    __name__, 
    external_stylesheets=[dbc.themes.BOOTSTRAP], 
    suppress_callback_exceptions=True
)

app.layout = dbc.Container([
    html.H2("Upload Two XLSX Files"),
    dbc.Row([
        dbc.Col([
            dcc.Upload(
                id='upload-file-1',
                children=html.Div(['Drag and Drop or ', html.A('Select File 1')]),
                style={'width': '100%', 'height': '60px', 'lineHeight': '60px',
                       'borderWidth': '1px', 'borderStyle': 'dashed',
                       'borderRadius': '5px', 'textAlign': 'center'},
                accept='.xlsx'
            ),
            html.Div(id='output-file-name-1', style={'marginTop': 10})
        ]),
        dbc.Col([
            dcc.Upload(
                id='upload-file-2',
                children=html.Div(['Drag and Drop or ', html.A('Select File 2')]),
                style={'width': '100%', 'height': '60px', 'lineHeight': '60px',
                       'borderWidth': '1px', 'borderStyle': 'dashed',
                       'borderRadius': '5px', 'textAlign': 'center'},
                accept='.xlsx'
            ),
            html.Div(id='output-file-name-2', style={'marginTop': 10})
        ]),
    ]),
    html.Br(),
    html.Div(id='button-container'),
    html.Div(id='print-filenames', style={'color': 'green', 'marginTop': 20})
])

@app.callback(
    Output('output-file-name-1', 'children'),
    Input('upload-file-1', 'filename')
)
def update_filename_1(filename):
    if filename is not None:
        return f"Uploaded File 1: {filename}"

@app.callback(
    Output('output-file-name-2', 'children'),
    Input('upload-file-2', 'filename')
)
def update_filename_2(filename):
    if filename is not None:
        return f"Uploaded File 2: {filename}"

@app.callback(
    Output('button-container', 'children'),
    Input('upload-file-1', 'filename'),
    Input('upload-file-2', 'filename'),
)
def show_button(filename1, filename2):
    if filename1 and filename2:
        return html.Button("Run Function", id='run-function-btn', n_clicks=0)
    return ""

@app.callback(
    Output('print-filenames', 'children'),
    Input('run-function-btn', 'n_clicks'),
    State('upload-file-1', 'filename'),
    State('upload-file-2', 'filename'),
    prevent_initial_call=True
)
def run_function(n_clicks, filename1, filename2):
    if n_clicks > 0:
        # Your custom function would go here
        print(f"File 1: {filename1}")
        print(f"File 2: {filename2}")
        return f"Function run!\nFile 1: {filename1}\nFile 2: {filename2}"

if __name__ == '__main__':
    app.run(debug=True)
