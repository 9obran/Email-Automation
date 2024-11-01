from dash import html, dcc
import dash_bootstrap_components as dbc

def create_layout():
    return dbc.Container([
        html.H1("Email Automation Tool", className="text-center mb-4", style={'color': '#2C3E50'}),
        create_upload_section(),
        create_email_config_section(),
        create_preview_section(),
        create_action_buttons(),
        create_progress_section(),
        create_store_components()
    ], fluid=True, style={'maxWidth': '1200px'})

def create_upload_section():
    return dbc.Card([
        dbc.CardBody([
            html.H4("Upload Files", className="card-title", style={'color': '#34495E'}),
            dbc.Row([
                dbc.Col([
                    dcc.Upload(
                        id='upload-excel',
                        children=dbc.Card([
                            dbc.CardBody([
                                html.I(className="fas fa-file-excel fa-3x mb-2", style={'color': '#27AE60'}),
                                html.Div('Drag and Drop or Click to Select Excel File')
                            ])
                        ], className="text-center"),
                        className='upload-box mb-3'
                    ),
                    html.Div(id='excel-upload-status', style={'color': '#27AE60'})
                ], width=6),
                dbc.Col([
                    dcc.Upload(
                        id='upload-template',
                        children=dbc.Card([
                            dbc.CardBody([
                                html.I(className="fas fa-file-alt fa-3x mb-2", style={'color': '#2980B9'}),
                                html.Div('Drag and Drop or Click to Select Template File')
                            ])
                        ], className="text-center"),
                        className='upload-box mb-3'
                    ),
                    html.Div(id='template-upload-status', style={'color': '#2980B9'})
                ], width=6),
            ])
        ])
    ], className="mb-4 shadow")

def create_email_config_section():
    return dbc.Card([
        dbc.CardBody([
            html.H4("Email Configuration", className="card-title", style={'color': '#34495E'}),
            dbc.Input(id="email-subject", placeholder="Email Subject", type="text", className="mb-3"),
            dbc.Row([
                dbc.Col([
                    dbc.Select(
                        id="email-service",
                        options=[
                            {"label": "Microsoft Outlook", "value": "outlook"}
                        ],
                        value="outlook",
                        className="mb-3"
                    ),
                ], width=4),
                dbc.Col([
                    dbc.Select(
                        id="font-family",
                        options=[
                            {"label": "Arial", "value": "Arial"},
                            {"label": "Calibri", "value": "Calibri"},
                            {"label": "Times New Roman", "value": "Times New Roman"},
                            {"label": "Verdana", "value": "Verdana"}
                        ],
                        value="Calibri",
                        className="mb-3",
                        placeholder="Select Font"
                    ),
                ], width=4),
                dbc.Col([
                    dbc.Select(
                        id="font-size",
                        options=[
                            {"label": "10", "value": "10"},
                            {"label": "11", "value": "11"},
                            {"label": "12", "value": "12"},
                            {"label": "14", "value": "14"}
                        ],
                        value="11",
                        className="mb-3",
                        placeholder="Select Size"
                    ),
                ], width=4)
            ])
        ])
    ], className="mb-4 shadow")

def create_preview_section():
    return dbc.Card([
        dbc.CardBody([
            html.H4("Email Preview", className="card-title", style={'color': '#34495E'}),
            dbc.Textarea(id="preview-content", className="border p-3 mb-3", style={'height': '300px'}),
            dbc.Row([
                dbc.Col([
                    dbc.Button("Previous", id="prev-preview", color="secondary", className="me-2"),
                    dbc.Button("Next", id="next-preview", color="secondary", className="me-2"),
                    dbc.Button("Send Current Email", id="send-current", color="success", className="me-2"),
                ], className="text-center")
            ])
        ])
    ], className="mb-4 shadow", id="preview-card", style={'display': 'none'})

def create_action_buttons():
    return dbc.Row([
        dbc.Col([
            dbc.Button("Preview Emails", id="preview-btn", color="info", className="me-2", 
                      style={'backgroundColor': '#3498DB'}),
            dbc.Button("Send All Emails", id="send-btn", color="primary",
                      style={'backgroundColor': '#27AE60'}),
        ], className="text-center")
    ], className="mb-4")

def create_progress_section():
    return dbc.Row([
        dbc.Col([
            dcc.Loading(
                id="loading",
                children=[
                    html.Div(id="progress-status", style={'color': '#2C3E50'}),
                    dbc.Progress(id="send-progress", value=0, className="mb-3",
                               style={'height': '20px'}),
                    html.Div(id="send-status", style={'color': '#27AE60'})
                ],
                type="circle"
            )
        ])
    ])

def create_store_components():
    return html.Div([
        dcc.Store(id='preview-index', data=0),
        dcc.Store(id='edited-templates', data={}),
        dcc.Store(id='sent-emails', data=[])
    ])