from dash import Input, Output, State, html
import dash_bootstrap_components as dbc
from src.app import app
from src.utils.excel_parser import parse_excel

@app.callback(
    [Output('excel-upload-status', 'children'),
     Output('template-upload-status', 'children'),
     Output('column-mapping-section', 'style'),
     Output('column-mapping-inputs', 'children'),
     Output('excel-columns', 'data')],
    [Input('upload-excel', 'contents'),
     Input('upload-template', 'contents')]
)
def update_upload_status(excel_contents, template_contents):
    """Update upload status and show column mapping interface"""
    excel_status = "✓ Excel file uploaded successfully" if excel_contents else "No file uploaded"
    template_status = "✓ Template file uploaded successfully" if template_contents else "No file uploaded"
    
    # Handle column mapping section
    if not excel_contents:
        return excel_status, template_status, {'display': 'none'}, [], []
        
    columns = parse_excel(excel_contents)
    if not columns:
        return excel_status, template_status, {'display': 'none'}, [], []
    
    mapping_inputs = []
    for col in columns:
        mapping_inputs.extend([
            dbc.Row([
                dbc.Col([
                    html.Label(f"Column: {col}", className="mt-2"),
                    dbc.Input(
                        id={'type': 'placeholder-input', 'column': col},
                        type="text",
                        placeholder="Enter placeholder character",
                        className="mb-2"
                    )
                ])
            ])
        ])
    
    return excel_status, template_status, {'display': 'block'}, mapping_inputs, columns

@app.callback(
    Output('placeholder-settings', 'data'),
    [Input('save-mapping', 'n_clicks')],
    [State('excel-columns', 'data'),
     State({'type': 'placeholder-input', 'column': ALL}, 'value')]
)
def save_placeholder_mapping(n_clicks, columns, placeholder_values):
    """Save the column-to-placeholder mapping"""
    if not n_clicks or not columns:
        return {}
    
    # Create mapping of columns to their placeholders
    mapping = {
        col: val for col, val in zip(columns, placeholder_values)
        if val  # Only include columns with placeholder values
    }
    
    return mapping
