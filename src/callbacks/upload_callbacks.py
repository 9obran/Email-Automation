from dash import Input, Output
from src.app import app

@app.callback(
    [Output('excel-upload-status', 'children'),
     Output('template-upload-status', 'children')],
    [Input('upload-excel', 'contents'),
     Input('upload-template', 'contents')]
)
def update_upload_status(excel_contents, template_contents):
    """Update upload status for both files"""
    excel_status = "✓ Excel file uploaded successfully" if excel_contents else "No file uploaded"
    template_status = "✓ Template file uploaded successfully" if template_contents else "No file uploaded"
    return excel_status, template_status