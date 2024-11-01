from dash import Input, Output, State, callback_context
from src.app import app
from src.utils.excel_parser import parse_excel
from src.utils.template_parser import parse_template
import logging

@app.callback(
    [Output('preview-card', 'style'),
     Output('preview-content', 'value'),
     Output('preview-index', 'data')],
    [Input('preview-btn', 'n_clicks'),
     Input('prev-preview', 'n_clicks'),
     Input('next-preview', 'n_clicks')],
    [State('upload-excel', 'contents'),
     State('upload-template', 'contents'),
     State('preview-index', 'data'),
     State('edited-templates', 'data'),
     State('preview-content', 'value')]
)
def handle_preview(preview_clicks, prev_clicks, next_clicks, excel_contents, 
                  template_contents, current_index, edited_templates, current_content):
    if not preview_clicks:
        return {'display': 'none'}, "", 0
    
    if not excel_contents or not template_contents:
        return {'display': 'block'}, "Error: Please upload both Excel and template files", 0
        
    df = parse_excel(excel_contents)
    template = parse_template(template_contents)
    
    if df is None:
        return {'display': 'block'}, "Error: Could not load Excel file. Please check format and required columns", 0
    if template is None:
        return {'display': 'block'}, "Error: Could not load template file. Please check format and placeholders", 0
        
    ctx = callback_context
    button_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    if button_id == 'prev-preview':
        current_index = max(0, current_index - 1)
    elif button_id == 'next-preview':
        if current_content:
            # Remove the preview header if it exists
            content_parts = current_content.split("\n\n", 1)
            if len(content_parts) > 1:
                edited_templates[str(current_index)] = content_parts[1]
            else:
                edited_templates[str(current_index)] = current_content
        current_index = min(len(df) - 1, current_index + 1)
    
    try:
        # Get the email content
        email_content = edited_templates.get(str(current_index), template)
        email_content = email_content.replace("X", str(df.iloc[current_index]["Last Name"]))
        email_content = email_content.replace("Y", str(df.iloc[current_index]["Fund Name"]))
        email_content = email_content.replace("Z", str(df.iloc[current_index]["Port-Co"]))
        
        # Add preview header
        preview_header = f"Preview {current_index + 1} of {len(df)}"
        preview_text = f"{preview_header}\n\n{email_content}"
        
    except Exception as e:
        logging.error(f"Error personalizing preview: {str(e)}")
        return {'display': 'block'}, "Error: Could not generate preview", current_index
        
    return {'display': 'block'}, preview_text, current_index

@app.callback(
    [Output('send-status', 'children'),
     Output('progress-status', 'children')],
    [Input('prev-preview', 'n_clicks'),
     Input('next-preview', 'n_clicks'),
     Input('preview-btn', 'n_clicks')]
)
def clear_status_messages(prev_clicks, next_clicks, preview_clicks):
    """Clear status messages when navigating previews"""
    return "", ""