from dash import Input, Output, State, callback_context
from src.app import app
from src.utils.excel_parser import parse_excel
from src.utils.template_parser import parse_template
from src.utils.email_sender import EmailSender
import logging

@app.callback(
    [Output('send-progress', 'value', allow_duplicate=True),
     Output('progress-status', 'children', allow_duplicate=True),
     Output('send-status', 'children', allow_duplicate=True)],
    [Input('send-btn', 'n_clicks'),
     Input('send-current', 'n_clicks'),
     Input('prev-preview', 'n_clicks'),
     Input('next-preview', 'n_clicks'),
     Input('preview-btn', 'n_clicks')],
    [State('upload-excel', 'contents'),
     State('upload-template', 'contents'),
     State('email-subject', 'value'),
     State('email-service', 'value'),
     State('font-family', 'value'),
     State('font-size', 'value'),
     State('edited-templates', 'data'),
     State('preview-content', 'value'),
     State('preview-index', 'data'),
     State('sent-emails', 'data')],
    prevent_initial_call=True
)
def handle_email_actions(send_all_clicks, send_current_clicks, prev_clicks, 
                        next_clicks, preview_clicks, excel_contents, 
                        template_contents, subject, email_service, 
                        font_family, font_size, edited_templates, 
                        current_content, current_index, sent_emails):
    """Combined callback to handle all email-related actions"""
    ctx = callback_context
    if not ctx.triggered:
        return 0, "", ""
    
    trigger_id = ctx.triggered[0]['prop_id'].split('.')[0]
    
    # Clear messages for navigation actions
    if trigger_id in ['prev-preview', 'next-preview', 'preview-btn']:
        return 0, "", ""
    
    # Handle send current email
    if trigger_id == 'send-current':
        if not send_current_clicks or current_index in sent_emails:
            return 0, "", ""
        
        try:
            df = parse_excel(excel_contents)
            if df is None:
                return 0, "", "Error: Could not load Excel file"
                
            # Remove preview text from email content
            email_content = current_content
            if "Preview" in email_content:
                email_content = email_content.split("\n\n", 1)[1]
                
            email_sender = EmailSender()
            if email_sender.send_email(
                recipient=df.iloc[current_index]['Email'],
                subject=subject,
                body=email_content,
                font_family=font_family,
                font_size=font_size
            ):
                sent_emails.append(current_index)
                return 0, "", f"✓ Email sent successfully to {df.iloc[current_index]['Email']}"
            else:
                return 0, "", "Error: Failed to send email"
            
        except Exception as e:
            logging.error(f"Error sending single email: {str(e)}")
            return 0, "", f"Error sending email: {str(e)}"
    
    # Handle send all emails
    if trigger_id == 'send-btn':
        if not send_all_clicks:
            return 0, "", ""
        
        if not all([excel_contents, template_contents, subject]):
            return 0, "Error: Please provide all required information", ""
        
        try:
            df = parse_excel(excel_contents)
            template = parse_template(template_contents)
            
            if df is None or template is None:
                return 0, "Error: Invalid file format or missing data", ""
                
            email_sender = EmailSender()
            successful, total_emails, failed_emails = email_sender.send_batch_emails(
                df=df,
                template=template,
                subject=subject,
                edited_templates=edited_templates,
                font_family=font_family,
                font_size=font_size
            )
                    
            progress = int((successful / total_emails) * 100)
            status_message = f"✓ Completed: {successful}/{total_emails} emails sent successfully"
            if failed_emails:
                status_message += f"\nFailed recipients: {', '.join(failed_emails)}"
                
            return progress, status_message, ""
        
        except Exception as e:
            logging.error(f"Error in email sending process: {str(e)}")
            return 0, f"Error: {str(e)}", ""
    
    return 0, "", ""