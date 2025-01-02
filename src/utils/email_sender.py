import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import logging
import os

class EmailSender:
    def __init__(self):
        self.smtp_server = os.getenv('SMTP_SERVER', 'smtp.office365.com')
        self.smtp_port = int(os.getenv('SMTP_PORT', '587'))
        self.email_address = os.getenv('EMAIL_ADDRESS')
        self.email_password = os.getenv('EMAIL_PASSWORD')
        
        if not all([self.email_address, self.email_password]):
            raise ValueError("Email credentials not properly configured")

    def get_signature(self):
        """Get signature from environment variable or return empty string"""
        return os.getenv('EMAIL_SIGNATURE', '')

    def format_email_body(self, body, font_family="Calibri", font_size="11"):
        """Format the email body with HTML"""
        style = 'body { ' + f'font-family: {font_family}; font-size: {font_size}pt;' + ' }'
        html_parts = [
            '<html>',
            '<head>',
            '<style>',
            style,
            '</style>',
            '</head>',
            '<body>',
            body.replace('\n', '<br>'),
            '<br><br>',
            self.get_signature(),
            '</body>',
            '</html>'
        ]
        return '\n'.join(html_parts)

    def send_email(self, recipient, subject, body, font_family="Calibri", font_size="11"):
        """Send a single email"""
        try:
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = self.email_address
            msg['To'] = recipient

            html_content = self.format_email_body(body, font_family, font_size)
            msg.attach(MIMEText(html_content, 'html'))

            with smtplib.SMTP(self.smtp_server, self.smtp_port) as server:
                server.starttls()
                server.login(self.email_address, self.email_password)
                server.send_message(msg)

            logging.info(f"Email sent successfully to {recipient}")
            return True
        except Exception as e:
            logging.error(f"Failed to send email to {recipient}: {str(e)}")
            return False

    def send_batch_emails(self, df, template, subject, placeholder_settings, edited_templates=None, 
                         font_family="Calibri", font_size="11"):
        """Send emails with dynamic placeholder replacement"""
        total_emails = len(df)
        successful = 0
        failed_emails = []
        
        # Get email column (assuming it's named 'email' or similar)
        email_column = next((col for col in df.columns if 'email' in col.lower()), None)
        if not email_column:
            logging.error("No email column found in DataFrame")
            return 0, total_emails, []
        
        for index, row in df.iterrows():
            try:
                email_body = edited_templates.get(str(index), template)
                if "Preview" in email_body:
                    email_body = email_body.split("\n\n", 1)[1]
                
                # Replace all placeholders dynamically
                for column, placeholder in placeholder_settings.items():
                    if placeholder and column in row:
                        email_body = email_body.replace(placeholder, str(row[column]))
                
                if self.send_email(row[email_column], subject, email_body, font_family, font_size):
                    successful += 1
                else:
                    failed_emails.append(row[email_column])
                    
            except Exception as e:
                failed_emails.append(row[email_column])
                logging.error(f"Failed to send email to {row[email_column]}: {str(e)}")
                
        return successful, total_emails, failed_emails
