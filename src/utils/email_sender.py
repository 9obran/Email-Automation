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

    def send_batch_emails(self, df, template, subject, edited_templates=None, font_family="Calibri", font_size="11"):
        """Send emails to multiple recipients"""
        total_emails = len(df)
        successful = 0
        failed_emails = []
        
        for index, row in df.iterrows():
            try:
                email_body = edited_templates.get(str(index), template)
                if "Preview" in email_body:
                    email_body = email_body.split("\n\n", 1)[1]
                
                email_body = email_body.replace("X", str(row["Last Name"]))
                email_body = email_body.replace("Y", str(row["Fund Name"]))
                email_body = email_body.replace("Z", str(row["Port-Co"]))
                
                if self.send_email(row['Email'], subject, email_body, font_family, font_size):
                    successful += 1
                else:
                    failed_emails.append(row['Email'])
                    
            except Exception as e:
                failed_emails.append(row['Email'])
                logging.error(f"Failed to send email to {row['Email']}: {str(e)}")
                
        return successful, total_emails, failed_emails
