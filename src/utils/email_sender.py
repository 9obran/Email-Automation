import win32com.client as win32
import pythoncom
import logging
import os
import re

class EmailSender:
    def __init__(self):
        self.outlook = None
        try:
            pythoncom.CoInitialize()
            self.outlook = win32.Dispatch("Outlook.Application")
        except Exception as e:
            logging.error(f"Failed to initialize Outlook: {str(e)}")
            raise

    def __del__(self):
        try:
            pythoncom.CoUninitialize()
        except:
            pass

    def get_signature(self):
        """Get the default signature from Outlook"""
        try:
            # Get the default signature from the current user's profile
            namespace = self.outlook.GetNamespace("MAPI")
            account = namespace.Accounts[0]  # Get the first email account
            signature = ""
            
            # Try to get signature from AppData
            appdata = os.getenv('APPDATA')
            signature_path = os.path.join(appdata, 'Microsoft', 'Signatures')
            
            if os.path.exists(signature_path):
                # Look for .htm or .rtf files
                signature_files = [f for f in os.listdir(signature_path) 
                                 if f.endswith(('.htm', '.rtf')) and not f.startswith('~')]
                if signature_files:
                    with open(os.path.join(signature_path, signature_files[0]), 'r', encoding='utf-8') as f:
                        signature = f.read()
            
            return signature
        except Exception as e:
            logging.error(f"Failed to get signature: {str(e)}")
            return ""

    def format_email_body(self, body, font_family="Calibri", font_size="11"):
        """Format the email body with HTML"""
        html_body = f"""
        <html>
        <head>
        <style>
            body {{ font-family: {font_family}; font-size: {font_size}pt; }}
        </style>
        </head>
        <body>
        {body.replace('\n', '<br>')}
        <br><br>
        {self.get_signature()}
        </body>
        </html>
        """
        return html_body

    def send_email(self, recipient, subject, body, font_family="Calibri", font_size="11"):
        """Send a single email"""
        try:
            email = self.outlook.CreateItem(0)
            email.Subject = subject
            email.HTMLBody = self.format_email_body(body, font_family, font_size)
            email.Recipients.Add(recipient)
            email.Send()
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

                # Extract the email body based on the "Preview" section
                email_body = self.extract_email_body(email_body)

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

    def extract_email_body(self, template):
        """Extracts the email body from a template with a "Preview" section"""
        preview_match = re.search(r"^\*\*Preview\*\*\n(.*)", template, re.MULTILINE)
        if preview_match:
            return preview_match.group(1).strip()
        else:
            # Handle cases where the "Preview" section is not found or has a different format
            logging.warning("Preview section not found in template. Using the entire template.")
            return template
