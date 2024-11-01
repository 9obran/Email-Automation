import base64
import logging
from src.config.settings import REQUIRED_PLACEHOLDERS

def parse_template(contents):
    """Parse uploaded template file"""
    # Check if contents are provided
    if contents is None:
        logging.error("No template file provided")
        return None
        
    try:
        # Decode the template file content
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        template = decoded.decode('utf-8')
        
        # Check if template is empty
        if not template.strip():
            logging.error("Template file is empty")
            return None
            
        # Validate the template's placeholders
        if not validate_template(template):
            return None
            
        return template
        
    except Exception as e:
        logging.error(f"Error parsing template file: {str(e)}")
        return None

def validate_template(template):
    """Validate template contents"""
    # Ensure all required placeholders are present
    if not all(placeholder in template for placeholder in REQUIRED_PLACEHOLDERS):
        logging.error("Template missing required placeholders")
        return False
    return True
