import logging
from datetime import datetime
import os

# Ensure logs directory exists
os.makedirs('logs', exist_ok=True)

# Logging configuration
LOGGING_CONFIG = {
    'filename': f'logs/email_automation_{datetime.now().strftime("%Y%m%d_%H%M%S")}.log',
    'level': logging.INFO,
    'format': '%(asctime)s - %(levelname)s - %(message)s'
}

# Required Excel columns
REQUIRED_COLUMNS = ["last name", "fund name", "port-co", "email"]
STANDARD_COLUMNS = ["Last Name", "Fund Name", "Port-Co", "Email"]

# Template placeholders
REQUIRED_PLACEHOLDERS = ["X", "Y", "Z"]

# Email configuration
DEFAULT_EMAIL_SERVICE = "outlook"