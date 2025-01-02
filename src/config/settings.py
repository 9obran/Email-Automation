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

# Dynamic columns and placeholders (will be set at runtime)
REQUIRED_COLUMNS = []
STANDARD_COLUMNS = []
REQUIRED_PLACEHOLDERS = []

# Storage keys for settings
PLACEHOLDER_SETTINGS_KEY = "placeholder_settings"
COLUMN_MAPPING_KEY = "column_mapping"

# Email configuration
DEFAULT_EMAIL_SERVICE = "outlook"
