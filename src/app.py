import dash
import dash_bootstrap_components as dbc
from datetime import datetime
import logging
from src.components.layout import create_layout
from src.config.settings import LOGGING_CONFIG

# Configure logging
logging.basicConfig(**LOGGING_CONFIG)

# Initialize the Dash app
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.BOOTSTRAP])

# Set app layout
app.layout = create_layout()

# Import callbacks - must be after layout
from src.callbacks import preview_callbacks
from src.callbacks import upload_callbacks
from src.callbacks import email_callbacks