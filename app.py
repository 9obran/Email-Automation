import os
import sys

# Add the project root directory to Python path
project_root = os.path.dirname(os.path.abspath(__file__))
sys.path.append(project_root)

from src.app import app
server = app.server

if __name__ == '__main__':
    app.run_server(debug=True)
