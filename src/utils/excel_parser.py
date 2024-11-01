import base64
import io
import pandas as pd
import openpyxl
import logging
from src.config.settings import REQUIRED_COLUMNS, STANDARD_COLUMNS

def parse_excel(contents):
    """Parse uploaded Excel file"""
    # Check if contents are provided
    if contents is None:
        logging.error("No Excel file provided")
        return None

    try:
        # Split content string to retrieve the actual file data
        content_type, content_string = contents.split(',')
        decoded = base64.b64decode(content_string)
        
        # Load workbook and select the active sheet
        workbook = openpyxl.load_workbook(io.BytesIO(decoded))
        sheet = workbook.active
        
        # Read data and convert it to DataFrame
        data = sheet.values
        cols = next(data)
        df = pd.DataFrame(data, columns=cols)
        
        # Check if DataFrame is empty
        if df.empty:
            logging.error("Excel file is empty")
            return None
            
        # Validate and process DataFrame
        return validate_and_process_dataframe(df)
        
    except Exception as e:
        logging.error(f"Error parsing Excel file: {str(e)}")
        return None

def validate_and_process_dataframe(df):
    """Validate and process the DataFrame"""
    # Convert column names to lowercase to avoid case-sensitivity issues
    df.columns = [str(col).strip().lower() for col in df.columns]
    
    # Ensure all required columns are present
    missing_columns = [col for col in REQUIRED_COLUMNS 
                      if not any(existing.lower() == col for existing in df.columns)]
    if missing_columns:
        logging.error(f"Excel file missing columns: {', '.join(missing_columns)}")
        return None

    # Standardize column names
    column_mapping = {
        col: next(std_col for std_col in STANDARD_COLUMNS 
                 if std_col.lower() == col.lower())
        for col in df.columns if col.lower() in REQUIRED_COLUMNS
    }
    df = df.rename(columns=column_mapping)
    
    # Validate the DataFrame's content
    if not validate_dataframe(df):
        return None
        
    return df

def validate_dataframe(df):
    """Validate DataFrame contents"""
    # Check for missing values in required columns
    if df[STANDARD_COLUMNS].isnull().any().any():
        logging.error("Excel file contains empty cells in required columns")
        return False
        
    # Verify email format in 'Email' column
    if not df["Email"].str.contains("@").all():
        logging.error("Invalid email format detected")
        return False
        
    return True
