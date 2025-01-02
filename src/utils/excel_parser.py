import base64
import io
import pandas as pd
import openpyxl
import logging

def parse_excel(contents, required_columns=None):
    """Parse uploaded Excel file with dynamic column validation"""
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
            
        # If this is the first load (no required columns specified)
        if not required_columns:
            # Return just the column names for mapping
            return list(df.columns)
        
        # Validate and process DataFrame with required columns
        return validate_and_process_dataframe(df, required_columns)
        
    except Exception as e:
        logging.error(f"Error parsing Excel file: {str(e)}")
        return None

def validate_and_process_dataframe(df, column_mapping):
    """Validate and process the DataFrame with dynamic column mapping"""
    try:
        # Convert column names to lowercase for comparison
        df.columns = [str(col).strip().lower() for col in df.columns]
        
        # Ensure all mapped columns are present
        required_columns = [col.lower() for col in column_mapping.keys()]
        missing_columns = [col for col in required_columns 
                         if not any(existing.lower() == col 
                                  for existing in df.columns)]
        if missing_columns:
            logging.error(f"Excel file missing columns: {', '.join(missing_columns)}")
            return None

        # Rename columns according to mapping
        df = df.rename(columns=column_mapping)
        
        # Validate the DataFrame's content
        if not validate_dataframe(df, list(column_mapping.values())):
            return None
            
        return df
        
    except Exception as e:
        logging.error(f"Error processing DataFrame: {str(e)}")
        return None

def validate_dataframe(df, required_columns):
    """Validate DataFrame contents with dynamic required columns"""
    # Check for missing values in required columns
    if df[required_columns].isnull().any().any():
        logging.error("Excel file contains empty cells in required columns")
        return False
        
    # Ensure there's at least one column designated as email
    email_columns = [col for col in required_columns if 'email' in col.lower()]
    if not email_columns:
        logging.error("No email column specified in mapping")
        return False
        
    # Verify email format in email column(s)
    for email_col in email_columns:
        if not df[email_col].str.contains("@").all():
            logging.error(f"Invalid email format detected in column {email_col}")
            return False
            
    return True
