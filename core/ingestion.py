"""
Data ingestion module for JV Cost Control System
Handles loading and preprocessing of HR master lists, TFR timesheets, and site attendance logs
"""

import pandas as pd
import pdfplumber
import io
from typing import Union, Dict, Any
import logging

logger = logging.getLogger(__name__)

def normalize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalize column names to standard format across partners
    """
    # Standard mapping for common variations
    column_mapping = {
        # Staff ID variations
        'staff_id': ['staff_id', 'employee_no', 'emp_id', 'staffid', 'employeeid'],
        'name': ['name', 'full_name', 'staff_name', 'employee_name', 'fullname'],
        'role': ['role', 'position', 'job_title', 'designation', 'position_title'],
        'rate': ['rate', 'daily_rate', 'hourly_rate', 'wage', 'salary', 'daily_wage'],
        'date': ['date', 'work_date', 'attendance_date', 'timesheet_date'],
        'hours': ['hours', 'worked_hours', 'hours_worked', 'total_hours'],
        'visa_permit': ['visa_permit', 'work_permit', 'permit_no', 'visa_number', 'permit_number'],
        'nationality': ['nationality', 'nationality_code', 'country'],
        'department': ['department', 'dept', 'division', 'section']
    }
    
    # Create reverse mapping from variations to standard names
    reverse_mapping = {}
    for standard_name, variations in column_mapping.items():
        for variation in variations:
            reverse_mapping[variation.lower()] = standard_name
    
    # Apply normalization
    normalized_columns = {}
    for col in df.columns:
        clean_col = col.strip().lower().replace(' ', '_').replace('-', '_')
        if clean_col in reverse_mapping:
            normalized_columns[col] = reverse_mapping[clean_col]
        else:
            # Keep original if not in mapping
            normalized_columns[col] = clean_col
    
    df.rename(columns=normalized_columns, inplace=True)
    return df

def extract_pdf_table(pdf_path: str, page_numbers: list = None) -> pd.DataFrame:
    """
    Extract table data from PDF using pdfplumber
    """
    all_tables = []
    
    with pdfplumber.open(pdf_path) as pdf:
        pages_to_process = pdf.pages if page_numbers is None else [pdf.pages[i-1] for i in page_numbers if i <= len(pdf.pages)]
        
        for page in pages_to_process:
            tables = page.extract_tables()
            for table in tables:
                if table:  # Only process non-empty tables
                    df = pd.DataFrame(table[1:], columns=table[0])  # First row as header
                    all_tables.append(df)
    
    if all_tables:
        combined_df = pd.concat(all_tables, ignore_index=True)
        return combined_df
    else:
        # If no tables found, try extracting text and parsing as structured data
        logger.warning("No tables found in PDF, attempting text extraction")
        text_data = []
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    # Split by lines and attempt to parse as structured data
                    lines = text.split('\n')
                    for line in lines:
                        # Simple heuristic to detect tabular data (multiple spaces or delimiters)
                        if '  ' in line or '\t' in line:  # Multiple spaces or tabs indicate columns
                            parts = line.split('\t') if '\t' in line else line.split('  ')
                            text_data.append(parts)
        
        if text_data:
            # Find the most common length to determine header row
            lengths = [len(row) for row in text_data]
            if lengths:
                max_len = max(set(lengths), key=lengths.count)  # Most common length
                # Assume first row with max length is header
                header_idx = next((i for i, row in enumerate(text_data) if len(row) == max_len), 0)
                header = text_data[header_idx]
                
                # Remove header from data
                data_rows = [row for i, row in enumerate(text_data) if i != header_idx and len(row) == max_len]
                
                if data_rows:
                    return pd.DataFrame(data_rows, columns=header)
    
    return pd.DataFrame()

def load_partner_data(file_path: Union[str, io.BytesIO], partner: str, source_type: str) -> pd.DataFrame:
    """
    Load and preprocess data from various file formats
    """
    logger.info(f"Loading {source_type} data for {partner}")
    
    # Determine file extension
    if isinstance(file_path, str):
        file_ext = file_path.lower().split('.')[-1]
    else:
        # For uploaded files in Streamlit, we get BytesIO objects
        # We'll need to handle this differently
        if hasattr(file_path, 'name'):
            file_ext = file_path.name.lower().split('.')[-1]
        else:
            raise ValueError("Unknown file type")
    
    try:
        if file_ext in ['csv']:
            df = pd.read_csv(file_path)
        elif file_ext in ['xlsx', 'xls']:
            df = pd.read_excel(file_path)
        elif file_ext in ['pdf']:
            # For PDFs, we need to save to temp file first
            if hasattr(file_path, 'getvalue'):
                import tempfile
                import os
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                    tmp_file.write(file_path.getvalue())
                    tmp_path = tmp_file.name
                
                try:
                    df = extract_pdf_table(tmp_path)
                finally:
                    os.unlink(tmp_path)
            else:
                df = extract_pdf_table(file_path)
        else:
            raise ValueError(f"Unsupported file format: {file_ext}")
        
        # Normalize column names
        df = normalize_column_names(df)
        
        # Add metadata columns
        df['partner'] = partner
        df['source_type'] = source_type
        
        logger.info(f"Loaded {len(df)} rows from {file_path}")
        return df
        
    except Exception as e:
        logger.error(f"Error loading data from {file_path}: {str(e)}")
        raise

def detect_schema_drift(base_df: pd.DataFrame, new_df: pd.DataFrame) -> Dict[str, Any]:
    """
    Detect schema changes between base and new data
    """
    base_cols = set(base_df.columns)
    new_cols = set(new_df.columns)
    
    added_cols = new_cols - base_cols
    removed_cols = base_cols - new_cols
    unchanged_cols = base_cols.intersection(new_cols)
    
    drift_info = {
        'added_columns': list(added_cols),
        'removed_columns': list(removed_cols),
        'unchanged_columns': list(unchanged_cols),
        'has_drift': bool(added_cols or removed_cols)
    }
    
    if drift_info['has_drift']:
        logger.warning(f"Schema drift detected: Added {len(added_cols)}, Removed {len(removed_cols)}")
    
    return drift_info