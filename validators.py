import pandas as pd
from typing import List, Tuple, Dict, Any

def clean_headers(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standardizes headers. 
    Equivalent to: public DataFrame cleanHeaders(DataFrame df)
    """
    # Annotating the list generation
    clean_columns: List[str] = [str(col).strip().lower() for col in df.columns]
    df.columns = clean_columns # type: ignore

    # Explicitly typed Dictionary
    header_mapping: Dict[str, str] = {
        'full name': 'name',
        'customer name': 'name',
        'customer': 'name',
        'email address': 'email',
        'e-mail address': 'email',
        'cell': 'phone',
        'phone number': 'phone',
        'dob': 'date of birth'
    }
    
    return df.rename(columns=header_mapping)

def validate_names(df: pd.DataFrame, file_name: str, sheet_name: str) -> Tuple[pd.DataFrame, List[pd.DataFrame]]:
    """Identifies and extracts rows with missing names."""
    error_list: List[pd.DataFrame] = []
    
    if 'name' in df.columns:
        # pd.Series is the 1D array type in Pandas
        is_missing_mask: pd.Series = df['name'].isna()
        
        # Extracting the 'Dirty' data
        bad_rows: pd.DataFrame = df[is_missing_mask].copy()
        
        if not bad_rows.empty:
            bad_rows['error type'] = 'Missing Name'
            bad_rows['doc id'] = file_name
            bad_rows['sheet name'] = sheet_name
            error_list.append(bad_rows)
            
        # Returning the 'Clean' data
        df = df[~is_missing_mask]
        
    return df, error_list

def validate_emails(df: pd.DataFrame, file_name: str, sheet_name: str) -> Tuple[pd.DataFrame, List[pd.DataFrame]]:
    """Identifies and extracts rows with invalid emails."""
    error_list: List[pd.DataFrame] = []
    
    if 'email' in df.columns:
        # Check for presence of '@' - result is a Boolean Series
        has_no_at_symbol: pd.Series = ~df['email'].astype(str).str.contains('@', na=False)
        
        bad_rows: pd.DataFrame = df[has_no_at_symbol].copy()
        
        if not bad_rows.empty:
            bad_rows['error type'] = 'Invalid Email'
            bad_rows['doc id'] = file_name
            bad_rows['sheet name'] = sheet_name
            error_list.append(bad_rows)
            
        df = df[~has_no_at_symbol]
        
    return df, error_list