"""Module for interacting with google sheets."""


import pandas as pd
from googleapiclient.discovery import build
from google.oauth2 import service_account
from IPython.display import display


def service():
    """Returns a service object to access google sheets."""
    
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    keys = 'keys.json'
    creds = None
    creds = service_account.Credentials.from_service_account_file(keys, scopes=scopes)
    service = build('sheets', 'v4', credentials=creds)
    
    return service


def clear(spreadsheet_id, range_):
    """Executes Google Sheets clear request and returns the response."""
    
    # Assign params
    params = {
        'spreadsheetId': spreadsheet_id,
        'range': range_
    }
    
    # Assign request
    request = service().spreadsheets().values().clear(**params)
    
    # Execute request and assign response
    response = request.execute()
    return response


def update(spreadsheet_id, range_, values, value_input_option='USER_ENTERED'):
    """Executes Google Sheets update request and returns the response."""
    
    # Assign body param
    body = {
        'values': values
    }
    
    # Assign params
    params = {
        'spreadsheetId': spreadsheet_id,
        'range': range_,
        'valueInputOption': value_input_option,
        'body': body
    }
    
    # Assign request
    request = service().spreadsheets().values().update(**params)
    
    # Execute request and return response
    response = request.execute()
    return response


def get(spreadsheet_id, range_):
    """Executes a Google Sheets get request and returns the response."""
    
    # Assign params
    params = {
        'spreadsheetId': spreadsheet_id,
        'range': range_
    }
    
    # Assign request
    request = service().spreadsheets().values().get(**params)
    
    # Execute request and return response
    response = request.execute()
    return response


def values(df, index=False):
    """Converts Pandas DataFrame to values list."""
    
    # Assign empty list
    values = []
    
    # If index is wanted in values list, put the index into the body and make it the first column
    if index:    
        df[df.index.name] = df.index
        index_col = df.pop(df.index.name)
        df.insert(0, df.index.name, index_col)
    
    # Append column headers to values
    values.append(list(df.columns))
    
    # Append each row of dataframe values to values
    for row in df.values.tolist():
        values.append(row)
        
    return values


def df(values, index=None):
    """Converts values list to Pandas DataFrame"""
    
    # Assign headers and body
    headers = values[0]
    body = values[1:]
    
    # Assign DataFrame
    df = pd.DataFrame(body, columns=headers)
    
    # If index is passed, set index
    if index != None:
        df.set_index(index, inplace=True)
    
    return df