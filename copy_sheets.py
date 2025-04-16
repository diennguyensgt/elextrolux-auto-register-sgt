from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os.path
import pickle

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_credentials():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    return creds

def copy_sheets(source_spreadsheet_id, source_range, destination_spreadsheet_id, destination_range):
    creds = get_credentials()
    service = build('sheets', 'v4', credentials=creds)

    # Get data from source spreadsheet
    result = service.spreadsheets().values().get(
        spreadsheetId=source_spreadsheet_id,
        range=source_range
    ).execute()
    values = result.get('values', [])

    if not values:
        print('No data found in source spreadsheet.')
        return

    # Prepare the data for writing
    body = {
        'values': values
    }

    # Write data to destination spreadsheet
    result = service.spreadsheets().values().update(
        spreadsheetId=destination_spreadsheet_id,
        range=destination_range,
        valueInputOption='RAW',
        body=body
    ).execute()
    print(f"{result.get('updatedCells')} cells updated.")

if __name__ == '__main__':
    # Thông tin bảng tính nguồn
    SOURCE_SPREADSHEET_ID = input("Nhập ID của bảng tính nguồn: ")
    SOURCE_RANGE = input("Nhập phạm vi dữ liệu nguồn (ví dụ: Sheet1!A1:Z1000): ")
    
    # Thông tin bảng tính đích
    DESTINATION_SPREADSHEET_ID = input("Nhập ID của bảng tính đích: ")
    DESTINATION_RANGE = input("Nhập vị trí bắt đầu sao chép trong bảng tính đích (ví dụ: Sheet1!A1): ")
    
    print("\nBắt đầu sao chép dữ liệu...")
    copy_sheets(SOURCE_SPREADSHEET_ID, SOURCE_RANGE, 
                DESTINATION_SPREADSHEET_ID, DESTINATION_RANGE)
    print("Hoàn thành sao chép dữ liệu!") 