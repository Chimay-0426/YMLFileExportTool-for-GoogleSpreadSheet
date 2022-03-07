import sys
import os
import ruamel.yaml
import ruamel
import gspread
import glob
import traceback

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

CREDENTIAL_DIR = "認証のJSONファイルへの絶対パス"

CLIENT_SECRET_FILE = os.path.join(CREDENTIAL_DIR, 'client_secret.json')

CREDENTIAL_PATH = os.path.join(CREDENTIAL_DIR, 'credential.json')

SCOPES = "https://www.googleapis.com/auth/spreadsheets"

WORKBOOK_ID ='WorkbookIDを書く'

def main():
     First = Get_Credential()
     Second = Read_Yml_Write(First)

def Get_Credential():
    creds = None
    if os.path.exists(CREDENTIAL_PATH):
        creds = Credentials.from_authorized_user_file(CREDENTIAL_PATH, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try: 
                creds.refresh(Request())
            except Exception as e:
                type_, value, _ = sys.exc_info()
                print(type_)
                print(value)
                sys.exit("An refreshing the credentials' access token failed. This error can happen after some changes, especially modyfing the scopes. Please delete the current credential.json and retry.")
        else:
            try:   
                 flow = InstalledAppFlow.from_client_secrets_file(
                 CLIENT_SECRET_FILE, SCOPES)
                 creds = flow.run_local_server(port=0)
                 with open('credential.json', 'w') as token:
                     token.write(creds.to_json())
            except Exception as e:
                type_, value, _ = sys.exc_info()
                print(type_)
                print(value)
                sys.exit("No such file linked to CLIENT_SECRET_FILE")
    return creds

def Read_Yml_Write(credential):
    yaml = ruamel.yaml.YAML()
    for h in range(1, len(sys.argv)):
        file = sys.argv[h]
        path = os.path.abspath(file)
        if not os.path.isfile(path) and not os.path.isdir(path):
            sys.exit('No such file or directory')
        elif os.path.isfile(path):
            try:
                with open(file) as s:
                    obj = yaml.load(s)
                    GSpread_Operate(credential, obj, file)              
            except Exception as e:
                type_, value, _ = sys.exc_info()
                print(type_)
                print(value)
        else:
            file_name_list = glob.glob(path + "/*.yml")
            if len(file_name_list) == 0:
                sys.exit("No such file in the target directory")
            else:
                pass
            for files in file_name_list:
                basename = os.path.basename(files)
                try:
                    with open(files) as s:
                        obj = yaml.load(s)
                        GSpread_Operate(credential, obj, basename)
                except Exception as e:
                        type_, value, _ = sys.exc_info()
                        print(type_)
                        print(value)

def GSpread_Operate(credential, obj, file):
#SpreadSheetの認証。
    gc = gspread.authorize(credential)
    title = []
    try:    
        sh = gc.open_by_key(WORKBOOK_ID)
        worksheet_list = sh.worksheets()
        for sheet in worksheet_list:
            title.append(sheet.title)    
    except:        
        sys.exit('Wrong WORKBOOK_ID or No such Workbook to the ID . Please verify the ID')
    if file in title:
        columns = Get_Column(obj)
        worksheet = sh.worksheet(file)
        worksheet.clear() 
        Write_Yml(credential, obj, file, columns, worksheet)
        if worksheet.col_count > len(columns):
            worksheet.delete_columns(len(columns) + 1, worksheet.col_count)  
        else:
            pass
       if worksheet.row_count > len(obj) + 1:
            worksheet.delete_rows(len(obj) + 2, worksheet.row_count)
        else:
            pass
    else:
        columns = Get_Column(obj)
        worksheet = sh.add_worksheet(title= file, rows=len(obj)+1, cols=len(columns))
        Write_Yml(credential, obj, file, columns, worksheet)

def Get_Column(obj):
    columns = []
    for i in range(0, len(obj)):
        list_keys = list(obj[i].keys())
        for l in range(0, len(list_keys)):
            if list_keys[l] in columns:
                pass
            else:
                columns.append(list_keys[l])
    return columns

def Write_Yml(credential, obj, file, columns, worksheet):
    try:
        worksheet.append_row(columns)
        values = []
        for i in range(0, len(obj)):
            value_units = [] 
            for x in columns:
                if x in obj[i].keys():
                    value_units.append(obj[i][x])
                else:
                    value_units.append(None)
            values.append(value_units)
        try:
            worksheet.append_rows(values, value_input_option='RAW', insert_data_option=None, table_range=None, include_values_in_response=False)
        except gspread.exceptions.APIError as e:    
            type_, value, _ = sys.exc_info()
            print(type_)
            print(value)
            sys.exit('There are some invalid values in the data to write in a worksheet. Thus, you cannot write them to a worksheet. Please refer to "Invalid values[(number)][(number)]" mentioned above, which means a position of the invalid values in the yml files. ')
    except Exception as e:
        type_, value, _ = sys.exc_info()
        print(type_)
        print(value)

if __name__ == '__main__':
    res = main()
    exit(res) 
