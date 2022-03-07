#python関連ライブラリ及びモジュール。
import sys
import os
import ruamel.yaml
import ruamel
import gspread
import glob
import traceback

#GoogleOauth認証関連ライブラリ及びモジュール。
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow

#ダウンロードした認証情報が保管してあるローカルのフォルダへのパス。GoogleAPIの認証画面で取得したjsonファイルを格##納するディレクトリーに格納。
CREDENTIAL_DIR = "認証のJSONファイルへの絶対パス"

#認証情報ファイルまでのパス
CLIENT_SECRET_FILE = os.path.join(CREDENTIAL_DIR, 'client_secret.json')

#一旦API認証が成功した後の認証情報を記録するファイルのパス
CREDENTIAL_PATH = os.path.join(CREDENTIAL_DIR, 'credential.json')

#以下で操作する範囲のAPIエンドポイントを定義。
SCOPES = "https://www.googleapis.com/auth/spreadsheets"

#対象のWorkbookのIDを以下に取得。
WORKBOOK_ID ='WorkbookIDを書く'

#関数を順に呼び出す。
def main():
     First = Get_Credential()
     Second = Read_Yml_Write(First)

#OAuth認証。
def Get_Credential():
 #credentialファイルが存在すれば、それで認証する。
    creds = None
    if os.path.exists(CREDENTIAL_PATH):
         #credsに返り値であるcredentialsが格納される。
        creds = Credentials.from_authorized_user_file(CREDENTIAL_PATH, SCOPES)
 # 未認証だった場合は許可を求める(ブラウザ認証)。
 #credsに代入されていない場合、または、tokenが代入されていてもexpireしている場合次のif文の処理に移る。
    if not creds or not creds.valid:
         #credsに代入されていて、expireされていて、OAuthがtokenをrefreshする場合。
        if creds and creds.expired and creds.refresh_token:
             #アクセスtokenをrefreshする。括弧内のparameterはhttp request。
            try: 
                creds.refresh(Request())
            except Exception as e:
                type_, value, _ = sys.exc_info()
                print(type_)
                print(value)
                sys.exit("An refreshing the credentials' access token failed. This error can happen after some changes, especially modyfing the scopes. Please delete the current credential.json and retry.")
        else:
             #client_secrets_fileからFlowインスタンスを生成。
            try:   
                 flow = InstalledAppFlow.from_client_secrets_file(
                 CLIENT_SECRET_FILE, SCOPES)
             #OAuth2.0のcreedentialに返す。credsに代入。
                 creds = flow.run_local_server(port=0)
 # 次回のためにcredential.jsonというファイルを書き込み権限で起案し、to_jsonでcredentialをJSON形式に生成。
                 with open('credential.json', 'w') as token:
                     token.write(creds.to_json())
            except Exception as e:
                type_, value, _ = sys.exc_info()
                print(type_)
                print(value)
                sys.exit("No such file linked to CLIENT_SECRET_FILE")
    return creds

#以下でymlファイルを読み込む。
def Read_Yml_Write(credential):
    yaml = ruamel.yaml.YAML()
#引数に取られたファイル及びdirectoryを順にfor文で処理する。
    for h in range(1, len(sys.argv)):
        file = sys.argv[h]
#引数に取られたものがdirectoryのケースとして、それを絶対パスに変換する。
        path = os.path.abspath(file)
#以下fileの場合とdirectoryの場合で分けて分岐し、directoryの場合はYMLファイルを析出。
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

#以下でspreadsheet上での操作。
def GSpread_Operate(credential, obj, file):
#SpreadSheetの認証。
    gc = gspread.authorize(credential)
    title = []
    try:
#ワークブックを開く。該当ワークブックのIDをカッコに格納。mainfileというワークブックを取得。
        sh = gc.open_by_key(WORKBOOK_ID)
        worksheet_list = sh.worksheets()
#対象fileがリストの中に同名のsheetがあるかないか確認。titleにsheet名を入れて、そこにfileと同名があるか否か。
        for sheet in worksheet_list:
            title.append(sheet.title)    
    except:        
        sys.exit('Wrong WORKBOOK_ID or No such Workbook to the ID . Please verify the ID')
#同名のファイル名があるケースと新規作成のケースを、取得したタイトルに名前があるかないかでchekck。
#上書きのケースは既存のものより上書きするものがレコード数が少ないケースを考慮し、空白の行・列がある場合は削除する仕様。
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
#以下該当のワークシートがなく新規作成のケース。まず、シートを作成。
        columns = Get_Column(obj)
        worksheet = sh.add_worksheet(title= file, rows=len(obj)+1, cols=len(columns))
        Write_Yml(credential, obj, file, columns, worksheet)

#sheetのcolumnの情報を取得する関数。
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

#情報を書き出す関数。
def Write_Yml(credential, obj, file, columns, worksheet):
    try:
#カラムを書き込む。
        worksheet.append_row(columns)
#valuesを書き出して、シートに書き込む。        
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
