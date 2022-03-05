# YMLFileExportTool for GoogleSpreadSheet

このスクリプトは引数に取られたYML形式のファイル（引数にdirectoryをとる場合はそこに存在する全てのYML形式のファイル）の中身をGOOgleSheetAPIを用いて、GoogleSpreadSheet上に書き出すスクリプトである。このスクリプトはdict型のデータが格納されたlist形式のYMLファイルの中身をkeyとvalueにわけ、spreadsheetの最初の行にはkeyをカラムcolumnsとして、それ以降の行に各カラムに対応する

It reads YMLfiles taken as arguments (when taking a directory, the script finds YMLfiles and educes keys from them, temporalily putting them into an empty list called columns. if it finish putting all keys into  And, Based on the list, it educes the values corrsponding to the keys in the list putting them into an empty list called values.
-Details  

It has mainly two parts. Tne first one is a function for GoogleAPI OAuth, enable GoogleAPI Authentification called OAuth2.0 to call GoogleSheet API.   



# DEMO
 
"hoge"の魅力が直感的に伝えわるデモ動画や図解を載せる
 
# Features
 
"hoge"のセールスポイントや差別化などを説明する

# Configuration instructions
 
# Requirement
 
Python relevant library
import sys
import os
import ruamel.yaml
import ruamel
import gspread
import glob
import traceback

GoogleOauth library  
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
 
* huga 3.5.2
* hogehuga 1.0.2
 
# Installation
 
Requirementで列挙したライブラリなどのインストール方法を説明する
 
```bash
pip install huga_package
```
 
# Usage
 
DEMOの実行方法など、"hoge"の基本的な使い方を説明する
 
```bash
git clone https://github.com/hoge/~
cd examples
python demo.py
```
 
# Note

As a premise, this is for YMLfiles consisting of a list of dict data.  
注意点などがあれば書く
 
# Author

Name Yamauchi SHIMMEI  
E-Mail saepo12100426@gmail.com
 
# License
ライセンスを明示する
 
"hoge" is under [MIT license](https://en.wikipedia.org/wiki/MIT_License).
 
社内向けなら社外秘であることを明示してる
 
"hoge" is Confidential.
