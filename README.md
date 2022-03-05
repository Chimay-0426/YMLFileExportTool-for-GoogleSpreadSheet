# YMLFileExportTool for GoogleSpreadSheet

このスクリプトは引数に取られたYML形式のファイル（引数にdirectoryをとる場合はそこに存在する全てのYML形式のファイル）の中身をGOOgleSheetAPIを用いて、GoogleSpreadSheet上に書き出すスクリプトである。このスクリプトはdict型のデータが格納されたlist形式のYMLファイルのレコードをkeyとvalueにわけ、spreadsheetの最初の行（raw）にはkeyをカラム（columns）として、それ以降の行に各カラムに対応する値（value）をカラム以後の行（raw）に書き込んでいく。

# DEMO
 
"hoge"の魅力が直感的に伝えわるデモ動画や図解を載せる
 
# Features
 
dict型のデータが格納されたlist形式のYMLファイルであれば、どんなにカラム数及びリスト内の要素数が多くとも、書き込むためのAPIcall数が2回で済む。なので、今回のようなOAuth認証を用いいたケースでは大量のデータ数を引数にかなり大量にとってもcall数の単位時間あたりの上限に引っかかることはまずない。  
 
# Requirement
 
Python relevant library
import ruamel.yaml
import ruamel
import gspread

GoogleOauth relevant library  
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
 
・scriptを実行の前に自身のGoogleAccountでAPIconsoleにログインし、プロジェクトを作成し、OAuth認証を作成し、認証ファイルであるjsonファイルを取得する。そして、関連API（ここでは、GoogleSheetAPI）を有効にする。

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
