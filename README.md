# YMLFileExportTool for GoogleSpreadSheet

このスクリプトは引数に取られたYML形式のファイル（引数にdirectoryをとる場合はそこに存在する全てのYML形式のファイル）の中身をGOOgleSheetAPIを用いて、GoogleSpreadSheet上に書き出すスクリプトである。このスクリプトはdict型のデータが格納されたlist形式のYMLファイルのレコードをkeyとvalueにわけ、spreadsheetの最初の行（raw）にはkeyをカラム（columns）として、それ以降の行に各カラムに対応する値（value）をカラム以後の行（raw）に書き込んでいく。

# DEMO
 以下、スクリプトが実行される際の挙動demo。  
 <img src="https://github.com/Chimay-0426/YMLFileExportTool-for-GoogleSpreadSheet/issues/1#issue-1160512823" width="400">
 
# CONFIGURATION
* ProductName:	Mac OS X
* ProductVersion:	10.15.7
* BuildVersion:	19H1217
* Python: 3.8.5
 
# Features
 
dict型のデータが格納されたlist形式のYMLファイルであれば、どんなにカラム数及びリスト内の要素数が多くとも、書き込むためのAPIcall数が2回で済む。なので、今回のようなOAuth認証を用いたケースでは大量のデータ数を引数にかなり大量にとってもcall数の単位時間あたりの上限に引っかかることはまずない。  
 
# Requirement
 Python関連ライブラリ
* ruame.yaml 0.17.20
* gspread 5.1.1 

 Google認証関連ライブラリ
* google.auth 2.3.2
* google_auth_oauthlib 0.4.6
 
# Installation
 
Requirementで列挙したライブラリのインストール方法。Macの場合はターミナルで実行。Windowsの場合はコマンドプロンプト(cmd)で実行。  

```bash
$ pip install ruamel.yaml
$ pip install gspread
$ pip install google-auth
$ pip install google-auth-oauthlib
```

# Usage
 
・scriptを実行の前に自身のGoogleAccountでAPIconsoleにログインし、プロジェクトを作成し、関連API（ここでは、GoogleSheetAPI）を有効にし、OAuth認証を作成し、認証ファイルであるjsonファイルを取得する。以下URLの【Google Sheets API の設定】を要参照（cf https://japan.appeon.com/technical/techblog/technicalblog019/）。  

From your command line:  
1. 対象リポジトリをclone
```bash
$ git clone https://github.com/~~~~
```
2. scriptが位置するdirectoryに移動。また同directoryに認証JSONfileをおく。
```bash
$ cd (a directory)
```
3. command lineでスクリプトを実行。command line引数には対象YMLfileが存在するdirectoryのパスを指定するか、スクリプトと同じdirectoryにあるYMLfileを指定する。
```bash
$ python ~ ./temp/ (or sample1.yml sample2.ym;)
```
*一回目の実行時にはGoogleAPIの初回ブラウザ認証が行われ、それを受け同じdirectoryにtokenが生成され、本スクリプトではcredential.jsonの名前でtokenが生成される。
 
# Note

・YMLを扱うためのライブラリとして比較的スタンダードに用いられるpyyamlは本スクリプトでは用いない。理由として、ユースケースとして考えられた対象YMLにYML1.1でbool型に解釈される文字列onが入って炊いたためで、1.2でloadするためruamel.yamlを用いた（cf https://yaml.org/type/bool.html）。

 
# Author

Name 山内　しんめい（Yamauchi SHIMMEI）  
E-Mail saepo12100426@gmail.com
 
