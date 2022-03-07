# YMLFileExportTool for GoogleSpreadSheet

このスクリプトは引数に取られたYML形式のファイル（引数にdirectoryをとる場合はそこに存在する全てのYML形式のファイル）の中身をGoogleSheetAPIを用いて、GoogleSpreadSheet上に書き出すスクリプトである。このスクリプトはdict型のデータが格納されたlist形式のYMLファイルのレコードをkeyとvalueにわけ、spreadsheetの最初の行（raw）にはkeyをカラム（columns）として、それ以降の行に各カラムに対応する値（value）をカラム以後の行（raw）に書き込んでいく。取り込むYMLファイルと同名のsheet名があれば上書きを行い、ない場合は新規作成としてYMLファイル名でsheetを作成し、書き出す。

# DEMO
 以下、スクリプトが実行される際のDEMO。上書きのケースで、すでに白紙の同名sheetが対象spreadsheetに存在するケース。  
 <img src="https://user-images.githubusercontent.com/89734301/156912194-520bdce3-ceba-4706-bbd9-a667233f1602.gif" width="600">
 
# CONFIGURATION
* ProductName:	Mac OS X
* ProductVersion:	10.15.7
* BuildVersion:	19H1217
* Python: 3.8.5
 
# Features
 
・dict型のデータが格納されたlist形式のYMLファイルであれば、どんなにカラム数及びリスト内の要素数が多くとも、書き込むためのAPIcall数が2回で済む。なので、今回のようなOAuth認証を用いたケースでは大量のデータ数を引数にとってもcall数の単位時間あたりの上限に引っかかることはまずない。  
・OAuth認証は基本的にGoogleのQuickstart(cf https://developers.google.com/people/quickstart/python )に依拠しているが、error handlingのため部分的に改変している。　　
 
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
 
scriptを実行の前に自身のGoogleAccountでAPIconsoleにログインし、プロジェクトを作成し、関連API（ここでは、GoogleSheetAPI）を有効にし、OAuth認証を作成し、認証ファイルであるJSONファイルを取得する。JSONファイル名を”client_secret.json”に変更する。以下URLの【Google Sheets API の設定】を要参照（cf https://japan.appeon.com/technical/techblog/technicalblog019/）。  

From your command line:  
1. 対象リポジトリをclone
```bash
$ git clone https://github.com/Chimay-0426/YMLFileExportTool-for-GoogleSpreadSheet.git
```
2. cloneを実行したscriptが位置するdirectoryに移動。また同directoryに認証JSONfileをおく。
```bash
$ cd (a directory)
```
3. command lineでスクリプトを実行。command line引数には対象YMLfileが存在するdirectoryのパスを指定するか、スクリプトと同じdirectoryにあるYMLfileを指定する。
```bash
$ python YMLFileExportTool.py ./temp/ (or sample1.yml sample2.yml)
```
*一回目の実行時にはGoogleAPIの初回ブラウザ認証が行われ（以下詳細な説明を割愛するが、ブラウザ認証のフロー。）、それを受け同じdirectoryにtokenが生成され、本スクリプトではcredential.jsonの名前でtokenが生成される。  
*本リポジトリにsampleのYMLファイルを3つあるので、試してみてください。(sample1.yml, sample2.yml, sample3.yml)  
https://github.com/Chimay-0426/YMLFileExportTool-for-GoogleSpreadSheet/issues/4#issue-1160943578

（以下初回のブラウザ認証のフロー）  
  <img src="https://user-images.githubusercontent.com/89734301/156970130-538e1cfe-29a0-4e31-93f1-80eeceea1934.gif" width="600">
 
 
# Note

・YMLを扱うためのライブラリとして比較的スタンダードに用いられるpyyamlは本スクリプトでは用いない。理由として、ユースケースとして考えられた対象YMLにYML1.1でbool型に解釈される文字列onが入っていたためで、1.2でloadするためruamel.yamlを用いた（cf https://yaml.org/type/bool.html）。

 
# Author

Name 山内　しんめい（Yamauchi SHIMMEI）  
E-Mail saepo12100426@gmail.com  
 
