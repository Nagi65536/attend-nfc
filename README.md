# 部活動出席確認システム
nfcタグによる出席管理システムです
Googleスプレッドシートで管理できます

## 必要なもの
* RaspberryPi
* PaSoRi

## セットアップ
1. [このサイト](https://www.twilio.com/blog/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python-jp)に従って鍵を生成し、ファイル名を `client_secret.json` にする
2. スプレッドシートを設定する
3. ライブラリをインストール
   ```
   pip install -r requirements.txt
   ```

## ライセンス
[MIT ライセンス](/LICENSE)