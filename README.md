# 部活動出席確認システム
nfcタグによる出席管理システムです
Googleスプレッドシートで管理できます

## 必要なもの
- RaspberryPi
- PaSoRi

## 使用サービス/言語
- Google スプレッドシート
- Apps Script
- Google Drive API
- Python 3.8~

## セットアップ
1. [このサイト](https://www.twilio.com/blog/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python-jp)に従って鍵を生成し、ファイル名を `client_secret.json` にする
2. ライブラリをインストール
   ```
   pip install -r requirements.txt
   ```
3. スプレッドシートに `名簿` `テンプレート` をセットする [参考](#シートの設定)
4. `gas.js`　を Apps Script にセットする
5. トリガーをセットする [参考](#トリガー)
6. `systemctl` 用のファイルを用意する [参考](#systemctl用ファイル)


### シートの設定
#### 名簿
| 連番用関数 | 氏名   | 学科 | 学年 | クラス | 番号 | 学籍番号 | RFID              | ソート用関数 |
|----------|--------|-----|-----|-------|------|--------|-------------------|------------|
| 0        | 佐藤 智 | J   | 3   | Z     |  46  | 65536  | ABCDEF1234567890  | 3JZ46      |

#### テンプレート
| 連番用関数 | 氏名   | 学科 | 学年 | クラス | 番号 | 学籍番号 | ソート用関数 | | 合計 | 1   | 2 | 3 | ... | 31  | 合計 |
|----------|--------|-----|-----|-------|------|--------|------------|-|------|----|---|---|------|-----|-----|
| 0        | 佐藤 智 | J   | 3   | Z     |  46  | 65536  | 3JZ46      | | 5    | 16 | ◯ |   |      |     | 5   |

#### 連番用関数
- 1列目のみ
```
=ARRAYFORMULA(IF(ISBLANK($G$1:$G$150),"",COUNTIFS($G$1:$G$150,"<>'",ROW($G$1:$G$150),"<="&ROW($G$1:$G$150)-1)))
```
#### ソート用関数
- 全列に追加
```
=IF(ISBLANK(B2),"", COUNTA(K2:AO2))
```

### トリガー

<img src="./images/gas-1.png" width="400">
<img src="./images/gas-2.png" width="400">


### systemctl用ファイル
```
cd /etc/systemd/system/
touch {check.service,check.timer}
```

### check.service
```
```

### check.timer
```
```


## ライセンス
[MIT ライセンス](/LICENSE)