# 部活動出席確認システム
nfcタグによる出席管理システムです。
Googleスプレッドシートで管理できます。
また設定時間に音楽が流れます

## 必要なもの
- RaspberryPi
- PaSoRi

## 使用サービス/言語
- Google スプレッドシート
- Apps Script
- Google Drive API
- Python 3.10

## セットアップ - スプレッドシート
1. [このサイト](https://www.twilio.com/blog/an-easy-way-to-read-and-write-to-a-google-spreadsheet-in-python-jp)に従って鍵を生成し、ファイル名を `client_secret.json` にする
2.  `名簿` `設定` シートを追加する
3. [gas.js](./gas.js)　を `Apps Script` にコピーし、トリガーを設定する

### シートの設定
#### 名簿
| A          | B        | C   | D   | E     | F    | G      | H                 |   | 4月 | 5月 | ... | 3月 | 計 | | 音楽 |
|------------|----------|-----|-----|-------|------|--------|-------------------|---|-----|----|-----|-----|---|-|------|
| [連番用関数] | 氏名     | 学科 | 学年 | クラス | 番号  | 学籍番号 | RFID             |   |      |    |     |    |   | | 16:50 |
|            | 佐藤 智   | J   | 2   | Z     |  46  | 65536  | ABCDEF1234567890  |   |     |    |      |    |   | | 17:50 |
|            | 多田 隆人 | J   | 3   | A     |  31  | 201018  | PENGUINISCUTE987 |   |      |    |     |     |   | |      |
   

#### 連番用関数
```
=ARRAYFORMULA(IF(ISBLANK($G$1:$G$150),"",COUNTIFS($G$1:$G$150,"<>'",ROW($G$1:$G$150),"<="&ROW($G$1:$G$150)-1)))
```

## セットアップ - Python/音源
1. ライブラリをインストール
   ```shell
   pip install -r requirements.txt
   ```
2. `check.py` と `periodic.py` 内の `TODO` 部分を変更する
3. 音源を各自ダウンロードし、`sounds/` に入れる

### 変更点
```python
[check.py]
SE_ENTRY = '記録時'
SE_ALREADY = '記録済時'
SE_NEW = '新規登録時'
SE_NONE = '未登録時'
SE_ERROR = 'システムエラー'
SE_BOOT = '起動音'

SPREADSHEET_APP = 'スプレッドシートのファイル名'
ROSTER_NAME = '名簿シート名'
```

```python
[periodic.py]
MUSIC_FILE = '流す音楽'
SE_REMOVE = '削除時'
SE_NEW = '新規登録時'

SPREADSHEET_APP = 'スプレッドシート名'
SETTING_SHEET = '設定シート名'

```


## セットアップ - systemctl
以下のファイルを `/etc/systemd/system/` に用意する  
`()` 内は適宜変更する

### check.service
```shell
[Unit]
Description = Start the attendance system

[Service]
ExecStart=/usr/bin/python3 (パス)/check.py
```

### check.timer
```shell
[Unit]
Description = Start the attendance system

[Timer]
OnBootSec=1min
Unit=fes_check.server

[Install]
WantedBy=timers.target
```

### periodic.service
```shell
[Unit]
Description = run periodically

[Service]
ExecStart=/usr/bin/python (パス)/periodic.py
```

### periodic.timer
```shell
[Unit]
Description = run periodically

[Timer]
OnCalendar=*-*-* *:*:00

[Install]
WantedBy=timers.target
```

## 起動/停止方法
```shell
// 自動起動有効
sudo systemctl enable check.timer
sudo systemctl enable periodic.timer

// 即時起動
sudo systemctl start check.timer
sudo systemctl start periodic.timer

// 自動起動無効
sudo systemctl disable check.timer
sudo systemctl disable periodic.timer

// 即時停止
sudo systemctl stop check.timer
sudo systemctl stop periodic.timer

// 再起動
sudo systemctl restart check.timer
sudo systemctl restart periodic.timer
```
- ※1分後に実行されます(WiFi等の起動を待つため)


## ライセンス
[MIT ライセンス](/LICENSE)
