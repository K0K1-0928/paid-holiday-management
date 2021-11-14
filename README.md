# paid-holiday-management

Google スプレッドシートの有給休暇管理表と併用し、  
社員の年次有給休暇付与日を迎えた際に自動で入力するプログラムです。  
スタンドアロン型の Google Apps Script ですので、  
clone 後に clasp 等でお使いの Apps Script 環境に適用すると扱いやすいと思います。

## How To Use

以下の値を 1 行目に入力している Google スプレッドシートがあることを前提とします。

| 社員名 | メールアドレス | 入社日 | 今年度付与日数 | 今年度分残日数 | 前年度付与日数 | 前年度分残日数 | 前々年度未消化 |
| :----- | :------------- | :----- | :------------- | :------------- | :------------- | :------------- | :------------- |
|        |                |        |                |                |                |                |                |

また、シート名は `有給休暇管理表` としています。

1. スプレッドシートの ID を取得します。ID は `docs.google.com/spreadsheets/d/xxxxx/edit#gid=0` 等の `xxxxx` の部分です。
2. code.ts の `const sheetId: string = 'SpreadSheetID';` の SpreadSheetID を 1 で取得した ID に書き換えます。
3. clasp 等で Apps Script 環境に本プログラムを用意します。
4. 1 度 Apps Script 環境で実行し、スプレッドシートへのアクセスを許可します。
5. Apps Script 環境でトリガーを設定します。1 日ごとの実行トリガーを推奨します。
