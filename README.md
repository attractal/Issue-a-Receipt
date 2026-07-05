# 領収書自動発行ツール

`receipt_template_source.docx` を元に、宛名・税込金額・但し書き・発行日・担当名を入れた領収書 DOCX を作成します。

## 1 枚だけ発行

```powershell
python .\generate_receipt.py `
  --recipient "宛名" 
  --amount 金額 
  --description "但し書き" 
  --staff "渉外担当もしくはリーダー" 
  --date 日付 
  --output ".\出力ファイル名.docx"
```

税込金額から税抜金額と消費税等は自動計算されます。`--staff` を省略した場合は `ScienceTechno` になります。

## CSV から発行

`receipts.csv`:
下の記入例を参考に csv ファイルを生成し、 `generate_receipt.py` と同じディレクトリにいれて実行してください。ファイル名は必ず `receipts.csv` にしてください。
```csv
recipient,amount,description,date,staff
宛名,10000,但し書き1,日付,リーダー
宛名,3000,但し書き2,日付,リーダー
```

```powershell
python .\generate_receipt.py --csv .\receipts.csv --output .\出力名.docx
```

元の DOCX は 1 ページに上下 2 枚の領収書があるため、CSV の先頭 2 行を上下に差し込みます。
1 行ずつ別ファイルにしたい場合は次のようにします。

```powershell
python .\generate_receipt.py --csv .\receipts.csv --output .\出力名.docx --keep-one-per-file
```

## よくあるエラーと対処

### `python` が認識されない

表示例:

```text
python : 用語 'python' は、コマンドレット、関数、スクリプト ファイル、または操作可能なプログラムの名前として認識されません。
```

原因:

PC に Python が入っていない、または PATH が通っていません。

対処:

Python をインストールしましょう。
Python インストール とかで調べると色々記事が出るはず。


### テンプレートが見つからない

表示例:

```text
FileNotFoundError: Template not found: ...\receipt_template_source.docx
```

原因:

`generate_receipt.py` と同じフォルダに `receipt_template_source.docx` がありません。

対処:

`receipt_template_source.docx` を `generate_receipt.py` と同じフォルダに置きます。
別の場所にあるテンプレートを使う場合は `--template` で指定します。
`receipt_template_source.docx` は過去の領収書をダウンロードして使ってもらえれば大丈夫です。形式は2026年以降発行に対応しています。ファイル名を変更するのを忘れずに。

```powershell
python .\generate_receipt.py --template "C:\path\to\receipt_template_source.docx" --recipient "宛名" --amount 10000 --description "但し書き"
```

### CSV の列名エラー

表示例:

```text
KeyError: 'recipient'
KeyError: 'amount'
KeyError: 'description'
```

原因:

CSV の 1 行目の列名が足りない、またはスペルが違います。

対処:

CSV の 1 行目を次の形にします。`staff` は省略できます。

```csv
recipient,amount,description,date,staff
```

### 金額の形式エラー

表示例:

```text
ValueError: invalid literal for int() with base 10
```

原因:

`--amount` や CSV の `amount` に、数字として読めない文字が入っています。

対処:

金額は `10000` または `10,000` のように書きます。全角数字や `円` は避けてください。

### 日付の形式エラー

表示例:

```text
ValueError: Unsupported date format: ...
```

原因:

日付が対応していない形式です。

対処:

次のどれかの形式で書きます。

```text
2026-01-14
2026/01/14
2026年01月14日
```

### 必須項目が足りない

表示例:

```text
ValueError: missing required option(s): --recipient, --amount, --description
```

原因:

CSV を使わない発行で、宛名・金額・但し書きのどれかが指定されていません。

対処:

最低限、次の 3 つを指定します。

```powershell
python .\generate_receipt.py --recipient "宛名" --amount 10000 --description "但し書き"
```

### CSV が文字化けする

原因:

Excel などで保存した CSV の文字コードが合っていない可能性があります。

対処:

CSV は UTF-8 で保存してください。Excel なら「CSV UTF-8」を選ぶと安定します。

### 出力ファイルが Word で開けない

原因:

出力中に処理が中断した、またはファイルを Word で開いたまま上書きした可能性があります。

対処:

Word で出力先の `.docx` を閉じてから、もう一度実行してください。別名で出す場合は `--output` を変えます。

```powershell
python .\generate_receipt.py --recipient "宛名" --amount 10000 --description "但し書き" --output .\issued_receipt_2.docx
```
