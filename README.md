# docs-convert-in-container

オフィス系ドキュメントをLLMで扱いやすい形式に変換するスクリプト群です

`docs/` ディレクトリ内の `pptx`, `xlsx`, `pdf` ファイルをすべて探索し、PNGやCSVに変換します。

変換されたファイルは、元ファイルと同じ階層に保存されます。
<details>

<summary>変換後のツリー例</summary>

```
docs/
├── presentation.pptx
├── presentation_png/       # presentation.pptxから生成
│   ├── presentation-1.png
│   ├── presentation-2.png
│   └── presentation-3.png
│
├── data.xlsx
├── data_csv/               # data.xlsxから生成
│   ├── data_sheet1.csv
│   └── data_sheet2.csv
│
├── project1/
│   ├── report.pptx
│   ├── report_png/         # report.pptxから生成
│   │   ├── report-1.png
│   │   └── report-2.png
│   │
│   └── analysis.xlsx
│   └── analysis_csv/       # analysis.xlsxから生成
│       └── analysis_sheet1.csv
│
└── archive/
    ├── old_document.pdf
    └── old_document_png/   # old_document.pdfから生成
        ├── old_document-1.png
        └── old_document-2.png
```

</details>

## サポートされている拡張子と変換形式
- `pdf`: png（1ページ -> 1PNG）
- `pptx`: png（1スライド -> 1PNG）
- `xlsx`: csv（1シート -> 1CSV）

## 使い方

1. このリポジトリをclone
   ```bash
   git clone git@github.com:pHo9UBenaA/docs-convert-in-container.git
   ```

2. 変換したいドキュメント一覧を `docs` ディレクトリにコピー
   ```bash
   cp -rf <path/to/docs/> docs/
   ```

3. Dockerコンテナをビルド
   ```bash
   docker compose up -d --build
   ```
   ※LibreOfficeに依存しているため時間がかかります

4. 変換スクリプトを実行
   ```bash
   docker compose run --rm converter bash generate-batch-and-run.sh docs/
   ```
   ※変換に失敗した場合は `logs/` ディレクトリの `batch-failed_<yyyyMMdd>.log` に書き込まれます

## TODO

<details>

- ワークフローを整える
- `xls`等の拡張子をサポート
- Pythonスクリプトに移行

</details>
