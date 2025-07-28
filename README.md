# docs-convert-in-container

オフィス系ドキュメントをLLMで扱いやすい形式に変換するスクリプト群です

`docs`ディレクトリ内の `pptx`, `xlsx`, `pdf` ファイルをすべて探索し、
各ファイルをPNGやCSVに変換します。
変換されたファイルは、元ファイルと同じ階層に専用フォルダを作って保存されます。

### 変換後のツリー例
```
docs/
├── presentation.pptx
├── presentation_png/
│   ├── presentation-1.png
│   ├── presentation-2.png
│   └── presentation-3.png
├── data.xlsx
├── data_csv/
│   ├── data_sheet1.csv
│   └── data_sheet2.csv
├── project1/
│   ├── report.pptx
│   ├── report_png/
│   │   ├── report-1.png
│   │   └── report-2.png
│   └── analysis.xlsx
│   └── analysis_csv/
│       └── analysis_sheet1.csv
└── archive/
    ├── old_document.pdf
    └── old_document_png/
        ├── old_document-1.png
        └── old_document-2.png
```

## サポートされている拡張子と変換形式
- `pptx`: png（1スライド -> 1PNG）
- `xlsx`: csv（1シート -> 1CSV）
- `png`: png（1ページ -> 1PNG）

## 使い方

1. このリポジトリをclone

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
   docker compose run --rm converter bash generate-batch-and-run.sh docs/<dir>/
   ```
   ※変換が失敗したファイルは `logs/` ディレクトリの `batch-failed_<yyyyMMdd>.log` に記録されている想定

## TODO

<details>

- ワークフローを整える
- `xls`等の拡張子をサポート
- Pythonスクリプトに移行

</details>
