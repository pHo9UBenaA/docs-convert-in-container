# docs-convert-in-container

A collection of scripts designed to convert office document formats into a format more suitable for processing with LLMs.

This script recursively processes all `pptx`, `xlsx`, and `pdf` files within the `docs/` directory, converting them into PNG images or CSV files.

The converted files are saved in the same directory structure as their original counterparts.
<details>

<summary>Example of converted file hierarchy</summary>

```
docs/
├── presentation.pptx
├── presentation.pdf        # Generated from presentation.pptx
├── presentation_png/       # Generated from presentation.pptx
│   ├── slide-1.png
│   ├── slide-2.png
│   └── slide-3.png
├── presentation_jsonl/     # Generated from presentation.pptx
│   ├── slide-1.jsonl
│   ├── slide-2.jsonl
│   └── slide-3.jsonl
│
├── data.xlsx
├── data_csv/               # Generated from data.xlsx
│   ├── sheet-Sheet1.csv    # Using sheet name
│   └── sheet-Summary.csv   # Using sheet name
├── data_jsonl/             # Generated from data.xlsx
│   ├── sheet-Sheet1.jsonl  # Using sheet name
│   └── sheet-Summary.jsonl # Using sheet name
│
├── project1/
│   ├── report.pptx
│   ├── report.pdf          # Generated from report.pptx
│   ├── report_png/         # Generated from report.pptx
│   │   ├── slide-1.png
│   │   └── slide-2.png
│   ├── report_jsonl/       # Generated from report.pptx
│   │   ├── slide-1.jsonl
│   │   └── slide-2.jsonl
│   │
│   ├── analysis.xlsx
│   ├── analysis_csv/       # Generated from analysis.xlsx
│   │   └── sheet-Data.csv  # Using sheet name
│   └── analysis_jsonl/     # Generated from analysis.xlsx
│       └── sheet-Data.jsonl # Using sheet name
│
└── archive/
    ├── old_document.pdf
    └── old_document_png/   # Generated from old_document.pdf
        ├── page-1.png
        └── page-2.png
```

</details>

## Supported file extensions and conversion formats
- `pdf`: Converted to PNG images (one image per page)
- `pptx`:
  - PDF format (entire document as a single file)
  - PNG images (one image per slide)
  - JSONL format (structured data per slide)
- `xlsx`:
  - CSV format (one file per sheet)
  - JSONL format (structured data per sheet)

## Usage

1. Clone this repository
   ```bash
   git clone git@github.com:pHo9UBenaA/docs-convert-in-container.git
   ```

2. Copy the documents you wish to convert into the `docs/` directory
   ```bash
   cp -rf <path/to/docs/> docs/
   ```

3. Build the Docker container
   ```bash
   docker compose build converter 
   ```
   *NOTE: This process may take some time as it relies on LibreOffice.*

4. Run the conversion script
   ```bash
   docker compose run --rm converter bash generate-batch-and-run.sh docs/
   ```
   *Files that fail to convert will be logged in `logs/batch-failed_<yyyyMMdd>.log`.*

## Development

### Testing

```bash
bash tests/integration_test.sh
```

### Linting & Formatting

(TODO)

```bash
docker compose run --rm linter bash -lc '
set -euo pipefail
DIRS=("pptx-xml-to-jsonl" "xlsx-xml-to-jsonl" "shared-xml-to-jsonl")

# Linting
for dir in "${DIRS[@]}"; do
  echo "Building $dir..."
  (cd /scripts/"$dir" && dotnet build)
done

# Formatting
(cd /scripts && dotnet tool restore)
for dir in "${DIRS[@]}"; do
  echo "Formatting $dir..."
  (cd /scripts/"$dir" && dotnet format)
done
'
```
