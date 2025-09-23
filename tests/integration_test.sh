#!/bin/bash

set -e

# Create temporary directory for testing
TMP_DIR="tests/tmp_$(date +%s)"
export TMP_DIR
echo "Creating temporary directory for testing: $TMP_DIR"

# Set up cleanup process
cleanup() {
    echo "Cleaning up temporary directory: $TMP_DIR"
    rm -rf "$TMP_DIR"
}
trap cleanup EXIT
# Success message function
success_message() {
    echo ""
    echo -e "\033[32m === $1 === \033[0m"
}

error_message() {
    echo ""
    echo -e "\033[31m === $1 === \033[0m"
}

# 0. Copy contents of tests/docs to temporary directory
cp -r tests/docs "$TMP_DIR"

# 1. Build container image to ensure latest tooling
echo "1. Building converter image..."
docker compose build converter

# 2. Execute convert command inside container
echo "2. Executing document conversion..."

docker compose run --rm converter bash generate-batch-and-run.sh "$TMP_DIR/"

# 3. Check existence of PDF and CSV files
echo "3. Checking generated files..."

if [ ! -f "$TMP_DIR/sample_csv/sample_sheet1.csv" ]; then
    error_message "sample_csv/sample_sheet1.csv was not generated"
    exit 1
fi

success_message "sample_csv/sample_sheet1.csv was generated"

if [ ! -f "$TMP_DIR/sample.pdf" ]; then
    error_message "sample.pdf was not generated"
    exit 1
fi

success_message "sample.pdf was generated"

# 4. Check CSV file contents
echo "4. Checking contents of sample_csv/sample_sheet1.csv..."

expected_content=$(cat "$TMP_DIR/expected.csv")
actual_content=$(cat "$TMP_DIR/sample_csv/sample_sheet1.csv")

if [ "$actual_content" != "$expected_content" ]; then
    error_message "sample_sheet1.csv content differs from expected"
    echo "Expected:"
    echo "$expected_content"
    echo "Actual:"
    echo "$actual_content"
    exit 1
fi

success_message "sample_csv/sample_sheet1.csv output is as expected"

# 5. Check PNG file contents
echo "5. Checking contents of sample_png/sample-1.png..."

expected_content=$(sha256sum "$TMP_DIR/expected.png" | awk '{print $1}')
actual_content=$(sha256sum "$TMP_DIR/sample_png/sample-1.png" | awk '{print $1}')

if [ "$actual_content" != "$expected_content" ]; then
    error_message "sample-1.png content differs from expected"
    exit 1
fi

success_message "sample_png/sample-1.png output is as expected"

# 6. Check PPTX JSONL outputs
echo "6. Checking contents of sample_jsonl/..."

expected_pptx_jsonl_hash=$(sha256sum "$TMP_DIR/expected_pptx-1.jsonl" | awk '{print $1}')
actual_pptx_jsonl_hash=$(sha256sum "$TMP_DIR/sample_jsonl/sample_page-1.jsonl" | awk '{print $1}')

if [ "$actual_pptx_jsonl_hash" != "$expected_pptx_jsonl_hash" ]; then
    error_message "sample_page-1.jsonl content differs from expected"
    echo "Expected hash: $expected_pptx_jsonl_hash"
    echo "Actual hash: $actual_pptx_jsonl_hash"
    exit 1
fi

success_message "sample_jsonl/sample_page-1.jsonl output is as expected"

# 7. Check XLSX JSONL outputs
echo "7. Checking XLSX JSONL outputs..."

expected_xlsx_jsonl_hash=$(sha256sum "$TMP_DIR/expected_xlsx-1.jsonl" | awk '{print $1}')
actual_xlsx_jsonl_hash=$(sha256sum "$TMP_DIR/sample_jsonl/sample_sheet1.jsonl" | awk '{print $1}')

if [ "$actual_xlsx_jsonl_hash" != "$expected_xlsx_jsonl_hash" ]; then
    error_message "sample_sheet1.jsonl content differs from expected"
    echo "Expected hash: $expected_xlsx_jsonl_hash"
    echo "Actual hash: $actual_xlsx_jsonl_hash"
    exit 1
fi

success_message "sample_jsonl/sample_sheet1.jsonl output is as expected"

echo ""
success_message "All tests passed successfully!"
