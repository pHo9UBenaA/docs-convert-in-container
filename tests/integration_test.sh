#!/bin/bash

set -e

# Create temporary directory for testing
TMP_DIR="tests/tmp_$(date +%s)"
echo "Creating temporary directory for testing: $TMP_DIR"

# Set up cleanup process
cleanup() {
    echo "Cleaning up temporary directory: $TMP_DIR"
    rm -rf "$TMP_DIR"
}
trap cleanup EXIT

# Copy contents of tests/docs to temporary directory
cp -r tests/docs "$TMP_DIR"

# 1. Execute convert command inside container
echo "1. Executing document conversion..."

docker compose run --rm converter bash generate-batch-and-run.sh "$TMP_DIR/"

# 2. Check existence of PDF and CSV files
echo "2. Checking generated files..."

if [ ! -f "$TMP_DIR/sample_csv/sample_sheet1.csv" ]; then
    echo "Error: sample_csv/sample_sheet1.csv was not generated"
    exit 1
fi

echo "✓ sample_csv/sample_sheet1.csv was generated"
if [ ! -f "$TMP_DIR/sample.pdf" ]; then
    echo "Error: sample.pdf was not generated"
    exit 1
fi

echo "✓ sample.pdf was generated"

# 3. Check CSV file contents
echo "3. Checking contents of sample_csv/sample_sheet1.csv..."

expected_content=$(cat "$TMP_DIR/expected.csv")
actual_content=$(cat "$TMP_DIR/sample_csv/sample_sheet1.csv")

if [ "$actual_content" != "$expected_content" ]; then
    echo "Error: sample_sheet1.csv content differs from expected"
    echo "Expected:"
    echo "$expected_content"
    echo "Actual:"
    echo "$actual_content"
    exit 1
fi

echo "✓ sample_csv/sample_sheet1.csv output is as expected"

# 4. Check PNG file contents
echo "4. Checking contents of sample_png/sample-1.png..."

expected_content=$(sha256sum "$TMP_DIR/expected.png" | awk '{print $1}')
actual_content=$(sha256sum "$TMP_DIR/sample_png/sample-1.png" | awk '{print $1}')

if [ "$actual_content" != "$expected_content" ]; then
    echo "Error: sample-1.png content differs from expected"
    exit 1
fi

echo "✓ sample_png/sample-1.png output is as expected"

echo ""
echo -e "\033[32m=== All tests passed successfully! ===\033[0m"
