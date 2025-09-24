#!/bin/bash

# Exit on error
set -e

# Constants
readonly DEFAULT_LOCALE="ja_JP.UTF-8"
readonly DOCS_ROOT="/docs"
readonly PNG_SUFFIX="_png"
readonly CSV_SUFFIX="_csv"
readonly PDF_EXTENSION="pdf"
readonly JSONL_EXTENSION="jsonl"
readonly PNG_RESOLUTION_DPI=300
readonly CONVERSION_WAIT_SECONDS=2
readonly MAX_DISPLAY_FILES=10
readonly EXIT_CODE_SUCCESS=0
readonly EXIT_CODE_ERROR=1

# Supported formats
readonly SUPPORTED_FORMAT_PPTX="pptx"
readonly SUPPORTED_FORMAT_PDF="pdf"
readonly SUPPORTED_FORMAT_XLSX="xlsx"

# Set locale for proper Japanese character handling
export LANG="${DEFAULT_LOCALE}"
export LC_ALL="${DEFAULT_LOCALE}"

# Display usage
if [ $# -eq 0 ]; then
    echo "Usage: $0 <input_file>"
    echo "Supported formats: .pptx, .pdf, .xlsx"
    echo "Output: PNG images (one per slide/page) or CSV file (for .xlsx)"
    exit "${EXIT_CODE_ERROR}"
fi

# Function to check if file exists
is_file_exists() {
    local file_path="$1"
    [ -f "$file_path" ]
}

# Function to create directory if not exists
create_directory_if_not_exists() {
    local dir_path="$1"
    mkdir -p "$dir_path"
}

# Function to normalize and parse file paths
parse_file_path() {
    local input="$1"
    
    # Remove 'docs/' prefix if present
    [[ "$input" == docs/* ]] && input="${input#docs/}"
    
    # Create absolute path
    if [[ "$input" == /* ]]; then
        # Already an absolute path
        FULL_PATH="$input"
    else
        # Relative path, prepend /docs/
        FULL_PATH="/docs/$input"
    fi
    
    # Extract filename components
    local filename=$(basename "$FULL_PATH")
    BASENAME="${filename%.*}"
    EXTENSION="${filename##*.}"
    
    # Get directory path
    local dir=$(dirname "$FULL_PATH")
    
    # Set output paths
    OUTPUT_DIR="${dir}/${BASENAME}${PNG_SUFFIX}"
    CSV_OUTPUT_DIR="${dir}/${BASENAME}${CSV_SUFFIX}"
    PDF_FILE="${dir}/${BASENAME}.${PDF_EXTENSION}"
    PDF_OUTDIR="${dir}"
    JSONL_OUTPUT_DIR="${dir}/${BASENAME}_jsonl"
    XLSX_JSONL_OUTPUT_DIR="${dir}/${BASENAME}_jsonl"
}

# Function to convert PDF to PNG images
convert_pdf_to_png_images() {
    local pdf_path="$1"
    local prefix="$2"  # Optional prefix parameter (e.g., "slide" or "page")

    # Use provided prefix or default to BASENAME for backward compatibility
    local output_prefix="${prefix:-${BASENAME}}"

    echo "Converting PDF to PNG images..."
    pdftoppm -png -r "${PNG_RESOLUTION_DPI}" "$pdf_path" "$OUTPUT_DIR/${output_prefix}"
}

# Function to check if PDF file exists
is_pdf_file_exists() {
    local pdf_path="$1"
    is_file_exists "$pdf_path"
}

# Function to convert PPTX to PDF using LibreOffice
convert_pptx_to_pdf_with_libreoffice() {
    local pptx_path="$1"
    local output_dir="$2"
    echo "Converting PPTX to PDF..."
    libreoffice --headless --convert-to pdf --outdir "$output_dir" "$pptx_path"
}

# Function to wait for file conversion to complete
wait_for_conversion_completion() {
    sleep "${CONVERSION_WAIT_SECONDS}"
}

# Function to convert PPTX to PNG via PDF
convert_pptx_to_png_via_pdf() {
    local pptx_path="$1"

    # Convert PPTX to PDF
    if ! is_pdf_file_exists "$PDF_FILE"; then
        convert_pptx_to_pdf_with_libreoffice "$pptx_path" "$PDF_OUTDIR"
    else
        echo "Using existing PDF: $PDF_FILE"
    fi

    # Wait for conversion to complete
    wait_for_conversion_completion

    # Check if PDF was created
    if ! is_pdf_file_exists "$PDF_FILE"; then
        echo "Error: Failed to convert PPTX to PDF"
        exit "${EXIT_CODE_ERROR}"
    fi

    # Convert PDF to PNG with "slide" prefix for PPTX files
    convert_pdf_to_png_images "$PDF_FILE" "slide"
}

# Function to convert PPTX package XML into JSONL
convert_pptx_to_jsonl() {
    local pptx_path="$1"
    local jsonl_output_dir="$2"

    echo "Extracting PPTX XML to JSONL..."
    pptx-xml-to-jsonl "$pptx_path" "$jsonl_output_dir"
}

# Function to convert XLSX package XML into JSONL
convert_xlsx_to_jsonl() {
    local xlsx_path="$1"
    local jsonl_output_dir="$2"

    echo "Extracting XLSX XML to JSONL..."
    xlsx-xml-to-jsonl "$xlsx_path" "$jsonl_output_dir"
}

# Function to convert Excel sheets to CSV files
convert_excel_sheets_to_csv_files() {
    local excel_path="$1"
    echo "Converting Excel sheets to individual CSV files..."
    # ssconvert uses %s for sheet number, we'll rename afterward to add hyphen
    ssconvert -S "$excel_path" "$CSV_OUTPUT_DIR/sheet%s.csv"

    # Rename files to use sheet-N format
    for file in "$CSV_OUTPUT_DIR"/sheet*.csv; do
        if [ -f "$file" ]; then
            # Extract the sheet number and rename
            base=$(basename "$file")
            num="${base#sheet}"
            num="${num%.csv}"
            mv "$file" "$CSV_OUTPUT_DIR/sheet-${num}.csv"
        fi
    done
}

# Function to count files by pattern
count_files_by_pattern() {
    local pattern="$1"
    ls -1 $pattern 2>/dev/null | wc -l
}

# Function to display CSV conversion results
display_csv_conversion_results() {
    local csv_count
    csv_count=$(count_files_by_pattern "$CSV_OUTPUT_DIR/*.csv")
    echo "Success: Created $csv_count CSV files in ${CSV_OUTPUT_DIR#${DOCS_ROOT}/}"
}

# Function to convert Excel to CSV (sheet by sheet)
convert_excel_to_csv_sheet_by_sheet() {
    local excel_path="$1"

    # Create CSV output directory
    create_directory_if_not_exists "$CSV_OUTPUT_DIR"

    # Convert all sheets to CSV files
    convert_excel_sheets_to_csv_files "$excel_path"

    # Count and display results
    display_csv_conversion_results
}

# Function to display PNG conversion results with file limit
display_png_conversion_results_with_limit() {
    local png_count
    png_count=$(count_files_by_pattern "$OUTPUT_DIR/*.png")

    if [ "$png_count" -gt 0 ]; then
        echo "Success: Created $png_count PNG files in ${OUTPUT_DIR#${DOCS_ROOT}/}"
        echo "Files created:"
        ls -1 "$OUTPUT_DIR"/*.png | sed "s|^${DOCS_ROOT}/||" | head -"${MAX_DISPLAY_FILES}"
        if [ "$png_count" -gt "${MAX_DISPLAY_FILES}" ]; then
            echo "... and $((png_count - MAX_DISPLAY_FILES)) more files"
        fi
    else
        echo "Error: No PNG files were created"
        exit "${EXIT_CODE_ERROR}"
    fi
}

# Function to display JSONL conversion results
display_jsonl_conversion_results() {
    local jsonl_dir="$1"
    local jsonl_count

    if [ -d "$jsonl_dir" ]; then
        jsonl_count=$(count_files_by_pattern "$jsonl_dir/*.jsonl")
        echo "Success: Created $jsonl_count JSONL files in ${jsonl_dir#${DOCS_ROOT}/}"
    else
        echo "Error: JSONL directory was not created"
        exit "${EXIT_CODE_ERROR}"
    fi
}

# Function to display conversion type message
display_conversion_type_message() {
    local extension="$1"
    if [ "$extension" = "${SUPPORTED_FORMAT_XLSX}" ]; then
        echo "Converting $INPUT_FILE to CSV..."
    else
        echo "Converting $INPUT_FILE to PNG..."
    fi
}

# Main execution
INPUT_FILE="$1"

# Parse the file path
parse_file_path "$INPUT_FILE"

# Check if file exists
if ! is_file_exists "$FULL_PATH"; then
    echo "Error: File not found: $FULL_PATH"
    exit "${EXIT_CODE_ERROR}"
fi

# Create output directory
create_directory_if_not_exists "$OUTPUT_DIR"

display_conversion_type_message "$EXTENSION"

case "$EXTENSION" in
    "${SUPPORTED_FORMAT_PPTX}")
        # Convert PPTX to PNG
        convert_pptx_to_png_via_pdf "$FULL_PATH"
        convert_pptx_to_jsonl "$FULL_PATH" "$JSONL_OUTPUT_DIR"
        # Display results
        display_png_conversion_results_with_limit
        display_jsonl_conversion_results "$JSONL_OUTPUT_DIR"
        ;;

    "${SUPPORTED_FORMAT_XLSX}")
        # Convert Excel to CSV and JSONL
        convert_excel_to_csv_sheet_by_sheet "$FULL_PATH"
        convert_xlsx_to_jsonl "$FULL_PATH" "$XLSX_JSONL_OUTPUT_DIR"
        # Display results
        display_jsonl_conversion_results "$XLSX_JSONL_OUTPUT_DIR"
        ;;

    "${SUPPORTED_FORMAT_PDF}")
        # Convert PDF directly to PNG with "page" prefix
        convert_pdf_to_png_images "$FULL_PATH" "page"
        # Display results
        display_png_conversion_results_with_limit
        ;;

    *)
        echo "Error: Unsupported format: $EXTENSION"
        echo "Supported formats: .pptx, .pdf, .xlsx"
        exit "${EXIT_CODE_ERROR}"
        ;;
esac
