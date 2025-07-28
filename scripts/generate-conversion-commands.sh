#!/bin/bash

# Display usage
if [ $# -eq 0 ]; then
    echo "Usage: $0 <extension>"
    echo "Example: $0 xlsx"
    echo "Example: $0 pdf"
    echo "Example: $0 pptx"
    echo ""
    echo "This script finds all files with the specified extension in the docs/ directory"
    echo "and outputs docker compose run commands to convert them."
    exit 1
fi

EXTENSION="$1"

# Find all files with the specified extension
# Store output in a variable to both display and copy
OUTPUT=$(find docs/ -type f -name "*.$EXTENSION" | while read -r file; do
    echo "bash scripts/convert.sh \"$file\";"
done)

# Display the output
echo "$OUTPUT"

# Copy to clipboard if pbcopy is available
if command -v pbcopy >/dev/null 2>&1; then
    echo "$OUTPUT" | pbcopy
    echo "Commands copied to clipboard!"
fi
