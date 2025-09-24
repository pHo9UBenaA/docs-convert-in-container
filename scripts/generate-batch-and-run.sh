#!/bin/bash

# Script to generate batch.txt from all pptx, pdf, and xlsx files
# and automatically run batch.sh

# Directory to process (default: /docs)
if [ -n "$1" ]; then
    TARGET_DIR="/$1"
else
    TARGET_DIR="/docs"
fi

# Log file for this script
SCRIPT_LOG="/logs/generate-batch-and-run_$(date +%Y%m%d_%H%M%S).log"

# Function to log with timestamp
log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $1" | tee -a "$SCRIPT_LOG"
}

# Start logging
log "=== Starting batch generation and execution ==="
log "Target directory: $TARGET_DIR"

# Check if target directory exists
if [ ! -d "$TARGET_DIR" ]; then
    log "Error: Directory '$TARGET_DIR' does not exist"
    exit 1
fi

# Clear previous batch.txt if exists
> /scripts/batch.txt

# Collect all PPTX files to identify generated PDFs
declare -A pptx_base_paths
while IFS= read -r -d '' pptx_file; do
    base_path="${pptx_file%.pptx}"
    pptx_base_paths["$base_path"]=1
done < <(find "$TARGET_DIR" -type f -name "*.pptx" -print0)

# Process each file type
declare -A file_counts
for ext in pptx pdf xlsx; do
    count=0
    skipped=0

    while IFS= read -r -d '' file; do
        # Skip PDFs that were generated from PPTX files
        if [ "$ext" = "pdf" ]; then
            base_path="${file%.pdf}"
            if [ "${pptx_base_paths[$base_path]}" = "1" ]; then
                log "Skipping $file (generated from PPTX)"
                ((skipped++))
                continue
            fi
        fi

        # Add to batch.txt
        relative_path="${file#/docs/}"
        echo "bash /scripts/convert.sh \"$relative_path\";" >> /scripts/batch.txt
        ((count++))
    done < <(find "$TARGET_DIR" -type f -name "*.$ext" -print0)

    # Log results
    if [ "$ext" = "pdf" ] && [ "$skipped" -gt 0 ]; then
        log "Found $count .$ext files (excluded $skipped PPTX-generated PDFs)"
    else
        log "Found $count .$ext files"
    fi
done

# Total commands generated
total_commands=$(grep -c "bash /scripts/convert.sh" /scripts/batch.txt 2>/dev/null || echo "0")
log "Total commands generated: $total_commands"

if [ "$total_commands" -eq 0 ]; then
    log "No files found to process. Exiting."
    exit 0
fi

# Display batch.txt content
log "Generated batch.txt content:"
cat /scripts/batch.txt | tee -a "$SCRIPT_LOG"

# Execute batch.sh
log "Starting batch.sh execution..."
bash /scripts/batch.sh

# Check exit status
if [ $? -eq 0 ]; then
    log "Batch execution completed successfully"
else
    log "Batch execution completed with errors"
fi

log "=== Batch generation and execution completed ==="
log "Script log saved to: $SCRIPT_LOG"