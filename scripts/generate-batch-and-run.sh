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
> /app/batch.txt

# Extensions to process
EXTENSIONS=("pptx" "pdf" "xlsx")

# Generate commands for each extension
for ext in "${EXTENSIONS[@]}"; do
    log "Generating commands for .$ext files..."
    
    # Find all files with the extension and generate commands
    find "$TARGET_DIR" -type f -name "*.$ext" | while read -r file; do
        # Remove /docs/ prefix to make it relative
        relative_path="${file#/docs/}"
        echo "bash /app/convert.sh \"$relative_path\";" >> /app/batch.txt
    done
    
    # Count files found
    file_count=$(find "$TARGET_DIR" -type f -name "*.$ext" | wc -l)
    log "Found $file_count .$ext files"
done

# Total commands generated
total_commands=$(grep -c "bash /app/convert.sh" /app/batch.txt 2>/dev/null || echo "0")
log "Total commands generated: $total_commands"

if [ "$total_commands" -eq 0 ]; then
    log "No files found to process. Exiting."
    exit 0
fi

# Display batch.txt content
log "Generated batch.txt content:"
cat /app/batch.txt | tee -a "$SCRIPT_LOG"

# Execute batch.sh
log "Starting batch.sh execution..."
bash /app/batch.sh

# Check exit status
if [ $? -eq 0 ]; then
    log "Batch execution completed successfully"
else
    log "Batch execution completed with errors"
fi

log "=== Batch generation and execution completed ==="
log "Script log saved to: $SCRIPT_LOG"