#!/bin/bash

# Log directory and file
LOG_DIR="/logs"
LOG_FILE="$LOG_DIR/batch_$(date +%Y%m%d_%H%M%S).log"
FAILED_LOG="$LOG_DIR/batch-failed_$(date +%Y%m%d_%H%M%S).log"

# Maximum parallel jobs
MAX_PARALLEL=5

# Function to log with timestamp
log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $1" | tee -a "$LOG_FILE"
}

# Function to execute a command and log output
execute_command() {
    local cmd="$1"
    local temp_log=$(mktemp)
    
    # Log start
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] Executing: $cmd" >> "$LOG_FILE"
    
    # Execute command and capture output
    eval "$cmd" > "$temp_log" 2>&1
    local exit_code=$?
    
    # Append output to main log
    cat "$temp_log" >> "$LOG_FILE"
    
    # Log completion status
    if [ $exit_code -eq 0 ]; then
        echo "[$(date '+%Y-%m-%d %H:%M:%S')] Completed successfully: $cmd" >> "$LOG_FILE"
    else
        echo "[$(date '+%Y-%m-%d %H:%M:%S')] Error occurred (exit code: $exit_code): $cmd" >> "$LOG_FILE"
        # Save failed command
        echo "$cmd" >> "$FAILED_LOG"
    fi
    
    # Clean up temp file
    rm -f "$temp_log"
    
    return $exit_code
}

# Start logging
log "=== Batch execution started ==="
log "Maximum parallel jobs: $MAX_PARALLEL"

# Read all commands into an array
commands=()
while IFS= read -r line; do
    # Skip empty lines and comment lines
    if [[ -z "$line" ]] || [[ "$line" =~ ^[[:space:]]*# ]]; then
        continue
    fi
    commands+=("$line")
done < /scripts/batch.txt

# Total number of commands
total_commands=${#commands[@]}
log "Total commands to execute: $total_commands"

# Execute commands with parallel processing
job_count=0
for i in "${!commands[@]}"; do
    # Start command in background
    execute_command "${commands[$i]}" &
    
    # Increment job counter
    ((job_count++))
    
    # If we've reached MAX_PARALLEL jobs, wait for one to finish
    if [ $job_count -ge $MAX_PARALLEL ]; then
        wait -n  # Wait for any background job to finish
        ((job_count--))
    fi
done

# Wait for all remaining jobs to complete
wait

log "=== All commands have been executed ==="
log "Log saved to: $LOG_FILE"