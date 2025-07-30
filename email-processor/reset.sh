#!/usr/bin/env bash
set -euo pipefail

# Get base directory from parameter or default to ../sample_emails
BASE_DIR="${1:-../sample_emails}"
BASE_DIR=$(realpath "$BASE_DIR")

echo "Resetting email processor state in: $BASE_DIR"

# Delete CSV file if it exists in base directory
CSV_FILE="$BASE_DIR/emails.csv"
if [ -f "$CSV_FILE" ]; then
    rm "$CSV_FILE"
    echo "Deleted $CSV_FILE"
else
    echo "No emails.csv file found in $BASE_DIR"
fi

# Move files from processed back to base directory
PROCESSED_DIR="$BASE_DIR/processed"

if [ -d "$PROCESSED_DIR" ]; then
    # Move all .msg files from processed back to base directory
    if ls "$PROCESSED_DIR"/*.msg 1> /dev/null 2>&1; then
        mv "$PROCESSED_DIR"/*.msg "$BASE_DIR/"
        echo "Moved .msg files from processed back to base directory"
        
        # Remove the processed directory if it's empty
        if [ -z "$(ls -A "$PROCESSED_DIR")" ]; then
            rmdir "$PROCESSED_DIR"
            echo "Removed empty processed directory"
        fi
    else
        echo "No .msg files found in processed directory"
    fi
else
    echo "No processed directory found"
fi

echo "Reset complete!"