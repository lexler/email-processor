#!/usr/bin/env bash
set -euo pipefail

echo "Resetting email processor state..."

# Delete CSV file if it exists
if [ -f "emails.csv" ]; then
    rm emails.csv
    echo "Deleted emails.csv"
else
    echo "No emails.csv file found"
fi

# Move files from processed back to sample_emails
PROCESSED_DIR="../sample_emails/processed"
SAMPLE_DIR="../sample_emails"

if [ -d "$PROCESSED_DIR" ]; then
    # Move all .msg files from processed back to sample_emails
    if ls "$PROCESSED_DIR"/*.msg 1> /dev/null 2>&1; then
        mv "$PROCESSED_DIR"/*.msg "$SAMPLE_DIR/"
        echo "Moved .msg files from processed back to sample_emails"
        
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