#!/usr/bin/env bash
set -euo pipefail

if [ $# -eq 0 ]; then
    echo "Usage: $0 <email_directory>"
    echo "Example: $0 ./sample_emails"
    exit 1
fi

EMAIL_DIR="$1"

if [ ! -d "$EMAIL_DIR" ]; then
    echo "Error: Directory '$EMAIL_DIR' does not exist"
    exit 1
fi

echo "Processing emails from directory: $EMAIL_DIR"
Rscript read_emails.R "$EMAIL_DIR"