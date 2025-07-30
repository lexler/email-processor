#!/usr/bin/env bash
set -euo pipefail

echo "Running email processor..."
uv run process_emails.py /Users/ladak/workspace/hackathon/email_data_processing/sample_emails