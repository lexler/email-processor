#!/usr/bin/env python3

import sys
import os

# Add the email-processor src to path
email_processor_path = os.path.join(os.path.dirname(__file__), 'email-processor', 'src')
sys.path.insert(0, email_processor_path)

from email_processor.main import main

if __name__ == "__main__":
    exit(main())