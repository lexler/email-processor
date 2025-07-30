# Email Processor

A Python-based email processor that extracts questionnaire responses from .msg files and saves structured data to CSV.

## Features

- **Email Parsing**: Processes Microsoft Outlook .msg files using the `extract-msg` library
- **Content Analysis**: Extracts structured responses for:
  - Questionnaire completion (yes/no)
  - Homework discussion (yes/no)
  - Coding analysis performed (yes/no)
- **CSV Export**: Saves clean, structured data with sender, date, and extracted responses
- **File Management**: Automatically moves processed files to a `processed/` folder
- **Reset Functionality**: Easy reset script to restore original state for testing

## Project Structure

```
email_data_processing/
├── email-processor/           # Main Python project (uv-managed)
│   ├── process_emails.py     # Main processing script
│   ├── run.sh               # Execute processor
│   ├── reset.sh             # Reset processed files
│   ├── pyproject.toml       # uv project configuration
│   └── emails.csv           # Generated output (after processing)
├── sample_emails/            # Input directory
│   ├── *.msg               # Email files to process
│   └── processed/          # Moved files after processing
└── README.md
```

## Setup

### Prerequisites
- Python 3.12+
- [uv](https://docs.astral.sh/uv/) package manager

### Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/lexler/email-processor.git
   cd email-processor
   ```

2. Navigate to the email-processor directory:
   ```bash
   cd email-processor
   ```

3. Install dependencies using uv:
   ```bash
   uv sync
   ```

## Usage

### Processing Emails

Place your .msg files in the `sample_emails/` directory, then run:

```bash
./run.sh
```

This will:
- Process all .msg files in the `sample_emails/` directory
- Extract sender, date, and questionnaire responses
- Save structured data to `emails.csv`
- Move processed files to `sample_emails/processed/`

### Reset for Testing

To restore original state (move files back, delete CSV):

```bash
./reset.sh
```

## Output Format

The generated `emails.csv` contains:

| Column | Description |
|--------|-------------|
| sender | Email sender address |
| date | Email date/time (YYYY-MM-DD HH:MM:SS) |
| questionnaire | Boolean - questionnaire completed |
| homework | Boolean - homework discussed |
| coding_analysis | Boolean - coding analysis performed |

### Example Output

```csv
sender,date,questionnaire,homework,coding_analysis
Lada Kesseler <ladak@logic2020.com>,2025-07-29 15:52:45,False,False,True
Lada Kesseler <ladak@logic2020.com>,2025-07-29 15:23:34,True,False,True
```

## How It Works

1. **File Discovery**: Scans `sample_emails/` for .msg files
2. **Email Parsing**: Uses `extract-msg` to decode Outlook message files  
3. **Content Analysis**: Searches message body for specific patterns:
   - "Questionnaire: yes/no"
   - "Asked about homework: yes/no" 
   - "Did coding analysis: yes/no"
4. **Data Export**: Writes structured data to CSV
5. **File Management**: Moves processed files to avoid reprocessing

## Dependencies

- `extract-msg` - Microsoft Outlook .msg file parser
- Built-in Python libraries: `csv`, `glob`, `shutil`, `datetime`

## Development

This project uses [uv](https://docs.astral.sh/uv/) for dependency management and virtual environment handling.

## License

MIT License