# Machine Setup Instructions

## Development Environment Setup

This document outlines the required software and dependencies for running the PCIT email processing system.

## Prerequisites

### macOS Setup

#### 1. Install Homebrew (if not already installed)
```bash
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
```

#### 2. Install R Programming Language
```bash
brew install r
```

If R is installed but not linked:
```bash
brew link r
```

Verify installation:
```bash
which R
which Rscript
```

Both commands should return paths like `/opt/homebrew/bin/R` and `/opt/homebrew/bin/Rscript`.

#### 3. R Package Dependencies

The script will automatically install required R packages:
- `msgxtractr` - For parsing .msg files directly
- `devtools` - For installing packages from GitHub

These are installed automatically when the script runs if not present.

### Windows Setup

#### 1. Install R
- Download R from [https://cran.r-project.org/](https://cran.r-project.org/)
- Install with default options
- Ensure R is added to PATH

#### 2. PowerShell Execution Policy
For the email export script to work:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### 3. Outlook Requirements
- Microsoft Outlook must be installed and configured
- The PowerShell script uses COM objects to interact with Outlook

## Project Dependencies

### R Libraries Used
- **msgxtractr**: Direct .msg file parsing without requiring Outlook
  - Repository: https://github.com/hrbrmstr/msgxtractr
  - Installation: Auto-installed via devtools from GitHub
- **devtools**: For GitHub package installation

### System Requirements
- **Disk Space**: At least 500MB for R installation and dependencies
- **Memory**: 4GB RAM recommended for processing large email batches
- **Network**: Internet connection required for initial R package installation

## Verification Steps

### Test R Installation
```bash
R --version
```

### Test Script Execution
```bash
./process_emails.sh ./sample_emails
```

This should:
1. Install required R packages (first run only)
2. Process .msg files in the specified directory
3. Create `emails.csv` with extracted data
4. Move processed files to `processed/` subdirectory

## Troubleshooting

### R Not Found
If you get "Rscript: command not found":
1. Verify R installation: `which R`
2. If missing, reinstall: `brew install r`
3. If installed but not linked: `brew link r`

### Permission Errors
On macOS, you may need to allow script execution:
```bash
chmod +x process_emails.sh
```

### Package Installation Failures
If R packages fail to install:
1. Check internet connection
2. Try manual installation in R console:
```r
install.packages("devtools")
devtools::install_github("hrbrmstr/msgxtractr")
```

### .msg File Reading Errors
If .msg files can't be parsed:
1. Ensure files are valid Outlook .msg format
2. Check file permissions
3. Verify files aren't corrupted

## Performance Notes

- Initial run will be slower due to package installation
- Subsequent runs should be much faster
- Processing time depends on email count and .msg file sizes
- Large attachments in emails will increase processing time

## Security Considerations

- The system processes email content which may contain sensitive information
- Ensure proper file permissions on the `emails.csv` output
- Consider encrypting the CSV file if it contains PII
- Regularly clean up processed files if disk space is limited