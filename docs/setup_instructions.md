# Email Processing Setup Instructions

## Overview
This document explains how to set up Outlook and Windows Task Scheduler to automatically export emails to files for R processing.

## Step 1: Outlook Rule Setup

### Create Email Processing Folders in Outlook
1. Right-click on your Inbox
2. Create new folders:
   - `PCIT-ToProcess` (for incoming emails)
   - `PCIT-Processed` (for completed emails)

### Create Outlook Rule
1. **File → Manage Rules & Alerts → New Rule**
2. **Start from a blank rule → Apply rule on messages I receive → Next**
3. **Conditions page**: 
   - Check "with specific words in the subject"
   - Click "specific words" → add "[PCIT Intermediary]" → Next
4. **Actions page**: 
   - Check "move it to the specified folder"
   - Click "specified folder" → select "PCIT-ToProcess" → Next
5. **Exceptions page**: Skip → Next
6. **Finish**: Name the rule "PCIT Email Processing" → Finish

## Step 2: Create Directory Structure

Create the following folders on your Windows machine:
```
C:\EmailProcessing\
├── scripts\
│   └── export-emails.ps1
├── inbox\          (where .msg files get saved)
├── processed\      (where R moves completed files)
└── logs\           (optional: for troubleshooting)
```

## Step 3: PowerShell Export Script

Create file: `C:\EmailProcessing\scripts\export-emails.ps1`

**Full PowerShell Script:**
```powershell
# Export emails from Outlook folder to .msg files
$outlook = New-Object -ComObject Outlook.Application
$folder = $outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders("PCIT-ToProcess")

foreach($mail in $folder.Items) {
    $filename = "C:\EmailProcessing\inbox\$($mail.ReceivedTime.ToString('yyyyMMdd_HHmmss')).msg"
    $mail.SaveAs($filename, 3)  # 3 = MSG format
    $mail.Move($outlook.GetNamespace("MAPI").GetDefaultFolder(6).Folders("PCIT-Processed"))
}
```

**Note:** A copy of this script is also saved in the project root as `export-emails.ps1` for reference.

## Step 4: Windows Task Scheduler Setup

### Create Scheduled Task
1. **Open Task Scheduler** (Windows search → "Task Scheduler")
2. **Create Basic Task**
   - Name: "Export PCIT Emails"
   - Description: "Exports PCIT emails from Outlook to file system for R processing"

### Configure Trigger
3. **Trigger**: Daily
   - Start date: Today
   - Recur every: 1 days
   - Click "Advanced settings"
   - Check "Repeat task every": 5 minutes
   - For a duration of: Indefinitely

### Configure Action
4. **Action**: Start a program
   - Program/script: `powershell.exe`
   - Add arguments: `-ExecutionPolicy Bypass -File "C:\EmailProcessing\scripts\export-emails.ps1"`
   - Start in: `C:\EmailProcessing\scripts\`

### Finish Setup
5. **Finish** → Check "Open the Properties dialog"
6. **Security options**: 
   - Check "Run whether user is logged on or not"
   - Check "Run with highest privileges"

## Step 5: Testing

### Test PowerShell Script
1. Send a test email with "[PCIT Intermediary]" in the subject
2. Manually run the PowerShell script: `powershell -ExecutionPolicy Bypass -File "C:\EmailProcessing\scripts\export-emails.ps1"`
3. Check if .msg file appears in `C:\EmailProcessing\inbox\`

### Test Task Scheduler
1. Right-click your task → "Run"
2. Check Task History for any errors
3. Verify emails are being exported every 5 minutes

## Troubleshooting

### Common Issues
- **Outlook not running**: The PowerShell script requires Outlook to be open
- **Permissions**: Ensure the task runs with appropriate user permissions
- **Folder paths**: Verify all folder paths exist and are accessible

### Log Files
Consider adding logging to the PowerShell script for debugging:
```powershell
Add-Content -Path "C:\EmailProcessing\logs\export-log.txt" -Value "$(Get-Date): Processing started"
```