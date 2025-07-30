# Email Ingestion
1. When email is sent to Rachel with [PCIT Intermediary] subj
2. It gets automagically moved to folder `PCIT Automated`
3. There's a daily task in task scheduler to run a powershell script `save-emails.ps1`
4. The powershell script saves that email to an .msg file and moves the email to `PCIT-Processed` folder
5. In R process, ingest the message file and process it
6. Save to csv file

# Process diagram

```mermaid
flowchart TD
    A[Email sent to Rachel<br/>Subject: PCIT Intermediary] --> B[Outlook Rule Triggers]
    B --> C[Email moved to<br/>PCIT Automated folder]
    C --> D[Windows Task Scheduler<br/>runs once a day]
    D --> E[PowerShell Script<br/>save-emails.ps1]
    E --> F[Email saved as<br/>.msg file to disk]
    F --> G[Email moved to<br/>PCIT-Processed folder]
    G --> H[R Script monitors<br/>inbox folder]
    H --> I[Parse .msg file into csv format<br/>]
```