# Email Processing

## Overview
This is a process written in the R programming language that is going to automatically monitor an email address for emails of a specific format and digest them into csv files.

## Tech Stack
- running on a Windows machine
- R programming language
- email is hosted on an outlook server
- task scheduler is triggered via windows task manager
- creating csv files

## Sample Ingestion Email
From: <clinitian@123.clinic>
Subj: [PCIT Intermediary]
Body:
```
Questionnaire: yes
Asked about homework: no
Did coding analysis: yes
```

## Features
Read emails
Show that email has been processed
Detect that it would be processed but missing backing data to make it be processed
Store data in csv files

## Backing Data
Need a backend csv file link an email address with a specific clinic name.

