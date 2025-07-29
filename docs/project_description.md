# Email Processing

## Overview
This is a process written in the R programming language that is going to automatically monitor an email address for emails of a specific format and digest them into SQLite.

## Tech Stack
- running on a Windows machine
- R programming language
- email is hosted on an outlook server
- task scheduler is triggered via windows task manager
- interfacing with an SQLite database

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
Store data in DB

## Backing Data
Need a backend table to link an email address with a specific clinic name.

