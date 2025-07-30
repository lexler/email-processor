#!/usr/bin/env python3

import os
import glob
from pathlib import Path
import extract_msg
from datetime import datetime


def get_msg_files(inbox_path):
    """Get all .msg files from the specified directory"""
    if not os.path.exists(inbox_path):
        print(f"Error: Directory {inbox_path} does not exist")
        return None
    
    pattern = os.path.join(inbox_path, "*.msg")
    msg_files = glob.glob(pattern)
    return msg_files


def parse_message_content(body):
    """Parse message content and return a map of question responses"""
    questionnaire = None
    homework = None
    coding_analysis = None
    
    if not body:
        return {
            "questionnaire": questionnaire,
            "homework": homework, 
            "coding_analysis": coding_analysis
        }
    
    lines = body.split('\n')
    for line in lines:
        line = line.strip().lower()
        if 'questionnaire:' in line:
            questionnaire = 'yes' in line
        if 'asked about homework:' in line:
            homework = 'yes' in line
        if 'did coding analysis:' in line:
            coding_analysis = 'yes' in line
    
    return {
        "questionnaire": questionnaire,
        "homework": homework,
        "coding_analysis": coding_analysis
    }


def process_single_email(file_path):
    """Process a single email message and extract basic info"""
    try:
        msg = extract_msg.Message(file_path)
        
        sender = msg.sender or "unknown"
        received_time = msg.date or datetime.fromtimestamp(os.path.getmtime(file_path))
        body = msg.body or ""
        
        parsed_data = parse_message_content(body)
        
        print(f"File: {os.path.basename(file_path)}")
        print(f"From: {sender}")
        print(f"Date: {received_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"Questionnaire: {parsed_data['questionnaire']}")
        print(f"Asked about homework: {parsed_data['homework']}")
        print(f"Did coding analysis: {parsed_data['coding_analysis']}")
        print("---")
        
        msg.close()
        return True
        
    except Exception as e:
        print(f"ERROR processing {os.path.basename(file_path)}: {e}")
        return False


def main():
    inbox_path = "../sample_emails"
    
    msg_files = get_msg_files(inbox_path)
    if msg_files is None:
        return 1
    
    msg_count = len(msg_files)
    print(f"Found {msg_count} .msg files in {inbox_path}\n")
    
    for file_path in msg_files:
        process_single_email(file_path)
    
    return 0


if __name__ == "__main__":
    exit(main())