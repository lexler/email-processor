#!/usr/bin/env python3

import os
import glob
import csv
import shutil
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


def write_to_csv(file_name, sender, date, questionnaire, homework, coding_analysis, raw_body, csv_file="emails.csv"):
    """Write email data to CSV file, creating it if it doesn't exist"""
    file_exists = os.path.isfile(csv_file)
    
    with open(csv_file, 'a', newline='', encoding='utf-8') as csvfile:
        fieldnames = ['file_name', 'sender', 'date', 'questionnaire', 'homework', 'coding_analysis', 'raw_body']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        if not file_exists:
            writer.writeheader()
            print(f"Created new CSV file: {csv_file}")
        
        writer.writerow({
            'file_name': file_name,
            'sender': sender,
            'date': date,
            'questionnaire': questionnaire,
            'homework': homework,
            'coding_analysis': coding_analysis,
            'raw_body': raw_body.replace('\n', ' | ') if raw_body else ''
        })
        print(f"Appended to CSV file: {csv_file}")


def ensure_processed_directory(inbox_path):
    """Create processed directory if it doesn't exist"""
    processed_dir = os.path.join(inbox_path, "processed")
    if not os.path.exists(processed_dir):
        os.makedirs(processed_dir)
        print(f"Created processed directory: {processed_dir}")
    return processed_dir


def move_processed_file(file_path, processed_dir):
    """Move processed file to the processed directory"""
    filename = os.path.basename(file_path)
    new_path = os.path.join(processed_dir, filename)
    shutil.move(file_path, new_path)
    print(f"Moved file to: {new_path}")
    return new_path


def process_single_email(file_path, inbox_path):
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
        print(f"Message content: {body[:200]}..." if len(body) > 200 else f"Message content: {body}")
        
        # Write to CSV
        write_to_csv(
            file_name=os.path.basename(file_path),
            sender=sender,
            date=received_time.strftime('%Y-%m-%d %H:%M:%S'),
            questionnaire=parsed_data['questionnaire'],
            homework=parsed_data['homework'],
            coding_analysis=parsed_data['coding_analysis'],
            raw_body=body
        )
        
        # Move file to processed directory
        processed_dir = ensure_processed_directory(inbox_path)
        move_processed_file(file_path, processed_dir)
        
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
        process_single_email(file_path, inbox_path)
    
    return 0


if __name__ == "__main__":
    exit(main())