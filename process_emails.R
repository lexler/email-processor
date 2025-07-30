library(RDCOMClient)
library(DBI)
library(RSQLite)

# Configuration
INBOX_PATH <- "C:/EmailProcessing/inbox"
PROCESSED_PATH <- "C:/EmailProcessing/processed"
DB_PATH <- "C:/EmailProcessing/pcit_data.sqlite"

# Create directories if they don't exist
if (!dir.exists(PROCESSED_PATH)) {
  dir.create(PROCESSED_PATH, recursive = TRUE)
}

# Initialize SQLite database
init_database <- function() {
  con <- dbConnect(RSQLite::SQLite(), DB_PATH)
  
  # Create tables if they don't exist
  dbExecute(con, "
    CREATE TABLE IF NOT EXISTS email_data (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      clinic_email TEXT,
      processed_date TEXT,
      questionnaire TEXT,
      asked_homework TEXT,
      coding_analysis TEXT,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )
  ")
  
  dbExecute(con, "
    CREATE TABLE IF NOT EXISTS clinic_lookup (
      email TEXT PRIMARY KEY,
      clinic_name TEXT
    )
  ")
  
  dbDisconnect(con)
}

# Parse email body for key-value pairs
parse_email_body <- function(body_text) {
  # Extract key-value pairs from body
  questionnaire <- if(grepl("Questionnaire:\\s*yes", body_text, ignore.case = TRUE)) "yes" else "no"
  homework <- if(grepl("Asked about homework:\\s*yes", body_text, ignore.case = TRUE)) "yes" else "no"
  coding <- if(grepl("Did coding analysis:\\s*yes", body_text, ignore.case = TRUE)) "yes" else "no"
  
  return(list(
    questionnaire = questionnaire,
    asked_homework = homework,
    coding_analysis = coding
  ))
}

# Process a single .msg file
process_msg_file <- function(file_path) {
  tryCatch({
    # Create Outlook application object
    outlook <- COMCreate("Outlook.Application")
    
    # Open the .msg file
    msg <- outlook$CreateItemFromTemplate(file_path)
    
    # Extract email data
    sender <- msg[["SenderEmailAddress"]]
    subject <- msg[["Subject"]]
    body <- msg[["Body"]]
    received_time <- msg[["ReceivedTime"]]
    
    cat("Processing email from:", sender, "\n")
    
    # Parse body content
    parsed_data <- parse_email_body(body)
    
    # Connect to database
    con <- dbConnect(RSQLite::SQLite(), DB_PATH)
    
    # Check if sender is in clinic lookup
    clinic_exists <- dbGetQuery(con, "SELECT COUNT(*) as count FROM clinic_lookup WHERE email = ?", 
                               params = list(sender))$count > 0
    
    if (clinic_exists) {
      # Insert data
      dbExecute(con, "
        INSERT INTO email_data (clinic_email, processed_date, questionnaire, asked_homework, coding_analysis)
        VALUES (?, ?, ?, ?, ?)
      ", params = list(
        sender,
        as.character(Sys.Date()),
        parsed_data$questionnaire,
        parsed_data$asked_homework,
        parsed_data$coding_analysis
      ))
      
      cat("✓ Data inserted successfully\n")
    } else {
      cat("⚠ Warning: Sender", sender, "not found in clinic lookup table\n")
    }
    
    dbDisconnect(con)
    
    # Move processed file
    file_name <- basename(file_path)
    processed_file <- file.path(PROCESSED_PATH, file_name)
    file.rename(file_path, processed_file)
    
    cat("✓ File moved to processed folder\n")
    
  }, error = function(e) {
    cat("✗ Error processing", file_path, ":", e$message, "\n")
  })
}

# Main processing function
process_emails <- function() {
  cat("Starting email processing...\n")
  
  # Get all .msg files in inbox
  msg_files <- list.files(INBOX_PATH, pattern = "\\.msg$", full.names = TRUE)
  
  if (length(msg_files) == 0) {
    cat("No .msg files found in inbox\n")
    return()
  }
  
  cat("Found", length(msg_files), ".msg files to process\n")
  
  # Process each file
  for (file_path in msg_files) {
    cat("\n--- Processing:", basename(file_path), "---\n")
    process_msg_file(file_path)
  }
  
  cat("\nEmail processing completed!\n")
}

# Initialize database on first run
init_database()

# Run the email processing
process_emails()