#!/usr/bin/env Rscript

library(RDCOMClient)

INBOX_PATH <- "/Users/ladak/workspace/hackathon/email_data_processing/PCIT Automated"

read_msg_file <- function(file_path) {
  tryCatch({
    outlook <- COMCreate("Outlook.Application")
    msg <- outlook$CreateItemFromTemplate(file_path)
    
    sender <- msg[["SenderEmailAddress"]]
    subject <- msg[["Subject"]]
    body <- msg[["Body"]]
    
    cat("=== Email ===\n")
    cat("From:", sender, "\n")
    cat("Subject:", subject, "\n")
    cat("Body:\n", body, "\n\n")
    
  }, error = function(e) {
    cat("Error reading", basename(file_path), ":", e$message, "\n")
  })
}

msg_files <- list.files(INBOX_PATH, pattern = "\\.msg$", full.names = TRUE)

if (length(msg_files) == 0) {
  cat("No .msg files found in inbox\n")
} else {
  for (file_path in msg_files) {
    read_msg_file(file_path)
  }
}