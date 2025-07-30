#!/usr/bin/env Rscript

options(repos = c(CRAN = "https://cran.rstudio.com/"))

if (!require(msgxtractr, quietly = TRUE)) {
  cat("Installing msgxtractr package...\n")
  if (!require(devtools, quietly = TRUE)) {
    install.packages("devtools")
  }
  devtools::install_github("hrbrmstr/msgxtractr")
  library(msgxtractr)
}

args <- commandArgs(trailingOnly = TRUE)

if (length(args) == 0) {
  cat("Usage: Rscript read_emails.R <email_directory>\n")
  quit(status = 1)
}

INBOX_PATH <- args[1]
CSV_FILE <- "emails.csv"

parse_email_data <- function(body) {
  questionnaire <- NA
  homework <- NA
  coding_analysis <- NA
  
  lines <- strsplit(body, "\n")[[1]]
  for (line in lines) {
    line <- trimws(line)
    if (grepl("Questionnaire:", line, ignore.case = TRUE)) {
      questionnaire <- ifelse(grepl("yes", line, ignore.case = TRUE), "yes", "no")
    }
    if (grepl("Asked about homework:", line, ignore.case = TRUE)) {
      homework <- ifelse(grepl("yes", line, ignore.case = TRUE), "yes", "no")
    }
    if (grepl("Did coding analysis:", line, ignore.case = TRUE)) {
      coding_analysis <- ifelse(grepl("yes", line, ignore.case = TRUE), "yes", "no")
    }
  }
  
  return(list(questionnaire = questionnaire, homework = homework, coding_analysis = coding_analysis))
}

process_msg_file <- function(file_path) {
  tryCatch({
    msg <- read_msg(file_path)
    
    sender <- if (!is.null(msg$sender$sender_email)) msg$sender$sender_email else "unknown"
    subject <- if (!is.null(msg$subject)) msg$subject else "no subject"
    body <- if (!is.null(msg$body$text)) msg$body$text else if (!is.null(msg$body$html)) msg$body$html else ""
    
    received_time <- if (!is.null(msg$headers$date_received)) {
      as.POSIXct(msg$headers$date_received)
    } else if (!is.null(msg$headers$date)) {
      as.POSIXct(msg$headers$date)
    } else {
      file.info(file_path)$mtime
    }
    
    parsed_data <- parse_email_data(body)
    
    cat("=== Processing Email ===\n")
    cat("File:", basename(file_path), "\n")
    cat("From:", sender, "\n")
    cat("Date:", format(received_time, "%Y-%m-%d %H:%M:%S"), "\n")
    cat("Subject:", subject, "\n")
    cat("Questionnaire:", parsed_data$questionnaire, "\n")
    cat("Asked about homework:", parsed_data$homework, "\n")
    cat("Did coding analysis:", parsed_data$coding_analysis, "\n")
    cat("Raw body:\n", body, "\n")
    cat("========================\n\n")
    
    csv_data <- data.frame(
      file_name = basename(file_path),
      sender = sender,
      date = format(received_time, "%Y-%m-%d %H:%M:%S"),
      subject = subject,
      questionnaire = parsed_data$questionnaire,
      homework = parsed_data$homework,
      coding_analysis = parsed_data$coding_analysis,
      raw_body = gsub("\n", " | ", body),
      stringsAsFactors = FALSE
    )
    
    if (!file.exists(CSV_FILE)) {
      write.csv(csv_data, CSV_FILE, row.names = FALSE)
      cat("Created new CSV file:", CSV_FILE, "\n")
    } else {
      write.table(csv_data, CSV_FILE, append = TRUE, sep = ",", row.names = FALSE, col.names = FALSE)
      cat("Appended to CSV file:", CSV_FILE, "\n")
    }
    
    processed_dir <- file.path(dirname(file_path), "processed")
    if (!dir.exists(processed_dir)) {
      dir.create(processed_dir, recursive = TRUE)
      cat("Created processed directory:", processed_dir, "\n")
    }
    
    new_path <- file.path(processed_dir, basename(file_path))
    file.rename(file_path, new_path)
    cat("Moved file to:", new_path, "\n\n")
    
    return(TRUE)
    
  }, error = function(e) {
    cat("ERROR processing", basename(file_path), ":", e$message, "\n")
    return(FALSE)
  })
}

if (!dir.exists(INBOX_PATH)) {
  cat("Error: Directory", INBOX_PATH, "does not exist\n")
  quit(status = 1)
}

msg_files <- list.files(INBOX_PATH, pattern = "\\.msg$", full.names = TRUE)

if (length(msg_files) == 0) {
  cat("No .msg files found in", INBOX_PATH, "\n")
} else {
  cat("Found", length(msg_files), ".msg files to process\n\n")
  
  processed_count <- 0
  for (file_path in msg_files) {
    if (process_msg_file(file_path)) {
      processed_count <- processed_count + 1
    }
  }
  
  cat("Successfully processed", processed_count, "out of", length(msg_files), "files\n")
}