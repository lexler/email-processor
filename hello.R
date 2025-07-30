#!/usr/bin/env Rscript

get_msg_files <- function(inbox_path) {
  if (!dir.exists(inbox_path)) {
    cat("Error: Directory", inbox_path, "does not exist\n")
    return(NULL)
  }
  
  msg_files <- list.files(inbox_path, pattern = "\\.msg$", full.names = TRUE)
  return(msg_files)
}

INBOX_PATH <- "sample_emails"

msg_files <- get_msg_files(INBOX_PATH)
if (!is.null(msg_files)) {
  msg_count <- length(msg_files)
  cat("Found", msg_count, ".msg files in", INBOX_PATH, "\n")
} else {
  quit(status = 1)
}