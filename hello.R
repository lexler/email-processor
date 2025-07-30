#!/usr/bin/env Rscript

INBOX_PATH <- "sample_emails"

if (!dir.exists(INBOX_PATH)) {
  cat("Error: Directory", INBOX_PATH, "does not exist\n")
  quit(status = 1)
}

msg_files <- list.files(INBOX_PATH, pattern = "\\.msg$", full.names = TRUE)
msg_count <- length(msg_files)

cat("Found", msg_count, ".msg files in", INBOX_PATH, "\n")