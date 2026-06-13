btk_log_message <- function(..., log_file = NULL) {
  msg <- paste(format(Sys.time(), "%Y-%m-%d %H:%M:%S"), "-", paste(..., collapse = " "))
  message(msg)

  if (!is.null(log_file) && !identical(log_file, "")) {
    cat(msg, "\n", file = log_file, append = TRUE)
  }

  invisible(msg)
}
