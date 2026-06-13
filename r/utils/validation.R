btk_require_file <- function(path, label = "file") {
  if (!file.exists(path)) {
    stop(sprintf("Required %s does not exist: %s", label, path), call. = FALSE)
  }
  invisible(TRUE)
}

btk_require_packages <- function(packages) {
  missing <- packages[!vapply(packages, requireNamespace, logical(1), quietly = TRUE)]
  if (length(missing) > 0) {
    stop(
      sprintf("Missing required R package(s): %s", paste(missing, collapse = ", ")),
      call. = FALSE
    )
  }
  invisible(TRUE)
}

btk_require_columns <- function(data, columns, label = "data") {
  missing <- setdiff(columns, names(data))
  if (length(missing) > 0) {
    stop(
      sprintf("%s is missing required column(s): %s", label, paste(missing, collapse = ", ")),
      call. = FALSE
    )
  }
  invisible(TRUE)
}
