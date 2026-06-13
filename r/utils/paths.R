btk_path <- function(...) {
  file.path(BTK_ROOT, ...)
}

btk_workbook_dir <- function(workbook_path) {
  if (is.null(workbook_path) || identical(workbook_path, "")) {
    stop("workbook_path is required.", call. = FALSE)
  }

  normalizePath(dirname(workbook_path), winslash = "/", mustWork = TRUE)
}

btk_default_run_dir <- function(tool_id, workbook_path) {
  file.path(
    btk_workbook_dir(workbook_path),
    "_BERTToolkitTemp",
    tool_id
  )
}

btk_normalize_path <- function(path) {
  normalizePath(path, winslash = "/", mustWork = FALSE)
}

btk_build_context <- function(tool_id, workbook_path, output_dir = NULL, create_output_dir = TRUE) {
  workbook_path <- btk_normalize_path(workbook_path)

  if (is.null(output_dir) || identical(output_dir, "")) {
    output_dir <- btk_default_run_dir(tool_id, workbook_path)
  }

  output_dir <- btk_normalize_path(output_dir)
  if (isTRUE(create_output_dir)) {
    dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)
  }

  list(
    tool_id = tool_id,
    toolkit_root = BTK_ROOT,
    workbook_path = workbook_path,
    workbook_dir = btk_workbook_dir(workbook_path),
    output_dir = output_dir,
    log_file = file.path(output_dir, "run_log.txt")
  )
}
