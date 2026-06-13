# Single BERT startup entry point.
# Source this file from BERT's functions.R.

btk_bootstrap_file <- function() {
  for (i in rev(seq_len(sys.nframe()))) {
    ofile <- sys.frame(i)$ofile
    if (!is.null(ofile)) {
      return(normalizePath(ofile, winslash = "/", mustWork = FALSE))
    }
  }

  NA_character_
}

btk_file <- btk_bootstrap_file()
btk_root_from_file <- if (!is.na(btk_file)) {
  normalizePath(file.path(dirname(btk_file), ".."), winslash = "/", mustWork = FALSE)
} else {
  NA_character_
}

BTK_ROOT <- Sys.getenv("BTK_ROOT", unset = btk_root_from_file)
if (is.na(BTK_ROOT) || !nzchar(BTK_ROOT)) {
  BTK_ROOT <- "C:/CompanyTools/BERTToolkit"
}

BTK_ROOT <- normalizePath(BTK_ROOT, winslash = "/", mustWork = FALSE)

source(file.path(BTK_ROOT, "r", "registry.R"))

BTK.Safe <- function(expr) {
  tryCatch(
    expr,
    error = function(e) {
      paste("ERROR:", conditionMessage(e))
    }
  )
}

BTK.Version <- function() {
  BTK.Safe(toolkit_version())
}

BTK.ListTools <- function() {
  BTK.Safe(paste(names(toolkit_registry()), collapse = ", "))
}

BTK.ValidateTool <- function(tool_id, workbook_path) {
  BTK.Safe(toolkit_validate(tool_id = tool_id, workbook_path = workbook_path))
}

BTK.RunTool <- function(tool_id, workbook_path, output_dir = NULL) {
  BTK.Safe(toolkit_run(tool_id = tool_id, workbook_path = workbook_path, output_dir = output_dir))
}

BTK.DispatchTool <- function(tool_id, action, workbook_path, output_dir = NULL) {
  BTK.Safe(toolkit_dispatch(tool_id = tool_id, action = action, workbook_path = workbook_path, output_dir = output_dir))
}

BTK.ToolActions <- function(tool_id) {
  BTK.Safe(toolkit_action_names(tool_id))
}

BTK.DefaultRunDir <- function(tool_id, workbook_path) {
  BTK.Safe(btk_default_run_dir(tool_id = tool_id, workbook_path = workbook_path))
}

BTK.RequiredPackages <- function() {
  c("sf", "ggplot2", "dplyr", "magick", "readxl", "openxlsx")
}

BTK.PackageRepository <- function() {
  "https://packagemanager.posit.co/cran/2019-06-01"
}

BTK.UserLibrary <- function() {
  lib <- Sys.getenv("R_LIBS_USER")
  if (!nzchar(lib)) {
    lib <- file.path(Sys.getenv("USERPROFILE"), "Documents", "R", "win-library", paste(R.version$major, R.version$minor, sep = "."))
  }
  dir.create(lib, recursive = TRUE, showWarnings = FALSE)
  .libPaths(unique(c(lib, .libPaths())))
  lib
}

BTK.MissingPackages <- function() {
  packages <- BTK.RequiredPackages()
  missing <- packages[!vapply(packages, requireNamespace, logical(1), quietly = TRUE)]
  if (length(missing) == 0) {
    "No missing packages."
  } else {
    paste(missing, collapse = ", ")
  }
}

BTK.InstallPackages <- function() {
  BTK.Safe({
    packages <- BTK.RequiredPackages()
    missing <- packages[!vapply(packages, requireNamespace, logical(1), quietly = TRUE)]
    if (length(missing) == 0) {
      return("No missing packages.")
    }

    lib <- BTK.UserLibrary()
    repos <- BTK.PackageRepository()

    install.packages(
      missing,
      lib = lib,
      repos = repos,
      type = "win.binary",
      dependencies = c("Depends", "Imports", "LinkingTo")
    )

    remaining <- missing[!vapply(missing, requireNamespace, logical(1), quietly = TRUE)]
    if (length(remaining) > 0) {
      paste(
        "ERROR: Still missing after install:",
        paste(remaining, collapse = ", "),
        "Repository:",
        repos,
        "Library:",
        lib
      )
    } else {
      paste("Installed packages:", paste(missing, collapse = ", "), "Library:", lib)
    }
  })
}
