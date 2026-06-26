xl_pricing_tool_rmstoyelt_gather <- function(workbook_path, output_dir = NULL, context = list()) {
  xl_pricing_tool_validate(workbook_path, context, action = "rmstoyelt_gather")

  rmstoyelt <- xl_pricing_tool_rmstoyelt_context(workbook_path, output_dir, context)
  sheet <- rmstoyelt$sheet
  xl_pricing_tool_rmstoyelt_check_marker(sheet)
  xl_pricing_tool_rmstoyelt_check_layout(sheet)

  input_folder <- xl_pricing_tool_rmstoyelt_normalize_path(xl_pricing_tool_cell_value(sheet, "C14"))
  output_folder <- xl_pricing_tool_rmstoyelt_normalize_path(xl_pricing_tool_cell_value(sheet, "D14"))
  declaration_file <- file.path(rmstoyelt$rmstoyelt_dir, "rmstoyelt_declarations.csv")

  if (!nzchar(input_folder)) {
    if (file.exists(declaration_file)) unlink(declaration_file)
    return(paste(
      "RMStoYELT gather completed.",
      sprintf("Output folder: %s", rmstoyelt$output_dir),
      sprintf("Active sheet: %s", sheet$name),
      "Bulk input folder C14 is blank; declarations were not changed.",
      sep = vb_newline()
    ))
  }
  if (!dir.exists(input_folder)) {
    stop(sprintf("RMStoYELT bulk input folder does not exist: %s", input_folder), call. = FALSE)
  }

  files <- list.files(input_folder, pattern = "\\.csv$", full.names = TRUE, ignore.case = TRUE)
  files <- files[file.info(files)$isdir %in% FALSE]
  if (length(files) == 0) {
    stop(sprintf("RMStoYELT gather found no CSV files in C14 folder: %s", input_folder), call. = FALSE)
  }

  rows <- lapply(files, function(path) {
    base <- basename(path)
    cause_id <- tools::file_path_sans_ext(base)
    data.frame(
      CauseID = cause_id,
      Peril = xl_pricing_tool_rmstoyelt_infer_peril(base),
      FilePath = normalizePath(path, winslash = "/", mustWork = FALSE),
      Active = "Y",
      OutDir = if (nzchar(output_folder)) output_folder else dirname(normalizePath(path, winslash = "/", mustWork = FALSE)),
      OutFileName = base,
      stringsAsFactors = FALSE
    )
  })
  declarations <- do.call(rbind, rows)
  declarations <- declarations[seq_len(min(nrow(declarations), 88L)), , drop = FALSE]

  utils::write.table(
    declarations,
    file = declaration_file,
    sep = ",",
    row.names = FALSE,
    col.names = FALSE,
    na = "",
    quote = TRUE
  )

  paste(
    "RMStoYELT gather completed.",
    sprintf("Output folder: %s", rmstoyelt$output_dir),
    sprintf("Active sheet: %s", sheet$name),
    sprintf("Bulk input folder: %s", input_folder),
    sprintf("Files listed: %s", nrow(declarations)),
    sprintf("Declaration output: %s", declaration_file),
    "CSV output is handled by the VBA client after BERT returns.",
    sep = vb_newline()
  )
}

xl_pricing_tool_rmstoyelt_update <- function(workbook_path, output_dir = NULL, context = list()) {
  xl_pricing_tool_validate(workbook_path, context, action = "rmstoyelt_update")

  rmstoyelt <- xl_pricing_tool_rmstoyelt_context(workbook_path, output_dir, context)
  sheet <- rmstoyelt$sheet
  xl_pricing_tool_rmstoyelt_check_marker(sheet)
  xl_pricing_tool_rmstoyelt_check_layout(sheet)

  simulation_years <- xl_pricing_tool_rmstoyelt_simulation_years(sheet)
  declarations <- xl_pricing_tool_rmstoyelt_read_declarations(sheet)
  active <- declarations[toupper(declarations$Active) %in% c("Y", "YES", "TRUE", "1"), , drop = FALSE]
  if (nrow(active) == 0) {
    stop("RMStoYELT update has no active declaration rows in B27:G114.", call. = FALSE)
  }

  summaries <- character()
  output_paths <- character()
  for (i in seq_len(nrow(active))) {
    row <- active[i, , drop = FALSE]
    job <- xl_pricing_tool_rmstoyelt_validate_job(row)
    elt <- xl_pricing_tool_rmstoyelt_read_rms_elt(job$FilePath)
    yelt <- xl_pricing_tool_rmstoyelt_simulate(elt, simulation_years)
    event_file <- xl_pricing_tool_rmstoyelt_ctr_ts_event_file(yelt)

    dir.create(job$OutDir, recursive = TRUE, showWarnings = FALSE)
    output_path <- file.path(job$OutDir, job$OutFileName)
    input_path <- normalizePath(job$FilePath, winslash = "/", mustWork = TRUE)
    output_norm <- normalizePath(output_path, winslash = "/", mustWork = FALSE)
    if (identical(tolower(input_path), tolower(output_norm))) {
      stop(sprintf("RMStoYELT output would overwrite source file for %s: %s", job$CauseID, output_path), call. = FALSE)
    }

    xl_pricing_tool_rmstoyelt_write_ctr_ts_event_file(event_file, output_path)
    output_paths <- c(output_paths, output_path)
    summaries <- c(summaries, sprintf(
      "%s: lambda=%.8f, input events=%s, simulated occurrence rows=%s, output=%s",
      job$CauseID,
      sum(elt$Rate),
      nrow(elt),
      nrow(yelt),
      output_path
    ))
  }

  paste(
    "RMStoYELT update completed.",
    sprintf("Output folder: %s", rmstoyelt$output_dir),
    sprintf("Active sheet: %s", sheet$name),
    sprintf("Simulation years: %s", simulation_years),
    sprintf("Active jobs: %s", nrow(active)),
    "Outputs:",
    paste(summaries, collapse = vb_newline()),
    "Generated CTR/TS-style event files: // escape-delay header plus Year, Loss, LossType, LossOccurrence, Delay, EventID rows.",
    sep = vb_newline()
  )
}

xl_pricing_tool_rmstoyelt_build <- function(workbook_path, output_dir = NULL, context = list()) {
  "No build action defined for RMStoYELT. Use gather to populate declarations and update to generate YELT files."
}

xl_pricing_tool_rmstoyelt_context <- function(workbook_path, output_dir = NULL, context = list()) {
  if (is.null(output_dir) || identical(output_dir, "")) {
    output_dir <- if (!is.null(context$output_dir)) {
      context$output_dir
    } else {
      btk_default_run_dir("xl_pricing_tool", workbook_path)
    }
  }

  rmstoyelt_dir <- file.path(output_dir, "rmstoyelt")
  dir.create(rmstoyelt_dir, recursive = TRUE, showWarnings = FALSE)

  list(
    workbook_path = workbook_path,
    output_dir = output_dir,
    rmstoyelt_dir = rmstoyelt_dir,
    sheet = xl_pricing_tool_action_sheet(workbook_path, context, "RMStoYELT")
  )
}

xl_pricing_tool_rmstoyelt_check_marker <- function(sheet) {
  marker <- xl_pricing_tool_normalize_marker(xl_pricing_tool_cell_value(sheet, "A1"))
  if (!marker %in% c("rmstoyelt", "rms toyelt", "rms to yelt")) {
    stop("RMStoYELT sheet must have <<RMStoYELT>> in A1.", call. = FALSE)
  }
  "RMStoYELT marker found in A1."
}

xl_pricing_tool_rmstoyelt_check_layout <- function(sheet) {
  expected <- c("CauseID", "Peril", "FilePath", "Active", "OutDir", "OutFileName")
  headers <- vapply(2:7, function(col) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), 26)))
  }, character(1))
  if (!all(headers == expected)) {
    stop("RMStoYELT declaration table must have expected headers in B26:G26.", call. = FALSE)
  }
  "RMStoYELT declaration layout valid."
}

xl_pricing_tool_rmstoyelt_simulation_years <- function(sheet) {
  value <- suppressWarnings(as.numeric(gsub(",", "", trimws(xl_pricing_tool_cell_value(sheet, "C6")), fixed = TRUE)))
  if (!is.finite(value) || value <= 0 || value != floor(value)) {
    stop("RMStoYELT Simulation Years in C6 must be a positive integer.", call. = FALSE)
  }
  as.integer(value)
}

xl_pricing_tool_rmstoyelt_read_declarations <- function(sheet) {
  rows <- list()
  for (row in 27:114) {
    values <- vapply(2:7, function(col) {
      trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), row)))
    }, character(1))
    if (!any(nzchar(values))) next
    rows[[length(rows) + 1L]] <- data.frame(
      CauseID = values[[1]],
      Peril = values[[2]],
      FilePath = xl_pricing_tool_rmstoyelt_normalize_path(values[[3]]),
      Active = values[[4]],
      OutDir = xl_pricing_tool_rmstoyelt_normalize_path(values[[5]]),
      OutFileName = values[[6]],
      stringsAsFactors = FALSE
    )
  }

  if (length(rows) == 0) {
    return(data.frame(
      CauseID = character(),
      Peril = character(),
      FilePath = character(),
      Active = character(),
      OutDir = character(),
      OutFileName = character(),
      stringsAsFactors = FALSE
    ))
  }
  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

xl_pricing_tool_rmstoyelt_validate_job <- function(row) {
  for (field in c("CauseID", "FilePath", "OutDir", "OutFileName")) {
    if (!nzchar(row[[field]][[1]])) {
      stop(sprintf("RMStoYELT active row is missing %s.", field), call. = FALSE)
    }
  }
  if (!file.exists(row$FilePath[[1]])) {
    stop(sprintf("RMStoYELT source file not found for %s: %s", row$CauseID[[1]], row$FilePath[[1]]), call. = FALSE)
  }
  if (grepl("[\\\\/:*?\"<>|]", row$OutFileName[[1]])) {
    stop(sprintf("RMStoYELT OutFileName for %s must be a file name only, not a path: %s", row$CauseID[[1]], row$OutFileName[[1]]), call. = FALSE)
  }
  if (!grepl("\\.csv$", row$OutFileName[[1]], ignore.case = TRUE)) {
    row$OutFileName[[1]] <- paste0(row$OutFileName[[1]], ".csv")
  }
  as.list(row[1, , drop = FALSE])
}

xl_pricing_tool_rmstoyelt_read_rms_elt <- function(path) {
  raw <- xl_pricing_tool_rmstoyelt_read_csv_rectangular(path)
  if (nrow(raw) == 0 || ncol(raw) < 6) {
    stop(sprintf("RMStoYELT RMS ELT is empty or has fewer than 6 columns: %s", path), call. = FALSE)
  }

  expected <- c("Event ID", "Event Rate", "Mean Loss", "Std Dev Ind", "Std Dev Corr", "Exposure")
  header_row <- NA_integer_
  for (i in seq_len(nrow(raw))) {
    normalized <- gsub("^Evend\\.ID$", "Event.ID", make.names(trimws(as.character(raw[i, seq_len(6)]))))
    if (all(normalized == c("Event.ID", "Event.Rate", "Mean.Loss", "Std.Dev.Ind", "Std.Dev.Corr", "Exposure"))) {
      header_row <- i
      break
    }
  }

  if (is.na(header_row)) {
    event_id <- trimws(as.character(raw[[1]]))
    numeric_rows <- suppressWarnings(as.numeric(event_id))
    valid <- which(nzchar(event_id) & is.finite(numeric_rows))
    if (length(valid) == 0) {
      stop(sprintf("RMStoYELT could not find RMS ELT headers or numeric event rows in: %s", path), call. = FALSE)
    }
    data <- raw[valid[[1]]:nrow(raw), seq_len(6), drop = FALSE]
  } else {
    data <- raw[(header_row + 1L):nrow(raw), seq_len(6), drop = FALSE]
  }
  names(data) <- expected
  data <- xl_pricing_tool_rmstoyelt_drop_blank_rows(data)

  out <- data.frame(
    EventID = trimws(as.character(data[["Event ID"]])),
    Rate = suppressWarnings(as.numeric(data[["Event Rate"]])),
    Mean = suppressWarnings(as.numeric(data[["Mean Loss"]])),
    Sdi = suppressWarnings(as.numeric(data[["Std Dev Ind"]])),
    Sdc = suppressWarnings(as.numeric(data[["Std Dev Corr"]])),
    Exposure = suppressWarnings(as.numeric(data[["Exposure"]])),
    stringsAsFactors = FALSE
  )
  out <- out[
    nzchar(out$EventID) &
      is.finite(out$Rate) & out$Rate > 0 &
      is.finite(out$Mean) & out$Mean > 0 &
      is.finite(out$Sdi) & out$Sdi >= 0 &
      is.finite(out$Sdc) & out$Sdc >= 0 &
      is.finite(out$Exposure) & out$Exposure > 0,
    ,
    drop = FALSE
  ]
  if (nrow(out) == 0) {
    stop(sprintf("RMStoYELT RMS ELT has no valid event rows after filtering: %s", path), call. = FALSE)
  }
  rownames(out) <- NULL
  out
}

xl_pricing_tool_rmstoyelt_simulate <- function(elt, simulation_years) {
  lambda <- sum(elt$Rate)
  if (!is.finite(lambda) || lambda <= 0) {
    stop("RMStoYELT event rates sum to zero.", call. = FALSE)
  }

  counts <- stats::rpois(simulation_years, lambda = lambda)
  total <- sum(counts)
  if (total == 0) {
    return(data.frame(
      Year = integer(),
      EventID = character(),
      Loss = numeric(),
      stringsAsFactors = FALSE
    ))
  }

  selected <- sample.int(nrow(elt), size = total, replace = TRUE, prob = elt$Rate)
  selected_elt <- elt[selected, , drop = FALSE]
  losses <- xl_pricing_tool_rmstoyelt_beta_losses(selected_elt)

  data.frame(
    Year = rep(seq_len(simulation_years), counts),
    EventID = selected_elt$EventID,
    Loss = losses,
    stringsAsFactors = FALSE
  )
}

xl_pricing_tool_rmstoyelt_ctr_ts_event_file <- function(events) {
  if (nrow(events) == 0) {
    return(data.frame(
      Year = integer(),
      Loss = numeric(),
      LossType = character(),
      LossOccurrence = numeric(),
      Delay = numeric(),
      EventID = character(),
      stringsAsFactors = FALSE
    ))
  }

  data.frame(
    Year = events$Year,
    Loss = events$Loss,
    LossType = "catastrophic",
    LossOccurrence = 0,
    Delay = stats::runif(nrow(events)),
    EventID = events$EventID,
    stringsAsFactors = FALSE
  )
}

xl_pricing_tool_rmstoyelt_write_ctr_ts_event_file <- function(event_file, output_path) {
  con <- file(output_path, open = "wt")
  on.exit(close(con), add = TRUE)
  writeLines("// escape-delay,,,,,", con = con)
  if (nrow(event_file) == 0) {
    return(invisible(output_path))
  }

  utils::write.table(
    event_file,
    file = con,
    sep = ",",
    row.names = FALSE,
    col.names = FALSE,
    quote = FALSE,
    na = ""
  )
  invisible(output_path)
}

xl_pricing_tool_rmstoyelt_beta_losses <- function(elt) {
  mean_loss <- pmax(elt$Mean, 0)
  exposure <- pmax(elt$Exposure, mean_loss)
  sd_loss <- pmax(elt$Sdi + elt$Sdc, 0)
  losses <- mean_loss

  variable <- sd_loss > 0 & exposure > mean_loss
  if (!any(variable)) {
    return(pmin(losses, exposure))
  }

  idx <- which(variable)
  m <- mean_loss[idx]
  e <- exposure[idx]
  s <- sd_loss[idx]
  max_sd <- sqrt(pmax(m * (e - m), 0))
  s <- pmin(s, max_sd * 0.999999)

  invalid <- !is.finite(s) | s <= 0 | !is.finite(max_sd) | max_sd <= 0
  valid_idx <- idx[!invalid]
  if (length(valid_idx) == 0) {
    return(pmin(losses, exposure))
  }

  m <- mean_loss[valid_idx]
  e <- exposure[valid_idx]
  s <- pmin(sd_loss[valid_idx], sqrt(m * (e - m)) * 0.999999)
  alpha <- (m / s) ^ 2 * (1 - m / e) - (m / e)
  beta <- alpha * (e / m - 1)
  good <- is.finite(alpha) & is.finite(beta) & alpha > 0 & beta > 0
  if (any(good)) {
    target <- valid_idx[good]
    losses[target] <- stats::qbeta(stats::runif(length(target)), alpha[good], beta[good]) * exposure[target]
  }

  pmin(pmax(losses, 0), exposure)
}

xl_pricing_tool_rmstoyelt_read_csv_rectangular <- function(csv_path) {
  lines <- readLines(csv_path, warn = FALSE)
  lines <- sub("\r$", "", lines)
  if (length(lines) == 0) return(data.frame())

  counter <- textConnection(lines)
  on.exit(close(counter), add = TRUE)
  field_counts <- utils::count.fields(counter, sep = ",", quote = "\"", blank.lines.skip = FALSE, comment.char = "")
  max_cols <- max(field_counts, na.rm = TRUE)
  if (!is.finite(max_cols) || max_cols < 1) return(data.frame())

  reader <- textConnection(lines)
  on.exit(close(reader), add = TRUE)
  utils::read.table(
    reader,
    sep = ",",
    header = FALSE,
    stringsAsFactors = FALSE,
    check.names = FALSE,
    colClasses = "character",
    na.strings = character(),
    blank.lines.skip = FALSE,
    fill = TRUE,
    quote = "\"",
    comment.char = "",
    col.names = paste0("V", seq_len(max_cols))
  )
}

xl_pricing_tool_rmstoyelt_drop_blank_rows <- function(data) {
  if (nrow(data) == 0) return(data)
  keep <- apply(data, 1, function(row) any(nzchar(trimws(as.character(row)))))
  data[keep, , drop = FALSE]
}

xl_pricing_tool_rmstoyelt_infer_peril <- function(file_name) {
  upper <- toupper(file_name)
  if (grepl("(^|[_ .-])EQ([_ .-]|$)", upper)) return("EQ")
  if (grepl("(^|[_ .-])(WS|WF)([_ .-]|$)", upper)) return("WS")
  ""
}

xl_pricing_tool_rmstoyelt_normalize_path <- function(value) {
  value <- trimws(as.character(value))
  if (!nzchar(value)) return("")
  normalizePath(value, winslash = "/", mustWork = FALSE)
}
