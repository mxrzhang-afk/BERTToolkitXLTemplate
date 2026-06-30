goecode_tool_validate <- function(workbook_path, context = list()) {
  btk_require_file(workbook_path, "workbook")

  sheets <- goecode_tool_sheet_names(workbook_path)
  if (!"Control" %in% sheets) {
    stop("Geocode workbook is missing required sheet: Control", call. = FALSE)
  }

  active_sheet <- goecode_tool_active_sheet(context, sheets)
  marker <- goecode_tool_cell(workbook_path, active_sheet, row = 1, col = 1)
  if (!identical(goecode_tool_marker_key(marker), "inputnorm")) {
    stop(sprintf("Active sheet '%s' is not an <<inputnorm>> tab.", active_sheet), call. = FALSE)
  }

  "OK"
}

goecode_tool_inputnorm_gather <- function(workbook_path, output_dir = NULL, context = list()) {
  ctx <- goecode_tool_context(workbook_path, output_dir, context)
  source_path <- goecode_tool_source_path(ctx$workbook_path, ctx$active_sheet)

  headers <- goecode_tool_source_headers(source_path)
  aliases <- goecode_tool_alias_dictionary(ctx$workbook_path)
  guesses <- goecode_tool_guess_headers(headers, aliases)

  out <- data.frame(
    header_raw = headers,
    header_mapped = guesses$header_mapped,
    Credibility = guesses$Credibility,
    stringsAsFactors = FALSE,
    check.names = FALSE
  )

  inputnorm_dir <- file.path(ctx$output_dir, "inputnorm")
  dir.create(inputnorm_dir, recursive = TRUE, showWarnings = FALSE)
  mapping_file <- file.path(inputnorm_dir, "inputnorm_header_mapping.csv")
  diagnostic_file <- file.path(inputnorm_dir, "inputnorm_gather_diagnostic.csv")
  goecode_tool_write_csv(out, mapping_file)
  goecode_tool_write_gather_diagnostic(
    source_path = source_path,
    headers = headers,
    diagnostic_file = diagnostic_file
  )

  paste(
    "InputNorm gather completed.",
    sprintf("Output folder: %s", ctx$output_dir),
    sprintf("Header mapping: %s", mapping_file),
    sprintf("Diagnostic: %s", diagnostic_file),
    sprintf("Headers read: %s", length(headers)),
    sep = vb_newline()
  )
}

goecode_tool_inputnorm_update <- function(workbook_path, output_dir = NULL, context = list()) {
  ctx <- goecode_tool_context(workbook_path, output_dir, context)
  source_path <- goecode_tool_source_path(ctx$workbook_path, ctx$active_sheet)
  mapping <- goecode_tool_read_header_mapping(ctx$workbook_path, ctx$active_sheet)
  if (nrow(mapping) == 0) {
    stop("No header mapping found in B26:C. Run Gather first, then configure C26:C.", call. = FALSE)
  }

  raw <- goecode_tool_read_source_table(source_path)
  goecode_tool_validate_mapping(mapping, raw)
  selected <- goecode_tool_apply_mapping(raw, mapping)
  selected <- data.frame(
    Entry_Index = seq_len(nrow(raw)),
    selected,
    stringsAsFactors = FALSE,
    check.names = FALSE
  )

  inputnorm_dir <- file.path(ctx$output_dir, "inputnorm")
  dir.create(inputnorm_dir, recursive = TRUE, showWarnings = FALSE)

  data_file <- file.path(inputnorm_dir, "inputnorm_data.csv")
  preview_file <- file.path(inputnorm_dir, "inputnorm_data_preview.csv")
  summary_file <- file.path(inputnorm_dir, "inputnorm_update_summary.csv")

  goecode_tool_write_csv(selected, data_file)

  max_excel_rows <- 1048576L - 13L
  max_excel_cols <- 27L
  fits_excel <- nrow(selected) <= max_excel_rows && ncol(selected) <= max_excel_cols
  rows_written <- if (fits_excel) nrow(selected) else 0L
  cols_written <- if (fits_excel) ncol(selected) else 0L
  if (fits_excel) {
    goecode_tool_write_csv(selected, preview_file)
  } else if (file.exists(preview_file)) {
    unlink(preview_file)
  }

  summary <- data.frame(
    Metric = c(
      "Source path",
      "Raw rows",
      "Raw columns",
      "Mapped columns",
      "Workbook refresh",
      "Rows written to workbook",
      "Columns written to workbook",
      "Full output",
      "Preview output"
    ),
    Value = c(
      source_path,
      nrow(raw),
      ncol(raw),
      ncol(selected) - 1L,
      if (fits_excel) "ready" else sprintf("refused: output is %s rows x %s columns; N14:AN capacity is %s rows x %s columns", nrow(selected), ncol(selected), max_excel_rows, max_excel_cols),
      rows_written,
      cols_written,
      data_file,
      if (fits_excel) preview_file else ""
    ),
    stringsAsFactors = FALSE
  )
  goecode_tool_write_csv(summary, summary_file)

  if (!fits_excel) {
    return(paste(
      "InputNorm update wrote the full normalized data, but workbook refresh was refused.",
      sprintf("Output folder: %s", ctx$output_dir),
      sprintf("Normalized data: %s", data_file),
      sprintf("Rows: %s", nrow(raw)),
      sprintf("Mapped columns: %s", ncol(selected) - 1L),
      sprintf("Workbook range N14:AN capacity: %s rows x %s columns", max_excel_rows, max_excel_cols),
      sprintf("Requested workbook output: %s rows x %s columns", nrow(selected), ncol(selected)),
      "No partial preview was pasted.",
      sep = vb_newline()
    ))
  }

  paste(
    "InputNorm update completed.",
    sprintf("Output folder: %s", ctx$output_dir),
    sprintf("Normalized data: %s", data_file),
    sprintf("Preview data: %s", preview_file),
    sprintf("Rows: %s", nrow(raw)),
    sprintf("Mapped columns: %s", ncol(selected) - 1L),
    sep = vb_newline()
  )
}

goecode_tool_inputnorm_build <- function(workbook_path, output_dir = NULL, context = list()) {
  ctx <- goecode_tool_context(workbook_path, output_dir, context)
  paste(
    "InputNorm build is reserved for the next geocode workflow step.",
    sprintf("Output folder: %s", ctx$output_dir),
    "Current tab contract: Gather maps headers; Update imports normalized raw data.",
    sep = vb_newline()
  )
}

goecode_tool_run <- goecode_tool_inputnorm_update

goecode_tool_context <- function(workbook_path, output_dir = NULL, context = list()) {
  if (length(context) == 0) {
    context <- btk_build_context(
      tool_id = "goecode_tool",
      workbook_path = workbook_path,
      output_dir = output_dir
    )
  }
  sheets <- goecode_tool_sheet_names(context$workbook_path)
  context$active_sheet <- goecode_tool_active_sheet(context, sheets)
  goecode_tool_validate(context$workbook_path, context)
  context
}

goecode_tool_active_sheet <- function(context, sheets) {
  if (!is.null(context$active_sheet) && nzchar(context$active_sheet)) {
    if (!context$active_sheet %in% sheets) {
      stop(sprintf("Workbook is missing active sheet: %s", context$active_sheet), call. = FALSE)
    }
    return(context$active_sheet)
  }

  inputnorm_sheets <- vapply(sheets, function(sheet) {
    marker <- tryCatch(goecode_tool_cell(context$workbook_path, sheet, row = 1, col = 1), error = function(e) "")
    identical(goecode_tool_marker_key(marker), "inputnorm")
  }, logical(1))

  matches <- sheets[inputnorm_sheets]
  if (length(matches) == 0) {
    stop("Workbook has no <<inputnorm>> sheet.", call. = FALSE)
  }

  matches[1]
}

goecode_tool_cell <- function(workbook_path, sheet, row, col) {
  ref <- paste0(goecode_tool_col_name(col), row)
  value <- goecode_tool_sheet_cell_value(goecode_tool_read_sheet_cached(workbook_path, sheet), ref)
  if (is.na(value)) {
    return("")
  }
  trimws(as.character(value))
}

goecode_tool_sheet_names <- function(workbook_path) {
  temp_dir <- tempfile("goecode_tool_sheets_")
  dir.create(temp_dir, recursive = TRUE, showWarnings = FALSE)
  on.exit(unlink(temp_dir, recursive = TRUE, force = TRUE), add = TRUE)

  workbook_xml <- utils::unzip(workbook_path, files = "xl/workbook.xml", exdir = temp_dir)
  if (length(workbook_xml) == 0 || !file.exists(workbook_xml[1])) {
    stop("Workbook package is missing xl/workbook.xml.", call. = FALSE)
  }

  workbook <- paste(readLines(workbook_xml[1], warn = FALSE), collapse = "")
  sheet_matches <- gregexpr("<sheet\\b[^>]*>", workbook, perl = TRUE)[[1]]
  if (identical(sheet_matches[1], -1L)) {
    stop("Workbook package does not list any worksheets.", call. = FALSE)
  }

  sheet_tags <- regmatches(workbook, list(sheet_matches))[[1]]
  unname(vapply(sheet_tags, function(tag) {
    goecode_tool_xml_unescape(sub('.*\\bname="([^"]+)".*', "\\1", tag))
  }, character(1)))
}

goecode_tool_read_sheet_cached <- function(workbook_path, sheet_name) {
  temp_dir <- tempfile("goecode_tool_sheet_")
  dir.create(temp_dir, recursive = TRUE, showWarnings = FALSE)
  on.exit(unlink(temp_dir, recursive = TRUE, force = TRUE), add = TRUE)

  workbook_xml <- utils::unzip(workbook_path, files = "xl/workbook.xml", exdir = temp_dir)
  workbook_rels <- utils::unzip(workbook_path, files = "xl/_rels/workbook.xml.rels", exdir = temp_dir)
  shared_xml <- suppressWarnings(utils::unzip(workbook_path, files = "xl/sharedStrings.xml", exdir = temp_dir))
  if (length(workbook_xml) == 0 || !file.exists(workbook_xml[1])) {
    stop("Workbook package is missing xl/workbook.xml.", call. = FALSE)
  }
  if (length(workbook_rels) == 0 || !file.exists(workbook_rels[1])) {
    stop("Workbook package is missing xl/_rels/workbook.xml.rels.", call. = FALSE)
  }

  workbook <- paste(readLines(workbook_xml[1], warn = FALSE), collapse = "")
  rels <- paste(readLines(workbook_rels[1], warn = FALSE), collapse = "")

  sheet_matches <- gregexpr("<sheet\\b[^>]*>", workbook, perl = TRUE)[[1]]
  if (identical(sheet_matches[1], -1L)) {
    stop("Workbook package does not list any worksheets.", call. = FALSE)
  }

  sheet_tags <- regmatches(workbook, list(sheet_matches))[[1]]
  sheet_names <- vapply(sheet_tags, function(tag) {
    goecode_tool_xml_unescape(sub('.*\\bname="([^"]+)".*', "\\1", tag))
  }, character(1))

  sheet_tag <- sheet_tags[tolower(sheet_names) == tolower(sheet_name)]
  if (length(sheet_tag) == 0) {
    stop(sprintf("Workbook is missing sheet: %s", sheet_name), call. = FALSE)
  }

  rid <- sub('.*r:id="([^"]+)".*', "\\1", sheet_tag[1])
  rel_tags <- regmatches(rels, gregexpr("<Relationship\\b[^>]*>", rels, perl = TRUE))[[1]]
  rel_tag <- rel_tags[grepl(sprintf('Id="%s"', rid), rel_tags, fixed = TRUE)]
  if (length(rel_tag) == 0) {
    stop(sprintf("Workbook relationship is missing for sheet: %s", sheet_name), call. = FALSE)
  }

  target <- sub('.*Target="([^"]+)".*', "\\1", rel_tag[1])
  sheet_file <- if (startsWith(target, "/")) {
    sub("^/", "", target)
  } else {
    file.path("xl", target)
  }

  sheet_xml <- utils::unzip(workbook_path, files = sheet_file, exdir = temp_dir)
  if (length(sheet_xml) == 0 || !file.exists(sheet_xml[1])) {
    stop(sprintf("Workbook package is missing XML for sheet: %s", sheet_name), call. = FALSE)
  }
  sheet_text <- paste(readLines(sheet_xml[1], warn = FALSE), collapse = "")

  shared <- character()
  if (length(shared_xml) > 0 && file.exists(shared_xml[1])) {
    shared_text <- paste(readLines(shared_xml[1], warn = FALSE), collapse = "")
    shared_items <- regmatches(shared_text, gregexpr("<si>.*?</si>", shared_text, perl = TRUE))[[1]]
    if (!identical(shared_items[1], -1L)) {
      shared <- vapply(shared_items, function(item) {
        item <- gsub("<phoneticPr[^>]*/>", "", item)
        item <- gsub("<rPr>.*?</rPr>", "", item, perl = TRUE)
        text <- gsub("<[^>]+>", "", item)
        goecode_tool_xml_unescape(text)
      }, character(1))
    }
  }

  cell_matches <- gregexpr("<c\\b[^>]*>.*?</c>", sheet_text, perl = TRUE)[[1]]
  cells <- list()
  if (!identical(cell_matches[1], -1L)) {
    cell_tags <- regmatches(sheet_text, list(cell_matches))[[1]]
    for (cell in cell_tags) {
      ref <- sub('.*\\br="([^"]+)".*', "\\1", cell)
      type <- if (grepl('\\bt="s"', cell)) "s" else if (grepl('\\bt="inlineStr"', cell)) "inlineStr" else ""
      value <- ""

      if (grepl("<v>", cell, fixed = TRUE)) {
        raw_value <- sub(".*<v>(.*?)</v>.*", "\\1", cell)
        if (identical(type, "s")) {
          idx <- suppressWarnings(as.integer(raw_value) + 1L)
          value <- if (!is.na(idx) && idx <= length(shared)) shared[idx] else raw_value
        } else {
          value <- goecode_tool_xml_unescape(raw_value)
        }
      } else if (identical(type, "inlineStr") && grepl("<t", cell, fixed = TRUE)) {
        text <- sub(".*<t[^>]*>(.*?)</t>.*", "\\1", cell)
        value <- goecode_tool_xml_unescape(text)
      }

      cells[[ref]] <- value
    }
  }

  list(name = sheet_name, cells = cells)
}

goecode_tool_sheet_cell_value <- function(sheet, ref) {
  value <- sheet$cells[[ref]]
  if (is.null(value)) {
    return("")
  }
  value
}

goecode_tool_col_name <- function(col) {
  name <- ""
  while (col > 0) {
    rem <- (col - 1) %% 26
    name <- paste0(intToUtf8(65 + rem), name)
    col <- (col - rem - 1) %/% 26
  }
  name
}

goecode_tool_xml_unescape <- function(value) {
  value <- gsub("&lt;", "<", value, fixed = TRUE)
  value <- gsub("&gt;", ">", value, fixed = TRUE)
  value <- gsub("&amp;", "&", value, fixed = TRUE)
  value <- gsub("&quot;", '"', value, fixed = TRUE)
  value <- gsub("&apos;", "'", value, fixed = TRUE)
  value
}

goecode_tool_marker_key <- function(value) {
  value <- tolower(trimws(as.character(value)))
  value <- gsub("^<+|>+$", "", value)
  value <- gsub("[-_]+", " ", value)
  value <- gsub("\\s+", " ", value)
  trimws(value)
}

goecode_tool_source_path <- function(workbook_path, sheet) {
  source_path <- goecode_tool_cell(workbook_path, sheet, row = 14, col = 3)
  if (!nzchar(source_path)) {
    stop(sprintf("InputNorm source path is required in %s!C14.", sheet), call. = FALSE)
  }
  btk_require_file(source_path, "source data file")
  source_path
}

goecode_tool_alias_dictionary <- function(workbook_path) {
  control <- goecode_tool_read_sheet_cached(workbook_path, "Control")
  aliases <- list()
  for (col in seq_len(8L)) {
    col_name <- goecode_tool_col_name(col)
    canonical <- trimws(as.character(goecode_tool_sheet_cell_value(control, paste0(col_name, "1"))))
    canonical <- sub(":$", "", canonical)
    if (!nzchar(canonical) || is.na(canonical)) next

    values <- vapply(seq_len(2000L), function(row) {
      goecode_tool_sheet_cell_value(control, paste0(col_name, row))
    }, character(1))
    values <- values[!is.na(values) & nzchar(trimws(values))]
    values <- gsub("^\\s*-\\s*", "", values)
    values <- unique(c(canonical, values))
    aliases[[canonical]] <- values
  }
  aliases
}

goecode_tool_guess_headers <- function(headers, aliases) {
  alias_keys <- list()
  for (canonical in names(aliases)) {
    for (alias in aliases[[canonical]]) {
      alias_keys[[goecode_tool_normalize_header(alias)]] <- canonical
    }
  }

  mapped <- character(length(headers))
  credibility <- character(length(headers))
  for (i in seq_along(headers)) {
    key <- goecode_tool_normalize_header(headers[i])
    if (nzchar(key) && !is.null(alias_keys[[key]])) {
      mapped[i] <- alias_keys[[key]]
      credibility[i] <- ""
    } else {
      mapped[i] <- ""
      credibility[i] <- ""
    }
  }

  data.frame(header_mapped = mapped, Credibility = credibility, stringsAsFactors = FALSE)
}

goecode_tool_normalize_header <- function(value) {
  value <- tolower(trimws(as.character(value)))
  value <- gsub("^\\s*-\\s*", "", value)
  value <- gsub("[[:punct:]_]+", " ", value)
  value <- gsub("\\s+", " ", value)
  trimws(value)
}

goecode_tool_source_headers <- function(source_path) {
  ext <- tolower(tools::file_ext(source_path))
  if (ext %in% c("csv", "txt")) {
    first_line <- goecode_tool_first_line(source_path)
    if (length(first_line) == 0) {
      stop("Source data file is empty.", call. = FALSE)
    }
    headers <- goecode_tool_parse_csv_line(first_line)
    return(as.character(headers))
  }

  if (ext %in% c("xlsx", "xlsm")) {
    btk_require_packages(c("openxlsx"))
    headers <- openxlsx::readWorkbook(
      source_path,
      sheet = 1,
      rows = 1,
      colNames = FALSE,
      skipEmptyRows = FALSE,
      skipEmptyCols = FALSE
    )
    return(as.character(headers[1, ]))
  }

  stop(sprintf("Unsupported source data file type: .%s", ext), call. = FALSE)
}

goecode_tool_write_gather_diagnostic <- function(source_path, headers, diagnostic_file) {
  first_line <- goecode_tool_first_line(source_path)
  raw_bytes <- goecode_tool_first_bytes(source_path, n = 32L)
  info <- file.info(source_path)

  rows <- data.frame(
    Metric = c(
      "source_path",
      "normalized_source_path",
      "file_exists",
      "file_size",
      "working_directory",
      "r_version",
      "locale",
      "encoding_option",
      "first_32_bytes_hex",
      "first_line",
      "headers_detected_count",
      paste0("header_", seq_along(headers))
    ),
    Value = c(
      source_path,
      normalizePath(source_path, winslash = "/", mustWork = FALSE),
      file.exists(source_path),
      if (nrow(info) > 0 && !is.na(info$size[1])) info$size[1] else NA,
      getwd(),
      paste(R.version$major, R.version$minor, sep = "."),
      paste(Sys.getlocale(), collapse = " | "),
      getOption("encoding"),
      raw_bytes,
      first_line,
      length(headers),
      headers
    ),
    stringsAsFactors = FALSE,
    check.names = FALSE
  )

  goecode_tool_write_csv(rows, diagnostic_file)
  invisible(diagnostic_file)
}

goecode_tool_first_line <- function(source_path) {
  bytes <- goecode_tool_read_file_bytes(source_path)
  if (length(bytes) == 0) {
    return("")
  }

  line_end <- which(bytes == as.raw(0x0A))[1]
  if (is.na(line_end)) {
    line_bytes <- bytes
  } else {
    line_bytes <- bytes[seq_len(line_end - 1L)]
  }

  if (length(line_bytes) > 0 && tail(line_bytes, 1L) == as.raw(0x0D)) {
    line_bytes <- line_bytes[-length(line_bytes)]
  }

  goecode_tool_decode_utf8_bytes(line_bytes)
}

goecode_tool_first_bytes <- function(source_path, n = 32L) {
  con <- file(source_path, open = "rb")
  on.exit(close(con), add = TRUE)
  bytes <- readBin(con, what = "raw", n = n)
  paste(sprintf("%02X", as.integer(bytes)), collapse = " ")
}

goecode_tool_read_file_text <- function(source_path) {
  bytes <- goecode_tool_read_file_bytes(source_path)
  goecode_tool_decode_utf8_bytes(bytes)
}

goecode_tool_read_file_bytes <- function(source_path) {
  con <- file(source_path, open = "rb")
  on.exit(close(con), add = TRUE)
  readBin(con, what = "raw", n = file.info(source_path)$size)
}

goecode_tool_decode_utf8_bytes <- function(bytes) {
  if (length(bytes) >= 3L && identical(bytes[1:3], as.raw(c(0xEF, 0xBB, 0xBF)))) {
    bytes <- bytes[-(1:3)]
  }
  text <- rawToChar(bytes)
  text <- iconv(text, from = "UTF-8", to = "UTF-8", sub = "")
  if (is.na(text)) {
    text <- rawToChar(bytes)
  }
  Encoding(text) <- "UTF-8"
  text
}

goecode_tool_write_csv <- function(data, path) {
  data <- as.data.frame(data, stringsAsFactors = FALSE, check.names = FALSE)
  data[] <- lapply(data, function(col) {
    col <- as.character(col)
    col[is.na(col)] <- ""
    Encoding(col) <- "UTF-8"
    col
  })

  con <- file(path, open = "wb")
  on.exit(close(con), add = TRUE)
  writeBin(as.raw(c(0xEF, 0xBB, 0xBF)), con)

  goecode_tool_write_csv_line(names(data), con)
  if (nrow(data) == 0) {
    return(invisible(path))
  }

  for (row in seq_len(nrow(data))) {
    values <- vapply(data[row, , drop = FALSE], function(value) value[[1]], character(1))
    goecode_tool_write_csv_line(values, con)
  }

  invisible(path)
}

goecode_tool_write_csv_line <- function(values, con) {
  line <- paste(vapply(values, goecode_tool_csv_quote, character(1)), collapse = ",")
  writeBin(charToRaw(enc2utf8(paste0(line, "\r\n"))), con)
}

goecode_tool_csv_quote <- function(value) {
  if (is.na(value)) {
    value <- ""
  }
  value <- as.character(value)
  Encoding(value) <- "UTF-8"
  value <- gsub('"', '""', value, fixed = TRUE)
  paste0('"', value, '"')
}

goecode_tool_parse_csv_text <- function(text, header = TRUE) {
  text <- sub("(\\r\\n|\\n|\\r)$", "", text)
  if (!nzchar(text)) {
    return(data.frame(stringsAsFactors = FALSE, check.names = FALSE))
  }

  lines <- strsplit(text, "\n", fixed = TRUE, useBytes = FALSE)[[1]]
  lines <- sub("\\r$", "", lines)
  rows <- lapply(lines, goecode_tool_parse_csv_line)

  if (length(rows) == 0) {
    return(data.frame(stringsAsFactors = FALSE, check.names = FALSE))
  }

  if (header) {
    headers <- rows[[1]]
    data_rows <- rows[-1]
  } else {
    data_rows <- rows
    width <- max(vapply(data_rows, length, integer(1)))
    headers <- paste0("V", seq_len(width))
  }

  if (length(headers) == 0) {
    return(data.frame(stringsAsFactors = FALSE, check.names = FALSE))
  }

  headers <- as.character(headers)
  headers[!nzchar(headers)] <- paste0("X", which(!nzchar(headers)))
  headers <- make.unique(headers)

  if (length(data_rows) == 0) {
    out <- as.data.frame(setNames(replicate(length(headers), character(), simplify = FALSE), headers), stringsAsFactors = FALSE)
    return(out)
  }

  normalized <- lapply(data_rows, function(row) {
    row <- as.character(row)
    if (length(row) < length(headers)) {
      row <- c(row, rep("", length(headers) - length(row)))
    }
    if (length(row) > length(headers)) {
      row <- row[seq_along(headers)]
    }
    row
  })

  matrix_data <- do.call(rbind, normalized)
  out <- as.data.frame(matrix_data, stringsAsFactors = FALSE, check.names = FALSE)
  names(out) <- headers
  out[] <- lapply(out, function(col) {
    Encoding(col) <- "UTF-8"
    col
  })
  out
}

goecode_tool_parse_csv_line <- function(line) {
  if (is.na(line)) {
    return("")
  }
  line <- as.character(line)
  Encoding(line) <- "UTF-8"

  chars <- strsplit(line, "", fixed = TRUE, useBytes = FALSE)[[1]]
  fields <- character()
  field <- character()
  in_quotes <- FALSE
  i <- 1L

  while (i <= length(chars)) {
    ch <- chars[[i]]
    if (identical(ch, '"')) {
      if (in_quotes && i < length(chars) && identical(chars[[i + 1L]], '"')) {
        field <- c(field, '"')
        i <- i + 1L
      } else {
        in_quotes <- !in_quotes
      }
    } else if (identical(ch, ",") && !in_quotes) {
      value <- paste(field, collapse = "")
      Encoding(value) <- "UTF-8"
      fields <- c(fields, value)
      field <- character()
    } else {
      field <- c(field, ch)
    }
    i <- i + 1L
  }

  value <- paste(field, collapse = "")
  Encoding(value) <- "UTF-8"
  fields <- c(fields, value)
  fields
}

goecode_tool_read_header_mapping <- function(workbook_path, sheet) {
  input_sheet <- goecode_tool_read_sheet_cached(workbook_path, sheet)
  rows <- lapply(26:2000, function(row) {
    data.frame(
      header_raw = trimws(as.character(goecode_tool_sheet_cell_value(input_sheet, paste0("B", row)))),
      header_mapped = trimws(as.character(goecode_tool_sheet_cell_value(input_sheet, paste0("C", row)))),
      stringsAsFactors = FALSE,
      check.names = FALSE
    )
  })
  values <- do.call(rbind, rows)
  values <- values[!(tolower(values$header_raw) == "header_raw" & tolower(values$header_mapped) == "header_mapped"), , drop = FALSE]
  values <- values[nzchar(values$header_raw), , drop = FALSE]
  values <- values[nzchar(values$header_mapped), , drop = FALSE]
  rownames(values) <- NULL
  values
}

goecode_tool_validate_mapping <- function(mapping, raw) {
  duplicate_mapped <- unique(mapping$header_mapped[duplicated(mapping$header_mapped)])
  duplicate_mapped <- duplicate_mapped[nzchar(duplicate_mapped)]
  if (length(duplicate_mapped) > 0) {
    stop(
      sprintf("Duplicate normalized header(s) in C26:C: %s", paste(duplicate_mapped, collapse = ", ")),
      call. = FALSE
    )
  }

  missing_raw <- setdiff(mapping$header_raw, names(raw))
  if (length(missing_raw) > 0) {
    stop(
      sprintf("Mapped raw header(s) are missing from source file: %s", paste(missing_raw, collapse = ", ")),
      call. = FALSE
    )
  }

  invisible(TRUE)
}

goecode_tool_read_source_table <- function(source_path) {
  ext <- tolower(tools::file_ext(source_path))
  if (ext %in% c("csv", "txt")) {
    text <- goecode_tool_read_file_text(source_path)
    data <- goecode_tool_parse_csv_text(text, header = TRUE)
    data[is.na(data)] <- ""
    return(data)
  }

  if (ext %in% c("xlsx", "xlsm")) {
    btk_require_packages(c("openxlsx"))
    data <- openxlsx::readWorkbook(
      source_path,
      sheet = 1,
      colNames = TRUE,
      skipEmptyRows = FALSE,
      skipEmptyCols = FALSE,
      detectDates = FALSE
    )
    data[] <- lapply(data, as.character)
    data[is.na(data)] <- ""
    return(data)
  }

  stop(sprintf("Unsupported source data file type: .%s", ext), call. = FALSE)
}

goecode_tool_apply_mapping <- function(raw, mapping) {
  out <- data.frame(.row_id = seq_len(nrow(raw)), stringsAsFactors = FALSE)
  out$.row_id <- NULL
  for (source in names(raw)) {
    idx <- match(source, mapping$header_raw)
    if (is.na(idx)) next
    target <- mapping$header_mapped[idx]
    if (!nzchar(target)) next
    out[[target]] <- raw[[source]]
  }
  out
}
