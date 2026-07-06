goecode_tool_validate <- function(workbook_path, context = list()) {
  btk_require_file(workbook_path, "workbook")

  sheets <- goecode_tool_sheet_names(workbook_path)
  if (!"Control" %in% sheets) {
    stop("Geocode workbook is missing required sheet: Control", call. = FALSE)
  }

  active_sheet <- goecode_tool_active_sheet(context, sheets)
  marker <- goecode_tool_cell(workbook_path, active_sheet, row = 1, col = 1)
  if (!goecode_tool_marker_key(marker) %in% c("inputnorm", "histcross", "generalcrosswalk")) {
    stop(sprintf("Active sheet '%s' is not a supported Geocode tab.", active_sheet), call. = FALSE)
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

  preview_limit <- goecode_tool_workbook_preview_rows()
  max_excel_cols <- 27L
  fits_excel <- ncol(selected) <= max_excel_cols
  preview <- selected[seq_len(min(nrow(selected), preview_limit)), , drop = FALSE]
  rows_written <- if (fits_excel) nrow(preview) else 0L
  cols_written <- if (fits_excel) ncol(preview) else 0L
  if (fits_excel) {
    goecode_tool_write_csv(preview, preview_file)
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
      "Workbook preview row cap",
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
      if (fits_excel) if (nrow(selected) > preview_limit) "ready: capped preview pasted; full output written to CSV" else "ready" else sprintf("refused: output has %s columns; N14:AN capacity is %s columns", ncol(selected), max_excel_cols),
      preview_limit,
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
      sprintf("Workbook range N14:AN capacity: %s columns", max_excel_cols),
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
    sprintf("Workbook preview rows: %s%s", rows_written, if (nrow(selected) > preview_limit) " (capped)" else ""),
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

goecode_tool_workbook_preview_rows <- function() {
  10000L
}

goecode_tool_histcross_gather <- function(workbook_path, output_dir = NULL, context = list()) {
  ctx <- goecode_tool_context(workbook_path, output_dir, context)
  source_path <- goecode_tool_source_path(ctx$workbook_path, ctx$active_sheet)

  headers <- goecode_tool_source_headers(source_path)
  aliases <- goecode_tool_histcross_alias_dictionary(ctx$workbook_path)
  guesses <- goecode_tool_guess_headers(headers, aliases)

  out <- data.frame(
    header_raw = headers,
    header_mapped = guesses$header_mapped,
    Credibility = guesses$Credibility,
    stringsAsFactors = FALSE,
    check.names = FALSE
  )

  histcross_dir <- file.path(ctx$output_dir, "histcross")
  dir.create(histcross_dir, recursive = TRUE, showWarnings = FALSE)
  mapping_file <- file.path(histcross_dir, "histcross_header_mapping.csv")
  diagnostic_file <- file.path(histcross_dir, "histcross_gather_diagnostic.csv")
  goecode_tool_write_csv(out, mapping_file)
  goecode_tool_write_gather_diagnostic(
    source_path = source_path,
    headers = headers,
    diagnostic_file = diagnostic_file
  )

  paste(
    "HistCross gather completed.",
    sprintf("Output folder: %s", ctx$output_dir),
    sprintf("Header mapping: %s", mapping_file),
    sprintf("Diagnostic: %s", diagnostic_file),
    sprintf("Headers read: %s", length(headers)),
    sep = vb_newline()
  )
}

goecode_tool_histcross_update <- function(workbook_path, output_dir = NULL, context = list()) {
  ctx <- goecode_tool_context(workbook_path, output_dir, context)
  source_path <- goecode_tool_source_path(ctx$workbook_path, ctx$active_sheet)
  input_sheet_name <- goecode_tool_cell(ctx$workbook_path, ctx$active_sheet, row = 5, col = 3)
  if (!nzchar(input_sheet_name)) {
    stop("HistCross requires the related InputNorm tab name in C5.", call. = FALSE)
  }
  goecode_tool_validate_inputnorm_dependency(ctx$workbook_path, input_sheet_name)

  mapping <- goecode_tool_read_header_mapping(ctx$workbook_path, ctx$active_sheet)
  if (nrow(mapping) == 0) {
    stop("No header mapping found in B26:C. Run Gather first, then configure C26:C.", call. = FALSE)
  }

  prior_raw <- goecode_tool_read_source_table(source_path)
  goecode_tool_validate_mapping(mapping, prior_raw)
  prior <- goecode_tool_apply_mapping(prior_raw, mapping)
  inputnorm_data_file <- goecode_tool_inputnorm_data_file(ctx$output_dir, input_sheet_name)
  current <- goecode_tool_read_source_table(inputnorm_data_file)
  if (!"Entry_Index" %in% names(current)) {
    stop(sprintf("InputNorm artifact for tab '%s' does not contain Entry_Index: %s", input_sheet_name, inputnorm_data_file), call. = FALSE)
  }

  key_cols <- mapping$header_mapped[grepl("_raw$", mapping$header_mapped)]
  key_cols <- unique(key_cols[key_cols %in% names(current) & key_cols %in% names(prior)])
  mapped_cols <- unique(mapping$header_mapped[grepl("^mapped_", mapping$header_mapped)])
  mapped_cols <- mapped_cols[mapped_cols %in% names(prior)]
  if (length(key_cols) == 0) {
    stop("HistCross requires at least one mapped *_raw join key that also exists in the InputNorm output.", call. = FALSE)
  }
  if (length(mapped_cols) == 0) {
    stop("HistCross requires at least one mapped_* result column in the prior-year file mapping.", call. = FALSE)
  }

  joined <- goecode_tool_histcross_join(current, prior, key_cols, mapped_cols)

  histcross_dir <- file.path(ctx$output_dir, "histcross")
  dir.create(histcross_dir, recursive = TRUE, showWarnings = FALSE)
  matched_file <- file.path(histcross_dir, "histcross_matched.csv")
  unmatched_file <- file.path(histcross_dir, "histcross_unmatched.csv")
  matched_preview_file <- file.path(histcross_dir, "histcross_matched_preview.csv")
  unmatched_preview_file <- file.path(histcross_dir, "histcross_unmatched_preview.csv")
  warnings_file <- file.path(histcross_dir, "histcross_warnings.csv")
  summary_file <- file.path(histcross_dir, "histcross_update_summary.csv")

  goecode_tool_write_csv(joined$matched, matched_file)
  goecode_tool_write_csv(joined$unmatched, unmatched_file)
  goecode_tool_write_csv(joined$warnings, warnings_file)

  preview_limit <- goecode_tool_workbook_preview_rows()
  max_excel_rows <- 1048576L - 13L
  matched_fits <- ncol(joined$matched) <= 27L
  unmatched_fits <- nrow(joined$unmatched) <= max_excel_rows && ncol(joined$unmatched) <= 27L
  matched_preview <- joined$matched[seq_len(min(nrow(joined$matched), preview_limit)), , drop = FALSE]
  unmatched_preview <- joined$unmatched
  if (matched_fits && unmatched_fits) {
    goecode_tool_write_csv(matched_preview, matched_preview_file)
    goecode_tool_write_csv(unmatched_preview, unmatched_preview_file)
  } else {
    if (file.exists(matched_preview_file)) unlink(matched_preview_file)
    if (file.exists(unmatched_preview_file)) unlink(unmatched_preview_file)
  }

  summary <- data.frame(
    Metric = c(
      "Prior-year source path",
      "InputNorm tab",
      "InputNorm data artifact",
      "Current input rows",
      "Prior-year rows",
      "Join keys",
      "Mapped result columns",
      "Matched rows",
      "Unmatched rows",
      "Warnings",
      "Workbook refresh",
      "Matched workbook preview row cap",
      "Matched rows written to workbook",
      "Unmatched rows written to workbook",
      "Matched output",
      "Unmatched output",
      "Warning output"
    ),
    Value = c(
      source_path,
      input_sheet_name,
      inputnorm_data_file,
      nrow(current),
      nrow(prior),
      paste(key_cols, collapse = ", "),
      paste(mapped_cols, collapse = ", "),
      nrow(joined$matched),
      nrow(joined$unmatched),
      nrow(joined$warnings),
      if (matched_fits && unmatched_fits) if (nrow(joined$matched) > preview_limit) "ready: matched preview capped; unmatched pasted in full" else "ready: unmatched pasted in full" else sprintf("refused: matched output must fit 27 columns; unmatched output must fit %s rows x 27 columns", max_excel_rows),
      preview_limit,
      if (matched_fits && unmatched_fits) nrow(matched_preview) else 0L,
      if (matched_fits && unmatched_fits) nrow(unmatched_preview) else 0L,
      matched_file,
      unmatched_file,
      warnings_file
    ),
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
  goecode_tool_write_csv(summary, summary_file)

  if (!matched_fits || !unmatched_fits) {
    return(paste(
      "HistCross update wrote outputs, but workbook refresh was refused.",
      sprintf("Output folder: %s", ctx$output_dir),
      sprintf("Matched rows: %s", nrow(joined$matched)),
      sprintf("Unmatched rows: %s", nrow(joined$unmatched)),
      sprintf("Warnings: %s", nrow(joined$warnings)),
      sprintf("Workbook unmatched capacity: %s rows x 27 columns", max_excel_rows),
      "No workbook output was pasted.",
      sep = vb_newline()
    ))
  }

  paste(
    "HistCross update completed.",
    sprintf("Output folder: %s", ctx$output_dir),
    sprintf("Matched rows: %s", nrow(joined$matched)),
    sprintf("Unmatched rows: %s", nrow(joined$unmatched)),
    sprintf("Workbook matched preview rows: %s%s", nrow(matched_preview), if (nrow(joined$matched) > preview_limit) " (capped)" else ""),
    sprintf("Workbook unmatched rows: %s (full)", nrow(unmatched_preview)),
    sprintf("Warnings: %s", nrow(joined$warnings)),
    sprintf("Join keys: %s", paste(key_cols, collapse = ", ")),
    sep = vb_newline()
  )
}

goecode_tool_histcross_build <- function(workbook_path, output_dir = NULL, context = list()) {
  ctx <- goecode_tool_context(workbook_path, output_dir, context)
  paste(
    "HistCross build is reserved for the next geocode workflow step.",
    sprintf("Output folder: %s", ctx$output_dir),
    "Current tab contract: Gather maps prior-year headers; Update splits matched and unmatched rows.",
    sep = vb_newline()
  )
}

goecode_tool_generalcrosswalk_gather <- function(workbook_path, output_dir = NULL, context = list()) {
  ctx <- goecode_tool_context(workbook_path, output_dir, context)
  input <- goecode_tool_generalcrosswalk_input(ctx)
  input <- goecode_tool_ensure_entry_index(input)

  config <- goecode_tool_generalcrosswalk_existing_config(ctx$workbook_path, ctx$active_sheet)
  config <- goecode_tool_generalcrosswalk_default_config(names(input), config)

  general_dir <- file.path(ctx$output_dir, "generalcrosswalk")
  dir.create(general_dir, recursive = TRUE, showWarnings = FALSE)
  input_file <- file.path(general_dir, "generalcrosswalk_input.csv")
  input_preview_file <- file.path(general_dir, "generalcrosswalk_input_preview.csv")
  config_file <- file.path(general_dir, "generalcrosswalk_config.csv")
  summary_file <- file.path(general_dir, "generalcrosswalk_gather_summary.csv")

  goecode_tool_write_csv(input, input_file)
  goecode_tool_write_csv(input, input_preview_file)
  goecode_tool_write_csv(config, config_file)

  summary <- data.frame(
    Metric = c("Input source", "Rows", "Columns", "Input artifact", "Config artifact"),
    Value = c(attr(input, "source_label"), nrow(input), ncol(input), input_file, config_file),
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
  goecode_tool_write_csv(summary, summary_file)

  paste(
    "GeneralCrossWalk gather completed.",
    sprintf("Output folder: %s", ctx$output_dir),
    sprintf("Input source: %s", attr(input, "source_label")),
    sprintf("Rows staged: %s", nrow(input)),
    sprintf("Columns staged: %s", ncol(input)),
    sprintf("Header config: %s", config_file),
    sep = vb_newline()
  )
}

goecode_tool_generalcrosswalk_update <- function(workbook_path, output_dir = NULL, context = list()) {
  ctx <- goecode_tool_context(workbook_path, output_dir, context)
  general_dir <- file.path(ctx$output_dir, "generalcrosswalk")
  dir.create(general_dir, recursive = TRUE, showWarnings = FALSE)

  input_file <- file.path(general_dir, "generalcrosswalk_input.csv")
  input <- if (file.exists(input_file)) {
    goecode_tool_read_source_table(input_file)
  } else {
    goecode_tool_generalcrosswalk_input(ctx)
  }
  input <- goecode_tool_ensure_entry_index(input)

  config <- goecode_tool_generalcrosswalk_read_config(ctx$workbook_path, ctx$active_sheet)
  selected_config <- goecode_tool_generalcrosswalk_selected_config(config, input)
  target_resolution <- goecode_tool_generalcrosswalk_target_resolution(ctx$workbook_path, ctx$active_sheet)
  threshold <- goecode_tool_generalcrosswalk_review_threshold(ctx$workbook_path, ctx$active_sheet)
  geo <- goecode_tool_generalcrosswalk_geo_index(ctx$workbook_path)
  zones <- goecode_tool_generalcrosswalk_zone_index(ctx$workbook_path, geo)

  ppc <- goecode_tool_generalcrosswalk_ppc_candidates(input, selected_config, geo, zones)
  address <- goecode_tool_generalcrosswalk_text_candidates(input, selected_config, geo, zones, "Address")
  insured <- goecode_tool_generalcrosswalk_text_candidates(input, selected_config, geo, zones, "Insured")
  promoted <- goecode_tool_generalcrosswalk_promote(input, selected_config, ppc, address, insured, target_resolution, threshold)

  config_file <- file.path(general_dir, "generalcrosswalk_config.csv")
  ppc_file <- file.path(general_dir, "generalcrosswalk_ppc_candidates.csv")
  address_file <- file.path(general_dir, "generalcrosswalk_address_candidates.csv")
  insured_file <- file.path(general_dir, "generalcrosswalk_insured_candidates.csv")
  final_file <- file.path(general_dir, "generalcrosswalk_final_mapped.csv")
  review_file <- file.path(general_dir, "generalcrosswalk_manual_review.csv")
  warnings_file <- file.path(general_dir, "generalcrosswalk_warnings.csv")
  summary_file <- file.path(general_dir, "generalcrosswalk_update_summary.csv")
  mapped_output <- goecode_tool_generalcrosswalk_mapped_output(input, promoted$mapped)

  goecode_tool_write_csv(input, input_file)
  goecode_tool_write_csv(config, config_file)
  goecode_tool_write_csv(ppc, ppc_file)
  goecode_tool_write_csv(address, address_file)
  goecode_tool_write_csv(insured, insured_file)
  goecode_tool_write_csv(mapped_output, final_file)
  goecode_tool_write_csv(promoted$review, review_file)
  goecode_tool_write_csv(promoted$warnings, warnings_file)

  summary <- data.frame(
    Metric = c(
      "Input rows",
      "Selected config rows",
      "Target resolution",
      "Review threshold",
      "PPC candidates",
      "Address candidates",
      "Insured candidates",
      "Mapped rows",
      "Manual review rows",
      "Warnings"
    ),
    Value = c(
      nrow(input),
      nrow(selected_config),
      target_resolution,
      threshold,
      sum(nzchar(ppc$candidate_status)),
      sum(nzchar(address$candidate_status)),
      sum(nzchar(insured$candidate_status)),
      nrow(promoted$mapped),
      nrow(promoted$review),
      nrow(promoted$warnings)
    ),
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
  goecode_tool_write_csv(summary, summary_file)

  paste(
    "GeneralCrossWalk update completed.",
    sprintf("Output folder: %s", ctx$output_dir),
    sprintf("Rows processed: %s", nrow(input)),
    sprintf("Mapped rows: %s", nrow(promoted$mapped)),
    sprintf("Manual review rows: %s", nrow(promoted$review)),
    sprintf("Review threshold: %s", threshold),
    sep = vb_newline()
  )
}

goecode_tool_generalcrosswalk_build <- function(workbook_path, output_dir = NULL, context = list()) {
  ctx <- goecode_tool_context(workbook_path, output_dir, context)
  paste(
    "GeneralCrossWalk build is reserved for later export workflow steps.",
    sprintf("Output folder: %s", ctx$output_dir),
    "Current contract: Gather stages input; Update creates mapped output and manual review queue.",
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

  action_sheets <- vapply(sheets, function(sheet) {
    marker <- tryCatch(goecode_tool_cell(context$workbook_path, sheet, row = 1, col = 1), error = function(e) "")
    goecode_tool_marker_key(marker) %in% c("inputnorm", "histcross", "generalcrosswalk")
  }, logical(1))

  matches <- sheets[action_sheets]
  if (length(matches) == 0) {
    stop("Workbook has no supported Geocode action sheet.", call. = FALSE)
  }

  matches[1]
}

goecode_tool_cell <- function(workbook_path, sheet, row, col) {
  ref <- paste0(goecode_tool_col_name(col), row)
  value <- goecode_tool_read_cell_cached(workbook_path, sheet, ref)
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

  cells <- list()
  for (cell in goecode_tool_extract_cell_tags(sheet_text)) {
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

  list(name = sheet_name, cells = cells)
}

goecode_tool_read_cell_cached <- function(workbook_path, sheet_name, ref) {
  temp_dir <- tempfile("goecode_tool_cell_")
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
  sheet_file <- goecode_tool_sheet_file(workbook, rels, sheet_name)
  sheet_xml <- utils::unzip(workbook_path, files = sheet_file, exdir = temp_dir)
  if (length(sheet_xml) == 0 || !file.exists(sheet_xml[1])) {
    stop(sprintf("Workbook package is missing XML for sheet: %s", sheet_name), call. = FALSE)
  }

  shared <- goecode_tool_shared_strings(shared_xml)
  cell <- goecode_tool_find_cell_tag_in_file(sheet_xml[1], ref)
  if (!nzchar(cell)) {
    return("")
  }
  goecode_tool_cell_tag_value(cell, shared)
}

goecode_tool_sheet_file <- function(workbook, rels, sheet_name) {
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
  if (startsWith(target, "/")) {
    sub("^/", "", target)
  } else {
    file.path("xl", target)
  }
}

goecode_tool_shared_strings <- function(shared_xml) {
  shared <- character()
  if (length(shared_xml) > 0 && file.exists(shared_xml[1])) {
    shared_text <- paste(readLines(shared_xml[1], warn = FALSE), collapse = "")
    parts <- strsplit(shared_text, "<si>", fixed = TRUE, useBytes = TRUE)[[1]]
    if (length(parts) > 1L) {
      parts <- parts[-1L]
      shared <- vapply(parts, function(item) {
      close <- regexpr("</si>", item, fixed = TRUE)[[1]]
        if (close > 0L) item <- substr(item, 1L, close - 1L)
        item <- gsub("<phoneticPr[^>]*/>", "", item)
        item <- gsub("<rPr>.*?</rPr>", "", item, perl = TRUE)
        text <- gsub("<[^>]+>", "", item)
        goecode_tool_xml_unescape(text)
      }, character(1), USE.NAMES = FALSE)
    }
  }
  shared
}

goecode_tool_find_cell_tag_in_file <- function(sheet_xml, ref) {
  con <- file(sheet_xml, open = "rb")
  on.exit(close(con), add = TRUE)

  pattern <- charToRaw(paste0('r="', ref, '"'))
  carry <- raw()
  repeat {
    chunk <- readBin(con, what = "raw", n = 1048576L)
    if (length(chunk) == 0) break
    buffer <- c(carry, chunk)
    loc <- goecode_tool_raw_match(pattern, buffer)
    if (!is.na(loc)) {
      text <- rawToChar(buffer)
      ref_pos <- loc
      before <- substr(text, 1L, ref_pos)
      start_matches <- gregexpr("<c", before, fixed = TRUE)[[1]]
      start <- tail(start_matches[start_matches > 0], 1L)
      after <- substr(text, ref_pos, nchar(text))
      close <- regexpr("</c>", after, fixed = TRUE)[[1]]
      if (length(start) == 1L && !is.na(start) && close > 0L) {
        return(substr(text, start, ref_pos + close + 3L))
      }
    }
    carry <- tail(buffer, min(length(buffer), 4096L))
  }
  ""
}

goecode_tool_raw_match <- function(pattern, buffer) {
  if (length(pattern) == 0 || length(buffer) < length(pattern)) {
    return(NA_integer_)
  }
  first <- pattern[[1]]
  starts <- which(buffer == first)
  starts <- starts[starts + length(pattern) - 1L <= length(buffer)]
  for (start in starts) {
    if (identical(buffer[start:(start + length(pattern) - 1L)], pattern)) {
      return(start)
    }
  }
  NA_integer_
}

goecode_tool_cell_tag_value <- function(cell, shared) {
  type <- if (grepl('\\bt="s"', cell)) "s" else if (grepl('\\bt="inlineStr"', cell)) "inlineStr" else ""

  if (grepl("<v>", cell, fixed = TRUE)) {
    raw_value <- sub(".*<v>(.*?)</v>.*", "\\1", cell)
    if (identical(type, "s")) {
      idx <- suppressWarnings(as.integer(raw_value) + 1L)
      return(if (!is.na(idx) && idx <= length(shared)) shared[idx] else raw_value)
    }
    return(goecode_tool_xml_unescape(raw_value))
  }

  if (identical(type, "inlineStr") && grepl("<t", cell, fixed = TRUE)) {
    text <- sub(".*<t[^>]*>(.*?)</t>.*", "\\1", cell)
    return(goecode_tool_xml_unescape(text))
  }

  ""
}

goecode_tool_extract_cell_tags <- function(sheet_text) {
  parts <- strsplit(sheet_text, "<c", fixed = TRUE, useBytes = TRUE)[[1]]
  if (length(parts) <= 1L) {
    return(character())
  }

  parts <- parts[-1L]
  cells <- vector("list", length(parts))
  used <- 0L
  for (idx in seq_along(parts)) {
    part <- parts[[idx]]
    close_cell <- regexpr("</c>", part, fixed = TRUE, useBytes = TRUE)[[1]]
    if (close_cell < 0L) next
    used <- used + 1L
    cells[[used]] <- paste0("<c", substr(part, 1L, close_cell + 3L))
  }
  if (used == 0L) {
    return(character())
  }
  unlist(cells[seq_len(used)], use.names = FALSE)
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
  value <- goecode_tool_xml_unescape_numeric(value, "&#([0-9]+);", 10L)
  value <- goecode_tool_xml_unescape_numeric(value, "&#x([0-9A-Fa-f]+);", 16L)
  value
}

goecode_tool_xml_unescape_numeric <- function(value, pattern, base) {
  matches <- gregexpr(pattern, value, perl = TRUE)
  regmatches(value, matches) <- lapply(regmatches(value, matches), function(items) {
    vapply(items, function(item) {
      code <- sub(pattern, "\\1", item, perl = TRUE)
      number <- suppressWarnings(strtoi(code, base = base))
      if (is.na(number)) item else intToUtf8(number)
    }, character(1))
  })
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

goecode_tool_validate_inputnorm_dependency <- function(workbook_path, input_sheet_name) {
  sheets <- goecode_tool_sheet_names(workbook_path)
  if (!input_sheet_name %in% sheets) {
    stop(sprintf("HistCross C5 points to a missing InputNorm tab: %s", input_sheet_name), call. = FALSE)
  }

  marker <- goecode_tool_cell(workbook_path, input_sheet_name, row = 1, col = 1)
  if (!identical(goecode_tool_marker_key(marker), "inputnorm")) {
    stop(sprintf("HistCross C5 must point to an <<inputnorm>> tab, but '%s' is not one.", input_sheet_name), call. = FALSE)
  }

  invisible(TRUE)
}

goecode_tool_inputnorm_data_file <- function(output_dir, input_sheet_name) {
  data_file <- file.path(output_dir, "inputnorm", "inputnorm_data.csv")
  if (!file.exists(data_file)) {
    stop(
      sprintf(
        "InputNorm data artifact is missing for tab '%s': %s. Run Update on the InputNorm tab first.",
        input_sheet_name,
        data_file
      ),
      call. = FALSE
    )
  }
  data_file
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

goecode_tool_histcross_alias_dictionary <- function(workbook_path) {
  aliases <- goecode_tool_alias_dictionary(workbook_path)
  mapped <- list(
    mapped_province = c("mapped_province", "Mapped_Province", "Mapped Province", "省映射", "标准省份"),
    mapped_prefecture = c("mapped_prefecture", "Mapped_Prefecture", "Mapped Prefecture", "市映射", "标准地市"),
    mapped_county = c("mapped_county", "Mapped_County", "Mapped County", "区县映射", "标准区县"),
    mapped_address = c("mapped_address", "Mapped_Address", "Mapped Address", "标准地址"),
    mapped_lon = c("mapped_lon", "mapped_longitude", "longitude", "lon", "lng", "经度"),
    mapped_lat = c("mapped_lat", "mapped_latitude", "latitude", "lat", "纬度"),
    mapped_confidence = c("mapped_confidence", "confidence", "match_confidence", "置信度"),
    mapped_source = c("mapped_source", "source", "match_source", "来源")
  )
  c(aliases, mapped)
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
    } else if (grepl("^mapped ", key)) {
      mapped[i] <- gsub(" ", "_", key)
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

  chunk_size <- 10000L
  starts <- seq.int(1L, nrow(data), by = chunk_size)
  for (start in starts) {
    end <- min(start + chunk_size - 1L, nrow(data))
    rows <- start:end
    quoted <- lapply(data, function(col) goecode_tool_csv_quote(col[rows]))
    lines <- do.call(paste, c(quoted, sep = ","))
    writeBin(charToRaw(enc2utf8(paste0(paste(lines, collapse = "\r\n"), "\r\n"))), con)
  }

  invisible(path)
}

goecode_tool_write_csv_line <- function(values, con) {
  line <- paste(vapply(values, goecode_tool_csv_quote, character(1)), collapse = ",")
  writeBin(charToRaw(enc2utf8(paste0(line, "\r\n"))), con)
}

goecode_tool_csv_quote <- function(value) {
  value <- as.character(value)
  value[is.na(value)] <- ""
  Encoding(value) <- "UTF-8"
  value <- gsub('"', '""', value, fixed = TRUE)
  paste0('"', value, '"')
}

goecode_tool_read_csv_fast <- function(source_path) {
  if (!requireNamespace("data.table", quietly = TRUE)) {
    return(NULL)
  }

  data <- tryCatch(
    data.table::fread(
      source_path,
      encoding = "UTF-8",
      na.strings = character(),
      colClasses = "character",
      data.table = FALSE,
      check.names = FALSE,
      blank.lines.skip = FALSE,
      showProgress = FALSE
    ),
    error = function(err) NULL
  )

  if (is.null(data)) {
    return(NULL)
  }

  data <- as.data.frame(data, stringsAsFactors = FALSE, check.names = FALSE)
  data[] <- lapply(data, function(col) {
    col <- as.character(col)
    col[is.na(col)] <- ""
    Encoding(col) <- "UTF-8"
    col
  })
  data
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

goecode_tool_read_workbook_table <- function(workbook_path, sheet_name, header_ref) {
  sheet <- goecode_tool_read_sheet_cached(workbook_path, sheet_name)
  header_col <- goecode_tool_ref_col(header_ref)
  header_row <- goecode_tool_ref_row(header_ref)

  headers <- character()
  for (col in header_col:(header_col + 200L)) {
    value <- trimws(goecode_tool_sheet_cell_value(sheet, paste0(goecode_tool_col_name(col), header_row)))
    if (!nzchar(value)) {
      if (length(headers) > 0) break
      next
    }
    headers <- c(headers, value)
  }
  if (length(headers) == 0) {
    stop(sprintf("No table headers found on sheet '%s' at %s.", sheet_name, header_ref), call. = FALSE)
  }
  headers <- make.unique(headers)

  rows <- list()
  blank_run <- 0L
  for (row in (header_row + 1L):(header_row + 1000000L)) {
    values <- vapply(seq_along(headers), function(i) {
      goecode_tool_sheet_cell_value(sheet, paste0(goecode_tool_col_name(header_col + i - 1L), row))
    }, character(1))
    if (!any(nzchar(trimws(values)))) {
      blank_run <- blank_run + 1L
      if (blank_run >= 20L) break
      next
    }
    blank_run <- 0L
    names(values) <- headers
    rows[[length(rows) + 1L]] <- as.data.frame(as.list(values), stringsAsFactors = FALSE, check.names = FALSE)
  }

  if (length(rows) == 0) {
    return(as.data.frame(setNames(replicate(length(headers), character(), simplify = FALSE), headers), stringsAsFactors = FALSE))
  }
  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

goecode_tool_ref_col <- function(ref) {
  letters <- toupper(gsub("[0-9]", "", ref))
  chars <- strsplit(letters, "", fixed = TRUE)[[1]]
  col <- 0L
  for (ch in chars) {
    col <- col * 26L + match(ch, LETTERS)
  }
  col
}

goecode_tool_ref_row <- function(ref) {
  as.integer(gsub("[A-Za-z]", "", ref))
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
    data <- goecode_tool_read_csv_fast(source_path)
    if (!is.null(data)) {
      return(data)
    }

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

goecode_tool_histcross_join <- function(current, prior, key_cols, mapped_cols) {
  current_keys_valid <- goecode_tool_complete_keys(current, key_cols)
  prior_keys_valid <- goecode_tool_complete_keys(prior, key_cols)
  current_key <- goecode_tool_join_key(current, key_cols)
  prior_key <- goecode_tool_join_key(prior, key_cols)

  prior_rows <- which(prior_keys_valid)
  warnings <- goecode_tool_histcross_duplicate_warnings(prior, prior_rows, prior_key)
  first_prior <- list()
  for (row in prior_rows) {
    key <- prior_key[[row]]
    if (is.null(first_prior[[key]])) {
      first_prior[[key]] <- row
    }
  }

  match_prior <- integer(nrow(current))
  for (row in seq_len(nrow(current))) {
    if (!current_keys_valid[[row]]) next
    key <- current_key[[row]]
    if (!is.null(first_prior[[key]])) {
      match_prior[[row]] <- first_prior[[key]]
    }
  }

  matched_rows <- which(match_prior > 0L)
  unmatched_rows <- which(match_prior == 0L)

  current_cols <- unique(c("Entry_Index", key_cols))
  current_cols <- current_cols[current_cols %in% names(current)]

  matched <- current[matched_rows, current_cols, drop = FALSE]
  if (length(mapped_cols) > 0) {
    mapped <- prior[match_prior[matched_rows], mapped_cols, drop = FALSE]
    rownames(mapped) <- NULL
    matched <- cbind(matched, mapped)
  }
  rownames(matched) <- NULL

  unmatched <- current[unmatched_rows, current_cols, drop = FALSE]
  rownames(unmatched) <- NULL

  list(matched = matched, unmatched = unmatched, warnings = warnings)
}

goecode_tool_complete_keys <- function(data, key_cols) {
  if (length(key_cols) == 0) {
    return(rep(FALSE, nrow(data)))
  }
  apply(data[, key_cols, drop = FALSE], 1, function(row) all(nzchar(trimws(as.character(row)))))
}

goecode_tool_join_key <- function(data, key_cols) {
  if (length(key_cols) == 0) {
    return(rep("", nrow(data)))
  }
  apply(data[, key_cols, drop = FALSE], 1, function(row) paste(enc2utf8(as.character(row)), collapse = "\r"))
}

goecode_tool_histcross_duplicate_warnings <- function(prior, prior_rows, prior_key) {
  if (length(prior_rows) == 0) {
    return(data.frame(
      warning_type = character(),
      source_rows = character(),
      join_key = character(),
      message = character(),
      stringsAsFactors = FALSE,
      check.names = FALSE
    ))
  }

  keys <- prior_key[prior_rows]
  dup_keys <- unique(keys[duplicated(keys)])
  if (length(dup_keys) == 0) {
    return(data.frame(
      warning_type = character(),
      source_rows = character(),
      join_key = character(),
      message = character(),
      stringsAsFactors = FALSE,
      check.names = FALSE
    ))
  }

  rows <- lapply(dup_keys, function(key) {
    source_rows <- prior_rows[keys == key] + 1L
    data.frame(
      warning_type = "duplicate_prior_key",
      source_rows = paste(source_rows, collapse = ";"),
      join_key = gsub("\r", " | ", key, fixed = TRUE),
      message = "Duplicate prior-year join key; first row was used for matching.",
      stringsAsFactors = FALSE,
      check.names = FALSE
    )
  })
  do.call(rbind, rows)
}

goecode_tool_read_workbook_table_range <- function(workbook_path, sheet_name, header_ref, max_cols = 200L, max_data_rows = NULL) {
  temp_dir <- tempfile("goecode_tool_range_")
  dir.create(temp_dir, recursive = TRUE, showWarnings = FALSE)
  on.exit(unlink(temp_dir, recursive = TRUE, force = TRUE), add = TRUE)

  workbook_xml <- utils::unzip(workbook_path, files = "xl/workbook.xml", exdir = temp_dir)
  workbook_rels <- utils::unzip(workbook_path, files = "xl/_rels/workbook.xml.rels", exdir = temp_dir)
  shared_xml <- suppressWarnings(utils::unzip(workbook_path, files = "xl/sharedStrings.xml", exdir = temp_dir))
  workbook <- paste(readLines(workbook_xml[1], warn = FALSE), collapse = "")
  rels <- paste(readLines(workbook_rels[1], warn = FALSE), collapse = "")
  sheet_file <- goecode_tool_sheet_file(workbook, rels, sheet_name)
  sheet_xml <- utils::unzip(workbook_path, files = sheet_file, exdir = temp_dir)
  sheet_text <- paste(readLines(sheet_xml[1], warn = FALSE), collapse = "")
  shared <- goecode_tool_shared_strings(shared_xml)

  header_col <- goecode_tool_ref_col(header_ref)
  header_row <- goecode_tool_ref_row(header_ref)
  max_col <- header_col + max_cols - 1L
  cells <- new.env(parent = emptyenv())
  row_parts <- strsplit(sheet_text, "<row", fixed = TRUE, useBytes = FALSE)[[1]]
  if (length(row_parts) > 1L) {
    for (row_part in row_parts[-1L]) {
      row_end <- regexpr("</row>", row_part, fixed = TRUE, useBytes = FALSE)[[1]]
      if (row_end < 0L) next
      row_text <- substr(row_part, 1L, row_end + 5L)
      row_ref <- sub('.*\\br="([0-9]+)".*', "\\1", substr(row_text, 1L, min(nchar(row_text), 200L)))
      row_num <- suppressWarnings(as.integer(row_ref))
      if (is.na(row_num) || row_num < header_row) next
      if (!is.null(max_data_rows) && row_num > header_row + max_data_rows + 20L) break

      cell_parts <- strsplit(row_text, "<c", fixed = TRUE, useBytes = FALSE)[[1]]
      if (length(cell_parts) <= 1L) next
      for (cell_part in cell_parts[-1L]) {
        close_cell <- regexpr("</c>", cell_part, fixed = TRUE, useBytes = FALSE)[[1]]
        if (close_cell < 0L) next
        cell <- paste0("<c", substr(cell_part, 1L, close_cell + 3L))
        ref <- sub('.*\\br="([^"]+)".*', "\\1", substr(cell, 1L, min(nchar(cell), 120L)))
        col <- goecode_tool_ref_col(ref)
        row <- goecode_tool_ref_row(ref)
        if (is.na(row) || is.na(col) || col < header_col || col > max_col || row < header_row) next
        cells[[ref]] <- goecode_tool_cell_tag_value(cell, shared)
      }
    }
  }

  headers <- character()
  for (col in header_col:max_col) {
    value <- trimws(as.character(cells[[paste0(goecode_tool_col_name(col), header_row)]] %||% ""))
    if (!nzchar(value)) {
      if (length(headers) > 0) break
      next
    }
    headers <- c(headers, value)
  }
  if (length(headers) == 0) {
    stop(sprintf("No table headers found on sheet '%s' at %s.", sheet_name, header_ref), call. = FALSE)
  }
  headers <- make.unique(headers)

  rows <- list()
  blank_run <- 0L
  current_row <- header_row + 1L
  repeat {
    values <- vapply(seq_along(headers), function(i) {
      ref <- paste0(goecode_tool_col_name(header_col + i - 1L), current_row)
      as.character(cells[[ref]] %||% "")
    }, character(1))
    if (!any(nzchar(trimws(values)))) {
      blank_run <- blank_run + 1L
      if (blank_run >= 20L) break
    } else {
      blank_run <- 0L
      names(values) <- headers
      rows[[length(rows) + 1L]] <- as.data.frame(as.list(values), stringsAsFactors = FALSE, check.names = FALSE)
    }
    current_row <- current_row + 1L
    if (current_row > 1048576L) break
  }

  if (length(rows) == 0) {
    return(as.data.frame(setNames(replicate(length(headers), character(), simplify = FALSE), headers), stringsAsFactors = FALSE, check.names = FALSE))
  }
  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

`%||%` <- function(x, y) {
  if (is.null(x)) y else x
}

goecode_tool_generalcrosswalk_input <- function(ctx) {
  source_path <- goecode_tool_cell(ctx$workbook_path, ctx$active_sheet, row = 14, col = 3)
  if (nzchar(source_path) && file.exists(source_path)) {
    input <- goecode_tool_read_source_table(source_path)
    attr(input, "source_label") <- source_path
    return(input)
  }

  input <- goecode_tool_read_workbook_table_range(ctx$workbook_path, ctx$active_sheet, "N14", max_cols = 29L)
  attr(input, "source_label") <- sprintf("%s!N14", ctx$active_sheet)
  input
}

goecode_tool_ensure_entry_index <- function(data) {
  data <- as.data.frame(data, stringsAsFactors = FALSE, check.names = FALSE)
  data[] <- lapply(data, function(col) {
    col <- as.character(col)
    col[is.na(col)] <- ""
    Encoding(col) <- "UTF-8"
    col
  })
  if (!"Entry_Index" %in% names(data)) {
    data <- data.frame(Entry_Index = seq_len(nrow(data)), data, stringsAsFactors = FALSE, check.names = FALSE)
  }
  data
}

goecode_tool_generalcrosswalk_existing_config <- function(workbook_path, sheet) {
  config <- goecode_tool_read_workbook_table_range(workbook_path, sheet, "B26", max_cols = 4L, max_data_rows = 500L)
  config <- config[!(tolower(config$header_raw) == "header_raw"), , drop = FALSE]
  config <- config[nzchar(config$header_raw), , drop = FALSE]
  names(config) <- c("header_raw", "Use", "Index Used", "Scoring Priority")
  rownames(config) <- NULL
  config
}

goecode_tool_generalcrosswalk_default_config <- function(headers, existing = NULL) {
  rows <- lapply(headers, function(header) {
    old <- NULL
    if (!is.null(existing) && nrow(existing) > 0) {
      idx <- match(header, existing$header_raw)
      if (!is.na(idx)) old <- existing[idx, , drop = FALSE]
    }
    guess <- goecode_tool_generalcrosswalk_guess_index(header)
    use <- if (guess %in% c("Province", "Prefecture", "County", "Address", "Insured")) "Y" else "N"
    data.frame(
      header_raw = header,
      Use = if (!is.null(old) && nzchar(old$Use)) old$Use else use,
      `Index Used` = if (!is.null(old) && nzchar(old$`Index Used`)) old$`Index Used` else guess,
      `Scoring Priority` = if (!is.null(old) && nzchar(old$`Scoring Priority`)) old$`Scoring Priority` else goecode_tool_generalcrosswalk_default_priority(guess),
      stringsAsFactors = FALSE,
      check.names = FALSE
    )
  })
  do.call(rbind, rows)
}

goecode_tool_generalcrosswalk_guess_index <- function(header) {
  key <- goecode_tool_normalize_header(header)
  if (grepl("province", key)) return("Province")
  if (grepl("prefecture|city", key)) return("Prefecture")
  if (grepl("county|district", key)) return("County")
  if (grepl("address", key)) return("Address")
  if (grepl("insured", key)) return("Insured")
  if (grepl("branch.*name", key)) return("Branch_Name")
  if (grepl("postal|postcode|zip", key)) return("Postal")
  if (grepl("branch.*code", key)) return("Branch_Code")
  ""
}

goecode_tool_generalcrosswalk_default_priority <- function(index_used) {
  switch(
    index_used,
    Province = "1",
    Prefecture = "1",
    County = "1",
    Address = "2",
    Insured = "3",
    ""
  )
}

goecode_tool_generalcrosswalk_read_config <- function(workbook_path, sheet) {
  config <- goecode_tool_generalcrosswalk_existing_config(workbook_path, sheet)
  if (nrow(config) == 0) {
    stop("GeneralCrossWalk configuration is missing in B26:E. Run Gather first, then configure the rows.", call. = FALSE)
  }
  config
}

goecode_tool_generalcrosswalk_selected_config <- function(config, input) {
  config$Use <- toupper(trimws(config$Use))
  config$`Index Used` <- vapply(config$`Index Used`, goecode_tool_generalcrosswalk_index_label, character(1))
  config$priority <- suppressWarnings(as.numeric(config$`Scoring Priority`))
  config$priority[is.na(config$priority)] <- 99
  selected <- config[config$Use %in% c("Y", "YES", "TRUE", "1"), , drop = FALSE]
  selected <- selected[nzchar(selected$header_raw) & nzchar(selected$`Index Used`), , drop = FALSE]
  selected <- selected[selected$`Index Used` %in% c("Province", "Prefecture", "County", "Address", "Insured"), , drop = FALSE]
  missing <- setdiff(selected$header_raw, names(input))
  if (length(missing) > 0) {
    stop(sprintf("GeneralCrossWalk selected header(s) are missing from input: %s", paste(missing, collapse = ", ")), call. = FALSE)
  }
  if (nrow(selected) == 0) {
    stop("GeneralCrossWalk requires at least one v1 config row marked Use=Y for Province, Prefecture, County, Address, or Insured.", call. = FALSE)
  }
  rownames(selected) <- NULL
  selected
}

goecode_tool_generalcrosswalk_index_label <- function(value) {
  key <- goecode_tool_normalize_header(value)
  if (key %in% c("province", "prov")) return("Province")
  if (key %in% c("prefecture", "city")) return("Prefecture")
  if (key %in% c("county", "district")) return("County")
  if (key %in% c("address", "addr")) return("Address")
  if (key %in% c("insured", "insured name")) return("Insured")
  if (key %in% c("branch name", "branch")) return("Branch_Name")
  if (key %in% c("postal", "postcode", "zip")) return("Postal")
  if (key %in% c("branch code")) return("Branch_Code")
  trimws(as.character(value))
}

goecode_tool_generalcrosswalk_target_resolution <- function(workbook_path, sheet) {
  value <- goecode_tool_generalcrosswalk_index_label(goecode_tool_cell(workbook_path, sheet, row = 42, col = 3))
  if (!value %in% c("Province", "Prefecture", "County")) {
    value <- "County"
  }
  value
}

goecode_tool_generalcrosswalk_review_threshold <- function(workbook_path, sheet) {
  value <- suppressWarnings(as.numeric(goecode_tool_cell(workbook_path, sheet, row = 43, col = 3)))
  if (is.na(value)) 70 else value
}

goecode_tool_generalcrosswalk_geo_index <- function(workbook_path) {
  geo <- goecode_tool_read_workbook_table_range(workbook_path, "Control", "O1", max_cols = 7L, max_data_rows = 5000L)
  required <- c("Postal", "province_name", "prefecture_name", "county_name", "province_short", "prefecture_short", "county_short")
  missing <- setdiff(required, names(geo))
  if (length(missing) > 0) {
    stop(sprintf("Control!O:U is missing required column(s): %s", paste(missing, collapse = ", ")), call. = FALSE)
  }
  geo <- geo[, required, drop = FALSE]
  geo$geo_row_id <- seq_len(nrow(geo))
  geo
}

goecode_tool_generalcrosswalk_zone_index <- function(workbook_path, geo) {
  zones <- goecode_tool_read_workbook_table_range(workbook_path, "Control", "AD1", max_cols = 4L, max_data_rows = 1000L)
  required <- c("zone_name", "province_name", "prefecture_name", "county_name")
  missing <- setdiff(required, names(zones))
  if (length(missing) > 0) {
    return(as.data.frame(setNames(replicate(length(c(required, "zone_status")), character(), simplify = FALSE), c(required, "zone_status")), stringsAsFactors = FALSE))
  }
  zones <- zones[, required, drop = FALSE]
  zones <- zones[nzchar(trimws(zones$zone_name)), , drop = FALSE]
  if (nrow(zones) == 0) {
    zones$zone_status <- character()
    return(zones)
  }
  key <- paste(geo$province_name, geo$prefecture_name, geo$county_name, sep = "\r")
  zone_key <- paste(zones$province_name, zones$prefecture_name, zones$county_name, sep = "\r")
  zones$zone_status <- ifelse(zone_key %in% key, "valid", "invalid_hierarchy")
  zones
}

goecode_tool_generalcrosswalk_ppc_candidates <- function(input, config, geo, zones) {
  ppc_config <- config[config$`Index Used` %in% c("Province", "Prefecture", "County"), , drop = FALSE]
  if (nrow(ppc_config) == 0) {
    rows <- lapply(input$Entry_Index, function(id) goecode_tool_generalcrosswalk_empty_candidate(id, "ppc"))
    return(goecode_tool_generalcrosswalk_candidate_frame(rows))
  }

  lookups <- list(
    Province = goecode_tool_generalcrosswalk_exact_lookup(geo, c("province_name", "province_short"), "geo_row_id"),
    Prefecture = goecode_tool_generalcrosswalk_exact_lookup(geo, c("prefecture_name", "prefecture_short"), "geo_row_id"),
    County = goecode_tool_generalcrosswalk_exact_lookup(geo, c("county_name", "county_short"), "geo_row_id"),
    Zone = goecode_tool_generalcrosswalk_exact_lookup(zones, "zone_name")
  )

  rows <- lapply(seq_len(nrow(input)), function(i) {
    goecode_tool_generalcrosswalk_ppc_row(input[i, , drop = FALSE], ppc_config, geo, zones, lookups)
  })
  goecode_tool_generalcrosswalk_candidate_frame(rows)
}

goecode_tool_generalcrosswalk_ppc_row <- function(row, ppc_config, geo, zones, lookups) {
  values <- split(ppc_config, ppc_config$`Index Used`)
  raw_values <- lapply(values, function(cfg) {
    vals <- unlist(row[, cfg$header_raw, drop = FALSE], use.names = FALSE)
    vals <- vals[goecode_tool_generalcrosswalk_has_value(vals)]
    unique(as.character(vals))
  })

  zone_hits <- character()
  if (!is.null(raw_values$County) && nrow(zones) > 0) {
    zone_hits <- goecode_tool_generalcrosswalk_lookup_ids(raw_values$County, lookups$Zone)
  }
  zone_hits <- unique(zone_hits)
  if (length(zone_hits) > 0) {
    return(goecode_tool_generalcrosswalk_zone_candidate(row$Entry_Index, "ppc", zones[zone_hits, , drop = FALSE], min(ppc_config$priority), "ppc_zone_match"))
  }

  if (sum(vapply(raw_values, length, integer(1))) == 0L) {
    return(goecode_tool_generalcrosswalk_empty_candidate(row$Entry_Index, "ppc"))
  }

  candidate_ids <- NULL
  evidence <- character()
  conflict <- character()
  for (kind in c("Province", "Prefecture", "County")) {
    vals <- raw_values[[kind]]
    if (is.null(vals) || length(vals) == 0) next
    ids <- goecode_tool_generalcrosswalk_lookup_ids(vals, lookups[[kind]])
    evidence <- c(evidence, sprintf("%s=%s", kind, paste(vals, collapse = "|")))
    if (length(ids) == 0) {
      conflict <- c(conflict, sprintf("No %s exact match for %s", kind, paste(vals, collapse = "|")))
      next
    }
    candidate_ids <- if (is.null(candidate_ids)) ids else intersect(candidate_ids, ids)
  }

  if (is.null(candidate_ids) || length(candidate_ids) == 0) {
    return(goecode_tool_generalcrosswalk_empty_candidate(row$Entry_Index, "ppc", "ppc_conflict", paste(conflict, collapse = "; ")))
  }

  cand <- geo[geo$geo_row_id %in% candidate_ids, , drop = FALSE]
  resolved <- goecode_tool_generalcrosswalk_resolve_geo(cand, raw_values)
  goecode_tool_generalcrosswalk_candidate(
    entry_index = row$Entry_Index,
    method = "ppc",
    province = resolved$province,
    prefecture = resolved$prefecture,
    county = resolved$county,
    level = resolved$level,
    status = resolved$status,
    priority = min(ppc_config$priority),
    confidence = goecode_tool_generalcrosswalk_base_confidence("ppc", resolved$level, FALSE),
    terms = paste(evidence, collapse = "; "),
    detail = resolved$detail,
    count = nrow(cand),
    source = paste(ppc_config$header_raw, collapse = ", ")
  )
}

goecode_tool_generalcrosswalk_exact_lookup <- function(table, fields, id_col = NULL) {
  if (nrow(table) == 0) return(list())
  row_ids <- if (!is.null(id_col) && id_col %in% names(table)) table[[id_col]] else seq_len(nrow(table))
  term_values <- character()
  term_ids <- integer()
  for (field in fields) {
    if (!field %in% names(table)) next
    values <- goecode_tool_generalcrosswalk_norm(table[[field]])
    keep <- nzchar(values)
    term_values <- c(term_values, values[keep])
    term_ids <- c(term_ids, row_ids[keep])
  }
  if (length(term_values) == 0) return(list())
  lapply(split(term_ids, term_values), unique)
}

goecode_tool_generalcrosswalk_lookup_ids <- function(values, lookup) {
  if (length(values) == 0 || length(lookup) == 0) return(integer())
  keys <- goecode_tool_generalcrosswalk_norm(values)
  keys <- keys[nzchar(keys)]
  if (length(keys) == 0) return(integer())
  unique(unlist(lookup[keys], use.names = FALSE))
}

goecode_tool_generalcrosswalk_text_candidates <- function(input, config, geo, zones, index_used) {
  text_config <- config[config$`Index Used` == index_used, , drop = FALSE]
  method <- tolower(index_used)
  if (nrow(text_config) == 0) {
    rows <- lapply(input$Entry_Index, function(id) goecode_tool_generalcrosswalk_empty_candidate(id, method))
    return(goecode_tool_generalcrosswalk_candidate_frame(rows))
  }

  row_texts <- vapply(seq_len(nrow(input)), function(i) {
    texts <- unlist(input[i, text_config$header_raw, drop = FALSE], use.names = FALSE)
    texts <- texts[goecode_tool_generalcrosswalk_has_value(texts)]
    paste(texts, collapse = " ")
  }, character(1))

  text_norms <- goecode_tool_generalcrosswalk_norm(row_texts)
  zone_hits <- goecode_tool_generalcrosswalk_row_hits(text_norms, zones, "zone_name")
  county_hits <- goecode_tool_generalcrosswalk_row_hits(text_norms, geo, c("county_name", "county_short"), "geo_row_id")
  prefecture_hits <- goecode_tool_generalcrosswalk_row_hits(text_norms, geo, c("prefecture_name", "prefecture_short"), "geo_row_id")
  province_hits <- goecode_tool_generalcrosswalk_row_hits(text_norms, geo, c("province_name", "province_short"), "geo_row_id")
  source_columns <- paste(text_config$header_raw, collapse = ", ")
  priority <- min(text_config$priority)

  rows <- lapply(seq_len(nrow(input)), function(i) {
    text <- row_texts[[i]]
    if (!goecode_tool_generalcrosswalk_has_value(text)) {
      return(goecode_tool_generalcrosswalk_empty_candidate(input$Entry_Index[[i]], method))
    }

    if (length(zone_hits[[i]]) > 0) {
      hits <- zones[zone_hits[[i]], , drop = FALSE]
      return(goecode_tool_generalcrosswalk_zone_candidate(input$Entry_Index[[i]], method, hits, priority, paste0(method, "_zone_match")))
    }

    if (length(county_hits[[i]]) > 0) {
      cand <- geo[geo$geo_row_id %in% county_hits[[i]], , drop = FALSE]
      if (length(province_hits[[i]]) > 0) {
        narrowed <- cand[cand$geo_row_id %in% province_hits[[i]], , drop = FALSE]
        if (nrow(narrowed) > 0) cand <- narrowed
      }
      if (length(prefecture_hits[[i]]) > 0) {
        narrowed <- cand[cand$geo_row_id %in% prefecture_hits[[i]], , drop = FALSE]
        if (nrow(narrowed) > 0) cand <- narrowed
      }
      return(goecode_tool_generalcrosswalk_geo_text_candidate(input$Entry_Index[[i]], method, cand, priority, "county", text, source_columns))
    }

    if (length(prefecture_hits[[i]]) > 0) {
      cand <- geo[geo$geo_row_id %in% prefecture_hits[[i]], , drop = FALSE]
      return(goecode_tool_generalcrosswalk_geo_text_candidate(input$Entry_Index[[i]], method, cand, priority, "prefecture", text, source_columns))
    }

    if (length(province_hits[[i]]) > 0) {
      cand <- geo[geo$geo_row_id %in% province_hits[[i]], , drop = FALSE]
      return(goecode_tool_generalcrosswalk_geo_text_candidate(input$Entry_Index[[i]], method, cand, priority, "province", text, source_columns))
    }

    goecode_tool_generalcrosswalk_empty_candidate(input$Entry_Index[[i]], method)
  })
  goecode_tool_generalcrosswalk_candidate_frame(rows)
}

goecode_tool_generalcrosswalk_row_hits <- function(text_norms, table, fields, id_col = NULL) {
  hits <- vector("list", length(text_norms))
  if (nrow(table) == 0 || length(text_norms) == 0) {
    return(hits)
  }

  text_norms <- goecode_tool_utf8_clean(text_norms)
  row_ids <- if (!is.null(id_col) && id_col %in% names(table)) table[[id_col]] else seq_len(nrow(table))
  term_values <- character()
  term_ids <- integer()
  for (field in fields) {
    if (!field %in% names(table)) next
    values <- goecode_tool_generalcrosswalk_norm(table[[field]])
    keep <- nzchar(values)
    term_values <- c(term_values, values[keep])
    term_ids <- c(term_ids, row_ids[keep])
  }
  if (length(term_values) == 0) {
    return(hits)
  }

  ids_by_term <- split(term_ids, term_values)
  for (term in names(ids_by_term)) {
    term <- goecode_tool_utf8_clean(term)
    if (!nzchar(term)) next
    matched <- grepl(term, text_norms, fixed = TRUE, useBytes = TRUE)
    if (!any(matched)) next
    ids <- unique(ids_by_term[[term]])
    for (row in which(matched)) {
      hits[[row]] <- c(hits[[row]], ids)
    }
  }

  lapply(hits, unique)
}

goecode_tool_generalcrosswalk_text_row <- function(entry_index, text, config, geo, zones, method) {
  if (!goecode_tool_generalcrosswalk_has_value(text)) {
    return(goecode_tool_generalcrosswalk_empty_candidate(entry_index, method))
  }

  zone_hits <- goecode_tool_generalcrosswalk_contains_rows(text, zones, "zone_name")
  if (nrow(zone_hits) > 0) {
    return(goecode_tool_generalcrosswalk_zone_candidate(entry_index, method, zone_hits, min(config$priority), paste0(method, "_zone_match")))
  }

  county_hits <- unique(goecode_tool_generalcrosswalk_contains_rows_multi(text, geo, c("county_name", "county_short")))
  if (length(county_hits) > 0) {
    narrowed <- goecode_tool_generalcrosswalk_filter_text_geo(text, geo[geo$geo_row_id %in% county_hits, , drop = FALSE])
    return(goecode_tool_generalcrosswalk_geo_text_candidate(entry_index, method, narrowed, min(config$priority), "county", text, paste(config$header_raw, collapse = ", ")))
  }

  prefecture_hits <- unique(goecode_tool_generalcrosswalk_contains_rows_multi(text, geo, c("prefecture_name", "prefecture_short")))
  if (length(prefecture_hits) > 0) {
    cand <- geo[geo$geo_row_id %in% prefecture_hits, , drop = FALSE]
    return(goecode_tool_generalcrosswalk_geo_text_candidate(entry_index, method, cand, min(config$priority), "prefecture", text, paste(config$header_raw, collapse = ", ")))
  }

  province_hits <- unique(goecode_tool_generalcrosswalk_contains_rows_multi(text, geo, c("province_name", "province_short")))
  if (length(province_hits) > 0) {
    cand <- geo[geo$geo_row_id %in% province_hits, , drop = FALSE]
    return(goecode_tool_generalcrosswalk_geo_text_candidate(entry_index, method, cand, min(config$priority), "province", text, paste(config$header_raw, collapse = ", ")))
  }

  goecode_tool_generalcrosswalk_empty_candidate(entry_index, method)
}

goecode_tool_generalcrosswalk_filter_text_geo <- function(text, candidates) {
  province_ids <- goecode_tool_generalcrosswalk_contains_rows_multi(text, candidates, c("province_name", "province_short"))
  if (length(province_ids) > 0) {
    candidates <- candidates[candidates$geo_row_id %in% province_ids, , drop = FALSE]
  }
  prefecture_ids <- goecode_tool_generalcrosswalk_contains_rows_multi(text, candidates, c("prefecture_name", "prefecture_short"))
  if (length(prefecture_ids) > 0) {
    candidates <- candidates[candidates$geo_row_id %in% prefecture_ids, , drop = FALSE]
  }
  candidates
}

goecode_tool_generalcrosswalk_geo_text_candidate <- function(entry_index, method, candidates, priority, requested_level, text, source_columns) {
  resolved <- goecode_tool_generalcrosswalk_resolve_geo(candidates, list())
  if (identical(requested_level, "province")) {
    resolved$prefecture <- ""
    resolved$county <- ""
    resolved$level <- "Province"
  } else if (identical(requested_level, "prefecture")) {
    resolved$county <- ""
    resolved$level <- "Prefecture"
  }
  confidence <- goecode_tool_generalcrosswalk_base_confidence(method, resolved$level, FALSE)
  goecode_tool_generalcrosswalk_candidate(
    entry_index = entry_index,
    method = method,
    province = resolved$province,
    prefecture = resolved$prefecture,
    county = resolved$county,
    level = resolved$level,
    status = paste0(method, "_", tolower(resolved$level), "_candidate"),
    priority = priority,
    confidence = confidence,
    terms = substr(text, 1L, 300L),
    detail = if (nrow(candidates) > 1) sprintf("%s candidates", nrow(candidates)) else "",
    count = nrow(candidates),
    source = source_columns
  )
}

goecode_tool_generalcrosswalk_zone_candidate <- function(entry_index, method, hits, priority, status) {
  unique_keys <- unique(paste(hits$province_name, hits$prefecture_name, hits$county_name, sep = "\r"))
  if (length(unique_keys) > 1) {
    return(goecode_tool_generalcrosswalk_candidate(
      entry_index = entry_index,
      method = method,
      province = "",
      prefecture = "",
      county = "",
      level = "",
      status = paste0(method, "_zone_ambiguous"),
      priority = priority,
      confidence = 0,
      terms = paste(unique(hits$zone_name), collapse = "|"),
      detail = "Multiple zone names mapped to different hierarchies.",
      count = nrow(hits),
      source = ""
    ))
  }
  hit <- hits[1, , drop = FALSE]
  goecode_tool_generalcrosswalk_candidate(
    entry_index = entry_index,
    method = method,
    province = hit$province_name,
    prefecture = hit$prefecture_name,
    county = hit$county_name,
    level = "County",
    status = status,
    priority = priority,
    confidence = goecode_tool_generalcrosswalk_base_confidence(method, "County", TRUE),
    terms = paste(unique(hits$zone_name), collapse = "|"),
    detail = if (any(hits$zone_status != "valid")) "Zone mapping is not found in Control O:U hierarchy." else "",
    count = nrow(hits),
    source = "zone_name"
  )
}

goecode_tool_generalcrosswalk_geo_exact_ids <- function(values, kind, geo) {
  fields <- switch(
    kind,
    Province = c("province_name", "province_short"),
    Prefecture = c("prefecture_name", "prefecture_short"),
    County = c("county_name", "county_short"),
    character()
  )
  value_keys <- goecode_tool_generalcrosswalk_norm(values)
  matched <- rep(FALSE, nrow(geo))
  for (field in fields) {
    matched <- matched | goecode_tool_generalcrosswalk_norm(geo[[field]]) %in% value_keys
  }
  geo$geo_row_id[matched]
}

goecode_tool_generalcrosswalk_contains_rows <- function(text, table, field) {
  if (nrow(table) == 0 || !field %in% names(table)) {
    return(table[0, , drop = FALSE])
  }
  text_norm <- goecode_tool_generalcrosswalk_norm(text)
  terms <- goecode_tool_generalcrosswalk_norm(table[[field]])
  keep <- nzchar(terms) & vapply(terms, function(term) grepl(term, text_norm, fixed = TRUE), logical(1))
  table[keep, , drop = FALSE]
}

goecode_tool_generalcrosswalk_contains_rows_multi <- function(text, table, fields) {
  if (nrow(table) == 0) return(integer())
  ids <- integer()
  for (field in fields) {
    if (!field %in% names(table)) next
    hits <- goecode_tool_generalcrosswalk_contains_rows(text, table, field)
    if (nrow(hits) > 0) ids <- c(ids, hits$geo_row_id)
  }
  unique(ids)
}

goecode_tool_generalcrosswalk_resolve_geo <- function(candidates, raw_values) {
  if (nrow(candidates) == 0) {
    return(list(province = "", prefecture = "", county = "", level = "", status = "no_candidate", detail = ""))
  }
  provinces <- unique(candidates$province_name)
  prefectures <- unique(paste(candidates$province_name, candidates$prefecture_name, sep = "\r"))
  counties <- unique(paste(candidates$province_name, candidates$prefecture_name, candidates$county_name, sep = "\r"))
  if (length(counties) == 1 && (!is.null(raw_values$County) || nrow(candidates) == 1)) {
    parts <- strsplit(counties[[1]], "\r", fixed = TRUE)[[1]]
    return(list(province = parts[[1]], prefecture = parts[[2]], county = parts[[3]], level = "County", status = "matched_county", detail = ""))
  }
  if (length(prefectures) == 1 && (length(counties) > 1 || !is.null(raw_values$Prefecture))) {
    parts <- strsplit(prefectures[[1]], "\r", fixed = TRUE)[[1]]
    return(list(province = parts[[1]], prefecture = parts[[2]], county = "", level = "Prefecture", status = "matched_prefecture", detail = sprintf("%s county candidates", length(counties))))
  }
  if (length(provinces) == 1) {
    return(list(province = provinces[[1]], prefecture = "", county = "", level = "Province", status = "matched_province", detail = sprintf("%s candidate rows", nrow(candidates))))
  }
  list(province = "", prefecture = "", county = "", level = "", status = "ambiguous", detail = sprintf("%s candidate rows", nrow(candidates)))
}

goecode_tool_generalcrosswalk_promote <- function(input, config, ppc, address, insured, target_resolution, threshold) {
  candidate_rows <- rbind(ppc, address, insured)
  candidates_by_entry <- split(candidate_rows, candidate_rows$Entry_Index)
  mapped_rows <- vector("list", nrow(input))
  review_rows <- list()
  warning_rows <- list()

  for (i in seq_len(nrow(input))) {
    entry_index <- input$Entry_Index[[i]]
    row_candidates <- candidates_by_entry[[as.character(entry_index)]]
    if (is.null(row_candidates)) {
      row_candidates <- candidate_rows[0, , drop = FALSE]
    }
    candidates <- row_candidates[nzchar(row_candidates$candidate_level), , drop = FALSE]
    diagnostics <- row_candidates[!nzchar(row_candidates$candidate_level) & nzchar(row_candidates$candidate_status), , drop = FALSE]
    result <- goecode_tool_generalcrosswalk_promote_row(entry_index, candidates, diagnostics, target_resolution, threshold)
    mapped_rows[[i]] <- result$mapped
    if (result$review) {
      review_rows[[length(review_rows) + 1L]] <- cbind(input[i, , drop = FALSE], result$review_row, stringsAsFactors = FALSE)
    }
    if (nzchar(result$mapped$mapping_warning)) {
      warning_rows[[length(warning_rows) + 1L]] <- data.frame(
        Entry_Index = entry_index,
        warning = result$mapped$mapping_warning,
        detail = result$mapped$mapping_conflict_detail,
        stringsAsFactors = FALSE,
        check.names = FALSE
      )
    }
  }

  mapped <- do.call(rbind, mapped_rows)
  review <- if (length(review_rows) == 0) {
    goecode_tool_generalcrosswalk_empty_review_frame(input)
  } else {
    do.call(rbind, review_rows)
  }
  warnings <- if (length(warning_rows) == 0) {
    data.frame(Entry_Index = character(), warning = character(), detail = character(), stringsAsFactors = FALSE, check.names = FALSE)
  } else {
    do.call(rbind, warning_rows)
  }
  list(mapped = mapped, review = review, warnings = warnings)
}

goecode_tool_generalcrosswalk_promote_row <- function(entry_index, candidates, diagnostics, target_resolution, threshold) {
  diagnostic_warning <- ""
  diagnostic_detail <- ""
  if (nrow(diagnostics) > 0) {
    diagnostic_warning <- "method_diagnostic"
    diagnostic_detail <- paste(sprintf("%s:%s:%s", diagnostics$method, diagnostics$candidate_status, diagnostics$candidate_detail), collapse = "; ")
  }

  if (nrow(candidates) == 0) {
    status <- if (nzchar(diagnostic_warning)) "unresolved_method_diagnostic" else "unresolved_no_candidate"
    detail <- if (nzchar(diagnostic_detail)) diagnostic_detail else "No v1 candidate was generated."
    mapped <- goecode_tool_generalcrosswalk_mapped_row(entry_index, "", "", "", "", "", 0, status, diagnostic_warning, detail)
    return(list(mapped = mapped, review = TRUE, review_row = goecode_tool_generalcrosswalk_review_row(mapped, "manual_no_candidate", detail)))
  }

  candidates$level_rank <- vapply(candidates$candidate_level, goecode_tool_generalcrosswalk_level_rank, numeric(1))
  candidates$method_rank <- match(candidates$method, c("ppc", "address", "insured"))
  candidates <- candidates[order(candidates$priority, -candidates$level_rank, candidates$method_rank), , drop = FALSE]
  winner <- candidates[1, , drop = FALSE]
  final <- winner
  warning <- ""
  detail <- character()

  if (nrow(candidates) > 1) {
    for (j in 2:nrow(candidates)) {
      other <- candidates[j, , drop = FALSE]
      if (goecode_tool_generalcrosswalk_compatible(final, other)) {
        if (goecode_tool_generalcrosswalk_level_rank(other$candidate_level) > goecode_tool_generalcrosswalk_level_rank(final$candidate_level)) {
          final <- other
        }
      } else {
        warning <- "conflict_resolved_by_priority"
        detail <- c(detail, sprintf("%s=%s/%s/%s", other$method, other$candidate_province, other$candidate_prefecture, other$candidate_county))
      }
    }
  }
  if (nzchar(diagnostic_warning)) {
    warning <- paste(c(warning, diagnostic_warning), collapse = ";")
    warning <- gsub("^;|;$", "", warning)
    detail <- c(detail, diagnostic_detail)
  }

  confidence <- as.numeric(final$confidence_score)
  if (nzchar(warning)) confidence <- max(0, confidence - 15)
  status <- paste0("mapped_", tolower(final$candidate_level))
  target_missed <- goecode_tool_generalcrosswalk_level_rank(final$candidate_level) < goecode_tool_generalcrosswalk_level_rank(target_resolution)
  if (target_missed) {
    warning <- paste(c(warning, "below_target_resolution"), collapse = ";")
    warning <- gsub("^;|;$", "", warning)
  }

  mapped <- goecode_tool_generalcrosswalk_mapped_row(
    entry_index = entry_index,
    province = final$candidate_province,
    prefecture = final$candidate_prefecture,
    county = final$candidate_county,
    resolution = final$candidate_level,
    method = final$method,
    confidence = confidence,
    status = status,
    warning = warning,
    detail = paste(detail, collapse = "; ")
  )

  review <- confidence < threshold || target_missed
  reason <- character()
  if (confidence < threshold) reason <- c(reason, "manual_low_confidence")
  if (target_missed) reason <- c(reason, "manual_below_target_resolution")
  if (nzchar(warning) && confidence < threshold) reason <- c(reason, "manual_warning")
  if (length(reason) == 0) reason <- "manual_review"

  list(mapped = mapped, review = review, review_row = goecode_tool_generalcrosswalk_review_row(mapped, paste(unique(reason), collapse = ";"), mapped$mapping_conflict_detail))
}

goecode_tool_generalcrosswalk_candidate <- function(entry_index, method, province, prefecture, county, level, status, priority, confidence, terms, detail, count, source) {
  data.frame(
    Entry_Index = as.character(entry_index),
    method = method,
    candidate_province = province,
    candidate_prefecture = prefecture,
    candidate_county = county,
    candidate_level = level,
    candidate_status = status,
    priority = priority,
    confidence_score = confidence,
    matched_terms = terms,
    candidate_detail = detail,
    candidate_count = count,
    source_columns = source,
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
}

goecode_tool_generalcrosswalk_empty_candidate <- function(entry_index, method, status = "", detail = "") {
  goecode_tool_generalcrosswalk_candidate(entry_index, method, "", "", "", "", status, 99, 0, "", detail, 0, "")
}

goecode_tool_generalcrosswalk_candidate_frame <- function(rows) {
  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

goecode_tool_generalcrosswalk_mapped_row <- function(entry_index, province, prefecture, county, resolution, method, confidence, status, warning, detail) {
  data.frame(
    Entry_Index = as.character(entry_index),
    mapped_province = province,
    mapped_prefecture = prefecture,
    mapped_county = county,
    mapped_resolution = resolution,
    mapping_method = method,
    confidence_score = confidence,
    mapping_status = status,
    mapping_warning = warning,
    mapping_conflict_detail = detail,
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
}

goecode_tool_generalcrosswalk_review_row <- function(mapped, reason, detail) {
  data.frame(
    proposed_mapped_province = mapped$mapped_province,
    proposed_mapped_prefecture = mapped$mapped_prefecture,
    proposed_mapped_county = mapped$mapped_county,
    proposed_resolution = mapped$mapped_resolution,
    confidence_score = mapped$confidence_score,
    manual_review_reason = reason,
    manual_review_detail = detail,
    winning_method = mapped$mapping_method,
    mapping_status = mapped$mapping_status,
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
}

goecode_tool_generalcrosswalk_empty_review_frame <- function(input) {
  review_cols <- data.frame(
    proposed_mapped_province = character(),
    proposed_mapped_prefecture = character(),
    proposed_mapped_county = character(),
    proposed_resolution = character(),
    confidence_score = character(),
    manual_review_reason = character(),
    manual_review_detail = character(),
    winning_method = character(),
    mapping_status = character(),
    stringsAsFactors = FALSE,
    check.names = FALSE
  )
  input[0, , drop = FALSE][, names(input), drop = FALSE]
  cbind(input[0, , drop = FALSE], review_cols, stringsAsFactors = FALSE)
}

goecode_tool_generalcrosswalk_mapped_output <- function(input, mapped) {
  mapped_cols <- mapped[, setdiff(names(mapped), "Entry_Index"), drop = FALSE]
  out <- cbind(input, mapped_cols, stringsAsFactors = FALSE)
  rownames(out) <- NULL
  out
}

goecode_tool_generalcrosswalk_compatible <- function(a, b) {
  pairs <- list(
    c(a$candidate_province, b$candidate_province),
    c(a$candidate_prefecture, b$candidate_prefecture),
    c(a$candidate_county, b$candidate_county)
  )
  all(vapply(pairs, function(x) !nzchar(x[[1]]) || !nzchar(x[[2]]) || identical(x[[1]], x[[2]]), logical(1)))
}

goecode_tool_generalcrosswalk_base_confidence <- function(method, level, zone = FALSE) {
  level <- as.character(level)
  score <- switch(
    method,
    ppc = switch(level, County = 95, Prefecture = 88, Province = 80, 0),
    address = switch(level, County = 86, Prefecture = 75, Province = 65, 0),
    insured = switch(level, County = 72, Prefecture = 62, Province = 55, 0),
    0
  )
  if (zone) min(100, score + 5) else score
}

goecode_tool_generalcrosswalk_level_rank <- function(level) {
  switch(as.character(level), Province = 1, Prefecture = 2, County = 3, 0)
}

goecode_tool_generalcrosswalk_has_value <- function(value) {
  value <- trimws(goecode_tool_utf8_clean(value))
  nzchar(value) & !tolower(value) %in% c("null", "na", "n/a", "nan")
}

goecode_tool_generalcrosswalk_norm <- function(value) {
  value <- goecode_tool_utf8_clean(value)
  value[is.na(value)] <- ""
  value <- trimws(value)
  value <- gsub("中国", "", value, fixed = TRUE)
  value <- gsub("[[:space:][:punct:]]+", "", value)
  value <- gsub("（|）|\\(|\\)|【|】|\\[|\\]", "", value)
  Encoding(value) <- "UTF-8"
  value
}

goecode_tool_utf8_clean <- function(value) {
  value <- as.character(value)
  value[is.na(value)] <- ""
  cleaned <- iconv(value, from = "UTF-8", to = "UTF-8", sub = "")
  fallback <- is.na(cleaned)
  if (any(fallback)) {
    cleaned[fallback] <- iconv(value[fallback], from = "", to = "UTF-8", sub = "")
  }
  cleaned[is.na(cleaned)] <- ""
  Encoding(cleaned) <- "UTF-8"
  cleaned
}
