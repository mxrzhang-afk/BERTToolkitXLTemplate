xl_pricing_tool_validate <- function(workbook_path, context = list(), action = NULL) {
  btk_require_file(workbook_path, "workbook")

  sheets <- xl_pricing_tool_sheet_names(workbook_path)
  if (length(sheets) == 0) {
    stop("Workbook package does not list any worksheets.", call. = FALSE)
  }

  if (!is.null(action) && nzchar(action)) {
    xl_pricing_tool_validate_action(workbook_path, action, context = context, sheets = sheets)
  }

  "OK"
}

xl_pricing_tool_validate_action <- function(workbook_path, action, context = list(), sheets = NULL) {
  if (is.null(sheets)) {
    sheets <- xl_pricing_tool_sheet_names(workbook_path)
  }

  action <- tolower(trimws(action))
  active_sheet <- if (!is.null(context$active_sheet) && nzchar(context$active_sheet)) {
    context$active_sheet
  } else {
    NULL
  }

  required_sheets <- switch(
    action,
    gather = character(),
    update = character(),
    build = character(),
    gnpi_gather = if (is.null(active_sheet)) "Input_GNPI" else active_sheet,
    gnpi_update = if (is.null(active_sheet)) "Input_GNPI" else active_sheet,
    gnpi_build = if (is.null(active_sheet)) "Input_GNPI" else active_sheet,
    agg_gather = if (is.null(active_sheet)) "Input_Agg" else active_sheet,
    agg_update = c(if (is.null(active_sheet)) "Input_Agg" else active_sheet, "Control_Module"),
    agg_build = if (is.null(active_sheet)) "Input_Agg" else active_sheet,
    profile_gather = if (is.null(active_sheet)) "Input_Profile" else active_sheet,
    profile_update = if (is.null(active_sheet)) "Input_Profile" else active_sheet,
    profile_build = c(if (is.null(active_sheet)) "Input_Profile" else active_sheet, "MBCurves", "ILFCurves"),
    riskfit_gather = if (is.null(active_sheet)) "Risk Loss Fitting" else active_sheet,
    riskfit_update = if (is.null(active_sheet)) "Risk Loss Fitting" else active_sheet,
    riskfit_build = if (is.null(active_sheet)) "Risk Loss Fitting" else active_sheet,
    cat_onlevel_gather = if (is.null(active_sheet)) "CatLossOnLevel" else active_sheet,
    cat_onlevel_update = if (is.null(active_sheet)) "CatLossOnLevel" else active_sheet,
    cat_onlevel_build = if (is.null(active_sheet)) "CatLossOnLevel" else active_sheet,
    xlsimulation_gather = if (is.null(active_sheet)) "Sim_Variations" else active_sheet,
    xlsimulation_update = if (is.null(active_sheet)) "Sim_Variations" else active_sheet,
    xlsimulation_build = if (is.null(active_sheet)) "Sim_Variations" else active_sheet,
    relativepricing_gather = if (is.null(active_sheet)) "RelativePrice" else active_sheet,
    relativepricing_update = c(if (is.null(active_sheet)) "RelativePrice" else active_sheet, "Input_Layers"),
    relativepricing_build = c(if (is.null(active_sheet)) "RelativePrice" else active_sheet, "Input_Layers"),
    rmstoyelt_gather = if (is.null(active_sheet)) "RMStoYELT" else active_sheet,
    rmstoyelt_update = if (is.null(active_sheet)) "RMStoYELT" else active_sheet,
    rmstoyelt_build = if (is.null(active_sheet)) "RMStoYELT" else active_sheet,
    character()
  )

  xl_pricing_tool_require_sheets(sheets, required_sheets, sprintf("XL pricing action '%s'", action))
  "OK"
}

xl_pricing_tool_require_sheets <- function(workbook_sheets, required_sheets, label = "workbook") {
  required_sheets <- unique(required_sheets[nzchar(required_sheets)])
  missing_sheets <- required_sheets[!tolower(required_sheets) %in% tolower(workbook_sheets)]
  if (length(missing_sheets) > 0) {
    stop(
      sprintf("%s is missing required sheet(s): %s", label, paste(missing_sheets, collapse = ", ")),
      call. = FALSE
    )
  }

  invisible(TRUE)
}

xl_pricing_tool_sheet_names <- function(workbook_path) {
  temp_dir <- tempfile("xl_pricing_tool_")
  dir.create(temp_dir, recursive = TRUE, showWarnings = FALSE)
  on.exit(unlink(temp_dir, recursive = TRUE, force = TRUE), add = TRUE)

  workbook_xml <- utils::unzip(workbook_path, files = "xl/workbook.xml", exdir = temp_dir)
  if (length(workbook_xml) == 0 || !file.exists(workbook_xml[1])) {
    stop("Workbook package is missing xl/workbook.xml.", call. = FALSE)
  }

  xml <- paste(readLines(workbook_xml[1], warn = FALSE), collapse = "")
  matches <- gregexpr("<sheet\\b[^>]*\\bname=\"[^\"]+\"", xml, perl = TRUE)[[1]]
  if (identical(matches[1], -1L)) {
    stop("Workbook package does not list any worksheets.", call. = FALSE)
  }

  sheet_tags <- regmatches(xml, list(matches))[[1]]
  sub('.*\\bname="([^"]+)".*', "\\1", sheet_tags)
}

xl_pricing_tool_gather <- function(workbook_path, output_dir = NULL, context = list()) {
  xl_pricing_tool_validate(workbook_path, context)
  xl_pricing_tool_action_pending("gather", workbook_path, output_dir, context)
}

xl_pricing_tool_update <- function(workbook_path, output_dir = NULL, context = list()) {
  xl_pricing_tool_validate(workbook_path, context)
  xl_pricing_tool_action_pending("update", workbook_path, output_dir, context)
}

xl_pricing_tool_build <- function(workbook_path, output_dir = NULL, context = list()) {
  xl_pricing_tool_validate(workbook_path, context)
  xl_pricing_tool_action_pending("build", workbook_path, output_dir, context)
}

xl_pricing_tool_run <- xl_pricing_tool_update

xl_pricing_tool_action_pending <- function(action, workbook_path, output_dir = NULL, context = list()) {
  if (is.null(output_dir) || identical(output_dir, "")) {
    output_dir <- if (!is.null(context$output_dir)) {
      context$output_dir
    } else {
      btk_default_run_dir("xl_pricing_tool", workbook_path)
    }
  }

  dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)

  paste(
    sprintf("XL pricing tool '%s' action is registered but not implemented yet.", action),
    "Next step: define tab-level behavior and expected outputs for this action.",
    sprintf("Output folder: %s", output_dir),
    sep = vb_newline()
  )
}

xl_pricing_tool_normalize_marker <- function(value) {
  value <- tolower(trimws(as.character(value)))
  value <- gsub("^<+|>+$", "", value)
  value <- gsub("[-_]+", " ", value)
  value <- gsub("\\s+", " ", value)
  trimws(value)
}

xl_pricing_tool_read_sheet <- function(workbook_path, sheet_name) {
  temp_dir <- tempfile("xl_pricing_tool_sheet_")
  dir.create(temp_dir, recursive = TRUE, showWarnings = FALSE)
  on.exit(unlink(temp_dir, recursive = TRUE, force = TRUE), add = TRUE)

  workbook_xml <- utils::unzip(workbook_path, files = "xl/workbook.xml", exdir = temp_dir)
  workbook_rels <- utils::unzip(workbook_path, files = "xl/_rels/workbook.xml.rels", exdir = temp_dir)
  shared_xml <- utils::unzip(workbook_path, files = "xl/sharedStrings.xml", exdir = temp_dir)

  workbook <- paste(readLines(workbook_xml[1], warn = FALSE), collapse = "")
  rels <- paste(readLines(workbook_rels[1], warn = FALSE), collapse = "")
  sheet_tags <- regmatches(workbook, gregexpr("<sheet\\b[^>]*>", workbook, perl = TRUE))[[1]]
  sheet_names <- vapply(sheet_tags, function(tag) {
    xl_pricing_tool_xml_unescape(sub('.*\\bname="([^"]+)".*', "\\1", tag))
  }, character(1))
  sheet_tag <- sheet_tags[tolower(sheet_names) == tolower(sheet_name)]
  if (length(sheet_tag) == 0) {
    stop(sprintf("Workbook is missing sheet: %s", sheet_name), call. = FALSE)
  }

  rid <- sub('.*r:id="([^"]+)".*', "\\1", sheet_tag[1])
  rel_tag <- regmatches(rels, gregexpr("<Relationship\\b[^>]*>", rels, perl = TRUE))[[1]]
  rel_tag <- rel_tag[grepl(sprintf('Id="%s"', rid), rel_tag, fixed = TRUE)]
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

  shared <- character()
  if (length(shared_xml) > 0 && file.exists(shared_xml[1])) {
    shared_text <- paste(readLines(shared_xml[1], warn = FALSE), collapse = "")
    shared_items <- regmatches(shared_text, gregexpr("<si>.*?</si>", shared_text, perl = TRUE))[[1]]
    if (!identical(shared_items[1], -1L)) {
      shared <- vapply(shared_items, function(item) {
        item <- gsub("<phoneticPr[^>]*/>", "", item)
        item <- gsub("<rPr>.*?</rPr>", "", item, perl = TRUE)
        text <- gsub("<[^>]+>", "", item)
        xl_pricing_tool_xml_unescape(text)
      }, character(1))
    }
  }

  sheet_text <- paste(readLines(sheet_xml[1], warn = FALSE), collapse = "")
  cell_tags <- regmatches(sheet_text, gregexpr("<c\\b[^>]*>.*?</c>", sheet_text, perl = TRUE))[[1]]
  cells <- list()
  if (!identical(cell_tags[1], -1L)) {
    for (tag in cell_tags) {
      ref <- sub('.*\\br="([^"]+)".*', "\\1", tag)
      type <- if (grepl('\\bt="s"', tag)) "s" else ""
      formula <- if (grepl("<f", tag, fixed = TRUE)) {
        f <- sub(".*<f[^>]*>(.*?)</f>.*", "\\1", tag)
        xl_pricing_tool_xml_unescape(f)
      } else {
        NA_character_
      }
      value <- if (grepl("<v>", tag, fixed = TRUE)) {
        v <- sub(".*<v>(.*?)</v>.*", "\\1", tag)
        if (identical(type, "s")) {
          idx <- suppressWarnings(as.integer(v) + 1L)
          if (!is.na(idx) && idx <= length(shared)) shared[idx] else v
        } else {
          xl_pricing_tool_xml_unescape(v)
        }
      } else if (grepl('\\bt="inlineStr"', tag) && grepl("<t", tag, fixed = TRUE)) {
        text <- sub(".*<t[^>]*>(.*?)</t>.*", "\\1", tag)
        xl_pricing_tool_xml_unescape(text)
      } else {
        ""
      }
      cells[[ref]] <- list(formula = formula, value = value)
    }
  }

  list(name = sheet_name, cells = cells)
}

xl_pricing_tool_action_sheet <- function(workbook_path, context, default_sheet) {
  sheet_name <- default_sheet
  if (!is.null(context$active_sheet) && nzchar(context$active_sheet)) {
    sheet_name <- context$active_sheet
  }
  xl_pricing_tool_read_sheet(workbook_path, sheet_name)
}

xl_pricing_tool_xml_unescape <- function(value) {
  value <- gsub("&lt;", "<", value, fixed = TRUE)
  value <- gsub("&gt;", ">", value, fixed = TRUE)
  value <- gsub("&amp;", "&", value, fixed = TRUE)
  value <- gsub("&quot;", '"', value, fixed = TRUE)
  value <- gsub("&apos;", "'", value, fixed = TRUE)
  value
}

xl_pricing_tool_cell_value <- function(sheet, ref) {
  cell <- sheet$cells[[ref]]
  if (is.null(cell)) {
    return("")
  }
  if (nzchar(cell$value)) {
    return(cell$value)
  }
  ""
}

xl_pricing_tool_column_values <- function(sheet, col, first_row) {
  refs <- names(sheet$cells)
  refs <- refs[grepl(sprintf("^%s[0-9]+$", col), refs)]
  rows <- suppressWarnings(as.integer(sub("^[A-Z]+", "", refs)))
  refs <- refs[!is.na(rows) & rows >= first_row]
  vapply(refs, function(ref) xl_pricing_tool_cell_value(sheet, ref), character(1))
}

xl_pricing_tool_read_vertical_values <- function(sheet, col, first_row) {
  values <- character()
  blank_run <- 0
  for (row in first_row:(first_row + 200)) {
    value <- trimws(xl_pricing_tool_cell_value(sheet, paste0(col, row)))
    if (!nzchar(value)) {
      blank_run <- blank_run + 1
      if (blank_run >= 3) break
      next
    }
    blank_run <- 0
    values <- c(values, value)
  }
  values
}

xl_pricing_tool_read_table <- function(sheet, header_row, start_col, end_col, first_data_row, key_cols) {
  headers <- vapply(start_col:end_col, function(col) {
    xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), header_row))
  }, character(1))
  headers <- make.names(headers, unique = TRUE)

  rows <- list()
  blank_run <- 0
  for (row in first_data_row:(first_data_row + 5000)) {
    values <- vapply(start_col:end_col, function(col) {
      xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), row))
    }, character(1))
    names(values) <- headers

    meaningful <- any(nzchar(trimws(values[names(values) %in% key_cols])))
    if (!meaningful) {
      blank_run <- blank_run + 1
      if (blank_run >= 20) break
      next
    }

    blank_run <- 0
    row_values <- unname(as.character(values))
    names(row_values) <- headers
    rows[[length(rows) + 1L]] <- as.data.frame(as.list(row_values), stringsAsFactors = FALSE)
  }

  if (length(rows) == 0) {
    return(data.frame())
  }
  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

xl_pricing_tool_col_name <- function(col) {
  name <- ""
  while (col > 0) {
    rem <- (col - 1) %% 26
    name <- paste0(intToUtf8(65 + rem), name)
    col <- (col - rem - 1) %/% 26
  }
  name
}
