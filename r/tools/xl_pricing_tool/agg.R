xl_pricing_tool_agg_update <- function(workbook_path, output_dir = NULL, context = list()) {
  xl_pricing_tool_validate(workbook_path, context)

  if (is.null(output_dir) || identical(output_dir, "")) {
    output_dir <- if (!is.null(context$output_dir)) {
      context$output_dir
    } else {
      btk_default_run_dir("xl_pricing_tool", workbook_path)
    }
  }

  agg_dir <- file.path(output_dir, "agg")
  dir.create(agg_dir, recursive = TRUE, showWarnings = FALSE)

  agg_sheet <- xl_pricing_tool_action_sheet(workbook_path, context, "Input_Agg")
  control_sheet <- xl_pricing_tool_read_sheet(workbook_path, "Control_Module")

  aggregate_input <- xl_pricing_tool_agg_read_input(agg_sheet)
  keyzone <- xl_pricing_tool_agg_read_keyzone(control_sheet)
  agg_output <- xl_pricing_tool_agg_summarize(aggregate_input, keyzone)

  output_file <- file.path(agg_dir, "agg_output.csv")
  utils::write.csv(agg_output, output_file, row.names = FALSE, na = "")

  paste(
    "Aggregate update completed.",
    sprintf("Output folder: %s", output_dir),
    sprintf("Input rows: %s", nrow(aggregate_input)),
    sprintf("Output rows: %s", nrow(agg_output)),
    sprintf("Aggregate output: %s", output_file),
    "CSV output is handled by the VBA client after BERT returns.",
    sep = vb_newline()
  )
}

xl_pricing_tool_agg_gather <- function(workbook_path, output_dir = NULL, context = list()) {
  "No gather action defined for Input_Agg because Gather From is blank."
}

xl_pricing_tool_agg_build <- function(workbook_path, output_dir = NULL, context = list()) {
  "No build output defined for Input_Agg."
}

xl_pricing_tool_agg_read_input <- function(sheet) {
  start_col <- 2
  province_col <- 3
  first_year_col <- 4
  header_row <- 12
  first_data_row <- 13

  peril_header <- trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(start_col), header_row)))
  province_header <- trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(province_col), header_row)))
  if (tolower(peril_header) != "peril" || tolower(province_header) != "province") {
    stop("Input_Agg must have Peril and Province headers in B12:C12.", call. = FALSE)
  }

  year_cols <- integer()
  years <- integer()
  blank_run <- 0
  for (col in first_year_col:(first_year_col + 60)) {
    value <- trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), header_row)))
    year <- suppressWarnings(as.integer(value))
    if (!nzchar(value)) {
      blank_run <- blank_run + 1
      if (blank_run >= 3) break
      next
    }
    blank_run <- 0
    if (is.na(year)) {
      break
    }
    year_cols <- c(year_cols, col)
    years <- c(years, year)
  }

  if (length(year_cols) == 0) {
    stop("Input_Agg has no numeric year headers starting at D12.", call. = FALSE)
  }

  rows <- list()
  blank_run <- 0
  for (row in first_data_row:(first_data_row + 5000)) {
    peril <- trimws(xl_pricing_tool_cell_value(sheet, paste0("B", row)))
    province <- trimws(xl_pricing_tool_cell_value(sheet, paste0("C", row)))

    if (!nzchar(peril) && !nzchar(province)) {
      blank_run <- blank_run + 1
      if (blank_run >= 20) break
      next
    }
    blank_run <- 0

    if (!nzchar(peril) || !nzchar(province)) {
      stop(sprintf("Input_Agg row %s has blank Peril or Province.", row), call. = FALSE)
    }

    peril <- toupper(peril)
    if (!peril %in% c("EQ", "WF")) {
      stop(sprintf("Input_Agg row %s has unsupported Peril '%s'. Supported perils: EQ, WF.", row, peril), call. = FALSE)
    }

    for (i in seq_along(year_cols)) {
      raw_value <- trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(year_cols[[i]]), row)))
      agg <- suppressWarnings(as.numeric(raw_value))
      if (is.na(agg)) {
        if (nzchar(raw_value)) {
          stop(sprintf("Input_Agg cell %s%s is not numeric: %s", xl_pricing_tool_col_name(year_cols[[i]]), row, raw_value), call. = FALSE)
        }
        agg <- 0
      }

      rows[[length(rows) + 1L]] <- data.frame(
        Peril = peril,
        Province = province,
        TreatyYear = years[[i]],
        Agg = agg,
        stringsAsFactors = FALSE
      )
    }
  }

  if (length(rows) == 0) {
    stop("Input_Agg has no aggregate input rows.", call. = FALSE)
  }

  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

xl_pricing_tool_agg_read_keyzone <- function(sheet) {
  headers <- vapply(11:15, function(col) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), 15)))
  }, character(1))
  expected <- c("Province", "Peril", "AIR", "RMS", "Blended")
  if (!all(headers == expected)) {
    stop("Control_Module keyzone table must have Province, Peril, AIR, RMS, Blended headers in K15:O15.", call. = FALSE)
  }

  rows <- list()
  blank_run <- 0
  for (row in 16:(16 + 5000)) {
    province <- trimws(xl_pricing_tool_cell_value(sheet, paste0("K", row)))
    peril <- trimws(xl_pricing_tool_cell_value(sheet, paste0("L", row)))

    if (!nzchar(province) && !nzchar(peril)) {
      blank_run <- blank_run + 1
      if (blank_run >= 20) break
      next
    }
    blank_run <- 0

    if (!nzchar(province) || !nzchar(peril)) {
      stop(sprintf("Control_Module keyzone row %s has blank Province or Peril.", row), call. = FALSE)
    }

    peril <- toupper(peril)
    if (!peril %in% c("EQ", "WF")) {
      stop(sprintf("Control_Module keyzone row %s has unsupported Peril '%s'. Supported perils: EQ, WF.", row, peril), call. = FALSE)
    }

    rows[[length(rows) + 1L]] <- data.frame(
      Province = province,
      Peril = peril,
      AIR = xl_pricing_tool_agg_flag(sheet, "M", row),
      RMS = xl_pricing_tool_agg_flag(sheet, "N", row),
      Blended = xl_pricing_tool_agg_flag(sheet, "O", row),
      stringsAsFactors = FALSE
    )
  }

  if (length(rows) == 0) {
    stop("Control_Module keyzone table has no rows.", call. = FALSE)
  }

  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

xl_pricing_tool_agg_flag <- function(sheet, col, row) {
  value <- suppressWarnings(as.numeric(trimws(xl_pricing_tool_cell_value(sheet, paste0(col, row)))))
  !is.na(value) && identical(value, 1)
}

xl_pricing_tool_agg_summarize <- function(aggregate_input, keyzone) {
  combos <- unique(aggregate_input[, c("Province", "Peril")])
  keyzone_key <- paste(keyzone$Province, keyzone$Peril, sep = "\r")
  input_key <- paste(combos$Province, combos$Peril, sep = "\r")
  missing <- combos[!input_key %in% keyzone_key, , drop = FALSE]
  if (nrow(missing) > 0) {
    missing_text <- paste(sprintf("%s/%s", missing$Province, missing$Peril), collapse = ", ")
    stop(sprintf("Keyzone table is missing Province/Peril combinations: %s", missing_text), call. = FALSE)
  }

  years <- sort(unique(aggregate_input$TreatyYear))
  perils <- c("EQ", "WF")
  methods <- c("Nationwide", "AIR", "RMS", "Blended")
  rows <- list()

  for (peril in perils) {
    for (year in years) {
      peril_year <- aggregate_input[aggregate_input$Peril == peril & aggregate_input$TreatyYear == year, , drop = FALSE]
      for (method in methods) {
        if (identical(method, "Nationwide")) {
          method_rows <- peril_year
        } else {
          kz <- keyzone[keyzone$Peril == peril & keyzone[[method]], c("Province", "Peril"), drop = FALSE]
          method_rows <- merge(peril_year, kz, by = c("Province", "Peril"))
        }

        rows[[length(rows) + 1L]] <- data.frame(
          Peril = peril,
          TreatyYear = year,
          Method = method,
          AggSum = sum(method_rows$Agg, na.rm = TRUE),
          stringsAsFactors = FALSE
        )
      }
    }
  }

  out <- do.call(rbind, rows)
  out <- out[order(out$Peril, out$TreatyYear, match(out$Method, methods)), , drop = FALSE]
  rownames(out) <- NULL
  out
}
