xl_pricing_tool_cat_onlevel_gather <- function(workbook_path, output_dir = NULL, context = list()) {
  xl_pricing_tool_validate(workbook_path, context)

  sheet <- xl_pricing_tool_action_sheet(workbook_path, context, "CatLossOnLevel")
  sources <- xl_pricing_tool_cat_onlevel_source_tabs(sheet)
  missing_source_names <- names(sources)[!nzchar(sources)]
  if (length(missing_source_names) > 0) {
    return(sprintf(
      "CatOnLevel gather needs source tab names in C6:C8. Please fill in: %s.",
      paste(missing_source_names, collapse = ", ")
    ))
  }

  workbook_sheets <- xl_pricing_tool_sheet_names(workbook_path)
  missing_tabs <- unname(sources)[!tolower(unname(sources)) %in% tolower(workbook_sheets)]
  if (length(missing_tabs) > 0) {
    stop(sprintf(
      "CatOnLevel source tab(s) do not exist in the active workbook: %s",
      paste(missing_tabs, collapse = ", ")
    ), call. = FALSE)
  }

  gnpi_source <- sources[["GNPI"]]
  cpi_source <- sources[["CPI"]]
  aggregate_source <- sources[["Aggregate"]]
  gnpi_sheet <- xl_pricing_tool_read_sheet(workbook_path, gnpi_source)
  cpi_sheet <- xl_pricing_tool_read_sheet(workbook_path, cpi_source)
  agg_sheet <- xl_pricing_tool_read_sheet(workbook_path, aggregate_source)

  checks <- c(
    xl_pricing_tool_cat_onlevel_check_gnpi(gnpi_sheet, gnpi_source),
    xl_pricing_tool_cat_onlevel_check_cpi(cpi_sheet, cpi_source),
    xl_pricing_tool_cat_onlevel_check_agg(agg_sheet, aggregate_source)
  )

  paste(
    "CatOnLevel gather validation completed.",
    sprintf("Active sheet: %s", sheet$name),
    sprintf("GNPI source: %s", gnpi_source),
    sprintf("CPI source: %s", cpi_source),
    sprintf("Aggregate source: %s", aggregate_source),
    paste(checks, collapse = vb_newline()),
    sep = vb_newline()
  )
}

xl_pricing_tool_cat_onlevel_update <- function(workbook_path, output_dir = NULL, context = list()) {
  xl_pricing_tool_validate(workbook_path, context)

  if (is.null(output_dir) || identical(output_dir, "")) {
    output_dir <- if (!is.null(context$output_dir)) {
      context$output_dir
    } else {
      btk_default_run_dir("xl_pricing_tool", workbook_path)
    }
  }

  cat_dir <- file.path(output_dir, "cat_onlevel")
  dir.create(cat_dir, recursive = TRUE, showWarnings = FALSE)

  sheet <- xl_pricing_tool_action_sheet(workbook_path, context, "CatLossOnLevel")
  sources <- xl_pricing_tool_cat_onlevel_source_tabs(sheet)
  xl_pricing_tool_cat_onlevel_validate_sources(workbook_path, sources)

  gnpi_sheet <- xl_pricing_tool_read_sheet(workbook_path, sources[["GNPI"]])
  cpi_sheet <- xl_pricing_tool_read_sheet(workbook_path, sources[["CPI"]])
  agg_sheet <- xl_pricing_tool_read_sheet(workbook_path, sources[["Aggregate"]])
  control_sheet <- xl_pricing_tool_read_sheet(workbook_path, sources[["CPI"]])

  events <- xl_pricing_tool_cat_onlevel_read_events(sheet)
  aggregate <- xl_pricing_tool_cat_onlevel_read_aggregate(agg_sheet)
  gnpi <- xl_pricing_tool_cat_onlevel_read_gnpi(gnpi_sheet)
  cpi <- xl_pricing_tool_cat_onlevel_read_cpi(cpi_sheet)
  keyzone <- xl_pricing_tool_cat_onlevel_read_keyzone(control_sheet)
  output <- xl_pricing_tool_cat_onlevel_calculate(events, aggregate, gnpi, cpi, keyzone)

  output_file <- file.path(cat_dir, "cat_onlevel_output.csv")
  utils::write.table(
    output[, c("from_Expo", "to_Expo"), drop = FALSE],
    file = output_file,
    sep = ",",
    row.names = FALSE,
    col.names = FALSE,
    na = "",
    quote = TRUE
  )

  warning_count <- sum(nzchar(output$warning))
  paste(
    "CatOnLevel update completed.",
    sprintf("Output folder: %s", output_dir),
    sprintf("Active sheet: %s", sheet$name),
    sprintf("Input rows: %s", nrow(events)),
    sprintf("Output rows: %s", nrow(output)),
    sprintf("Rows with warnings: %s", warning_count),
    sprintf("CatOnLevel output: %s", output_file),
    "CSV output is handled by the VBA client after BERT returns.",
    sep = vb_newline()
  )
}

xl_pricing_tool_cat_onlevel_build <- function(workbook_path, output_dir = NULL, context = list()) {
  "No build action defined for CatOnLevel."
}

xl_pricing_tool_cat_onlevel_source_tabs <- function(sheet) {
  c(
    GNPI = trimws(xl_pricing_tool_cell_value(sheet, "C6")),
    CPI = trimws(xl_pricing_tool_cell_value(sheet, "C7")),
    Aggregate = trimws(xl_pricing_tool_cell_value(sheet, "C8"))
  )
}

xl_pricing_tool_cat_onlevel_validate_sources <- function(workbook_path, sources) {
  missing_source_names <- names(sources)[!nzchar(sources)]
  if (length(missing_source_names) > 0) {
    stop(sprintf(
      "CatOnLevel update needs source tab names in C6:C8. Please fill in: %s.",
      paste(missing_source_names, collapse = ", ")
    ), call. = FALSE)
  }

  workbook_sheets <- xl_pricing_tool_sheet_names(workbook_path)
  missing_tabs <- unname(sources)[!tolower(unname(sources)) %in% tolower(workbook_sheets)]
  if (length(missing_tabs) > 0) {
    stop(sprintf(
      "CatOnLevel source tab(s) do not exist in the active workbook: %s",
      paste(missing_tabs, collapse = ", ")
    ), call. = FALSE)
  }
}

xl_pricing_tool_cat_onlevel_check_gnpi <- function(sheet, sheet_name) {
  expected <- c("LOB", "TreatyYear", "Type", "GNPI_Original", "Adjustment", "GNPI", "Latest_Current", "Latest_Prior")
  headers <- vapply(2:9, function(col) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), 12)))
  }, character(1))
  missing <- setdiff(expected, headers)
  if (length(missing) > 0) {
    stop(sprintf("GNPI source '%s' is missing expected B12:I12 header(s): %s", sheet_name, paste(missing, collapse = ", ")), call. = FALSE)
  }

  rows <- xl_pricing_tool_cat_onlevel_count_rows(sheet, first_row = 13, cols = c("B", "C", "D", "E", "G"))
  if (rows == 0) {
    stop(sprintf("GNPI source '%s' has no data in B13:I...", sheet_name), call. = FALSE)
  }
  sprintf("GNPI source range valid: %s rows found in '%s'!B13:I...", rows, sheet_name)
}

xl_pricing_tool_cat_onlevel_check_cpi <- function(sheet, sheet_name) {
  expected <- c("Year", "Inflat. %", "Inflat. Factor_Current", "Inflat. Factor_Prior")
  headers <- vapply(5:8, function(col) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), 15)))
  }, character(1))
  missing <- setdiff(expected, headers)
  if (length(missing) > 0) {
    stop(sprintf("CPI source '%s' is missing expected E15:H15 header(s): %s", sheet_name, paste(missing, collapse = ", ")), call. = FALSE)
  }

  rows <- xl_pricing_tool_cat_onlevel_count_rows(sheet, first_row = 16, cols = c("E", "G", "H"))
  if (rows == 0) {
    stop(sprintf("CPI source '%s' has no data in E16:H...", sheet_name), call. = FALSE)
  }
  sprintf("CPI source range valid: %s rows found in '%s'!E16:H...", rows, sheet_name)
}

xl_pricing_tool_cat_onlevel_check_agg <- function(sheet, sheet_name) {
  peril_header <- trimws(xl_pricing_tool_cell_value(sheet, "B12"))
  province_header <- trimws(xl_pricing_tool_cell_value(sheet, "C12"))
  if (peril_header != "Peril" || province_header != "Province") {
    stop(sprintf("Aggregate source '%s' must have Peril and Province headers in B12:C12.", sheet_name), call. = FALSE)
  }

  year_headers <- vapply(4:19, function(col) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), 12)))
  }, character(1))
  numeric_years <- suppressWarnings(as.integer(year_headers))
  if (!any(!is.na(numeric_years))) {
    stop(sprintf("Aggregate source '%s' has no numeric year headers in D12:S12.", sheet_name), call. = FALSE)
  }

  rows <- xl_pricing_tool_cat_onlevel_count_rows(sheet, first_row = 13, cols = c("B", "C", "D", "E", "F"))
  if (rows == 0) {
    stop(sprintf("Aggregate source '%s' has no data in B13:S...", sheet_name), call. = FALSE)
  }
  sprintf("Aggregate source range valid: %s rows found in '%s'!B13:S...", rows, sheet_name)
}

xl_pricing_tool_cat_onlevel_count_rows <- function(sheet, first_row, cols, max_rows = 10000) {
  count <- 0L
  blank_run <- 0L
  for (row in first_row:(first_row + max_rows)) {
    values <- vapply(cols, function(col) {
      trimws(xl_pricing_tool_cell_value(sheet, paste0(col, row)))
    }, character(1))
    if (!any(nzchar(values))) {
      blank_run <- blank_run + 1L
      if (blank_run >= 20L) {
        break
      }
      next
    }
    blank_run <- 0L
    count <- count + 1L
  }
  count
}

xl_pricing_tool_cat_onlevel_read_events <- function(sheet) {
  expected <- c(
    "AsAt", "ClaimID", "Method", "AY", "Affected_Provinces", "Peril",
    "from_year", "to_year", "Actual_Incurred",
    "In Prior/Current", "Prior/Current Incurred"
  )
  headers <- vapply(2:12, function(col) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), 12)))
  }, character(1))
  if (!all(headers == expected)) {
    stop("CatOnLevel event table must have expected headers in B12:L12.", call. = FALSE)
  }

  rows <- list()
  blank_run <- 0L
  for (row in 13:(13 + 10000)) {
    as_at <- trimws(xl_pricing_tool_cell_value(sheet, paste0("B", row)))
    claim_id <- trimws(xl_pricing_tool_cell_value(sheet, paste0("C", row)))
    method <- trimws(xl_pricing_tool_cell_value(sheet, paste0("D", row)))
    if (!nzchar(as_at) && !nzchar(claim_id) && !nzchar(method)) {
      blank_run <- blank_run + 1L
      if (blank_run >= 20L) break
      next
    }
    blank_run <- 0L

    rows[[length(rows) + 1L]] <- data.frame(
      Row = row,
      AsAt = as_at,
      ClaimID = claim_id,
      Method = method,
      AY = xl_pricing_tool_cat_onlevel_number(sheet, "E", row, "AY", required = FALSE),
      Affected_Provinces = trimws(xl_pricing_tool_cell_value(sheet, paste0("F", row))),
      Peril = toupper(trimws(xl_pricing_tool_cell_value(sheet, paste0("G", row)))),
      from_year = as.integer(xl_pricing_tool_cat_onlevel_number(sheet, "H", row, "from_year", required = TRUE)),
      to_year = as.integer(xl_pricing_tool_cat_onlevel_number(sheet, "I", row, "to_year", required = TRUE)),
      Actual_Incurred = xl_pricing_tool_cat_onlevel_number(sheet, "J", row, "Actual_Incurred", required = FALSE),
      In_Prior_Current = xl_pricing_tool_cat_onlevel_number(sheet, "K", row, "In Prior/Current", required = FALSE),
      stringsAsFactors = FALSE
    )
  }

  if (length(rows) == 0) {
    stop("CatOnLevel event table has no input rows.", call. = FALSE)
  }
  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

xl_pricing_tool_cat_onlevel_read_aggregate <- function(sheet) {
  year_cols <- 4:19
  years <- suppressWarnings(as.integer(vapply(year_cols, function(col) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), 12)))
  }, character(1))))
  valid <- !is.na(years)
  year_cols <- year_cols[valid]
  years <- years[valid]
  if (length(years) == 0) {
    stop("Aggregate source has no numeric year headers in D12:S12.", call. = FALSE)
  }

  rows <- list()
  blank_run <- 0L
  for (row in 13:(13 + 10000)) {
    peril <- toupper(trimws(xl_pricing_tool_cell_value(sheet, paste0("B", row))))
    province <- trimws(xl_pricing_tool_cell_value(sheet, paste0("C", row)))
    if (!nzchar(peril) && !nzchar(province)) {
      blank_run <- blank_run + 1L
      if (blank_run >= 20L) break
      next
    }
    blank_run <- 0L

    for (i in seq_along(year_cols)) {
      value <- suppressWarnings(as.numeric(gsub(",", "", trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(year_cols[[i]]), row))), fixed = TRUE)))
      rows[[length(rows) + 1L]] <- data.frame(
        Peril = peril,
        Province = province,
        ProvinceKey = xl_pricing_tool_cat_onlevel_key(province),
        Year = years[[i]],
        Exposure = ifelse(is.na(value), 0, value),
        stringsAsFactors = FALSE
      )
    }
  }

  if (length(rows) == 0) {
    stop("Aggregate source has no input rows.", call. = FALSE)
  }
  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

xl_pricing_tool_cat_onlevel_read_gnpi <- function(sheet) {
  rows <- list()
  blank_run <- 0L
  for (row in 13:(13 + 5000)) {
    lob <- trimws(xl_pricing_tool_cell_value(sheet, paste0("B", row)))
    year_raw <- trimws(xl_pricing_tool_cell_value(sheet, paste0("C", row)))
    type <- trimws(xl_pricing_tool_cell_value(sheet, paste0("D", row)))
    if (!nzchar(lob) && !nzchar(year_raw) && !nzchar(type)) {
      blank_run <- blank_run + 1L
      if (blank_run >= 20L) break
      next
    }
    blank_run <- 0L

    year <- suppressWarnings(as.integer(year_raw))
    original <- xl_pricing_tool_cat_onlevel_number(sheet, "E", row, "GNPI_Original", required = FALSE)
    adjustment <- xl_pricing_tool_cat_onlevel_number(sheet, "F", row, "Adjustment", required = FALSE)
    gnpi <- xl_pricing_tool_cat_onlevel_number(sheet, "G", row, "GNPI", required = FALSE)
    value <- if (is.finite(gnpi) && gnpi > 0) gnpi else original * ifelse(is.finite(adjustment) && adjustment > 0, adjustment, 1)
    if (!is.na(year) && is.finite(value) && value > 0) {
      rows[[length(rows) + 1L]] <- data.frame(Year = year, Type = type, GNPI = value, stringsAsFactors = FALSE)
    }
  }
  if (length(rows) == 0) {
    stop("GNPI source has no usable GNPI rows.", call. = FALSE)
  }
  by_type <- stats::aggregate(GNPI ~ Year + Type, data = do.call(rbind, rows), FUN = sum)
  by_type$TypePriority <- match(tolower(by_type$Type), c("actual", "revised", "estimate"))
  by_type$TypePriority[is.na(by_type$TypePriority)] <- 99L
  by_type <- by_type[order(by_type$Year, by_type$TypePriority), , drop = FALSE]
  out <- by_type[!duplicated(by_type$Year), c("Year", "GNPI"), drop = FALSE]
  rownames(out) <- NULL
  out
}

xl_pricing_tool_cat_onlevel_read_cpi <- function(sheet) {
  rows <- list()
  blank_run <- 0L
  for (row in 16:(16 + 5000)) {
    year <- xl_pricing_tool_cat_onlevel_number(sheet, "E", row, "CPI Year", required = FALSE)
    current <- xl_pricing_tool_cat_onlevel_number(sheet, "G", row, "Inflat. Factor_Current", required = FALSE)
    prior <- xl_pricing_tool_cat_onlevel_number(sheet, "H", row, "Inflat. Factor_Prior", required = FALSE)
    if (!is.finite(year) && !is.finite(current) && !is.finite(prior)) {
      blank_run <- blank_run + 1L
      if (blank_run >= 20L) break
      next
    }
    blank_run <- 0L
    if (is.finite(year)) {
      rows[[length(rows) + 1L]] <- data.frame(
        Year = as.integer(year),
        Current = current,
        Prior = prior,
        stringsAsFactors = FALSE
      )
    }
  }
  if (length(rows) == 0) {
    stop("CPI source has no usable CPI rows.", call. = FALSE)
  }
  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

xl_pricing_tool_cat_onlevel_read_keyzone <- function(sheet) {
  rows <- list()
  blank_run <- 0L
  for (row in 16:(16 + 5000)) {
    province <- trimws(xl_pricing_tool_cell_value(sheet, paste0("K", row)))
    peril <- toupper(trimws(xl_pricing_tool_cell_value(sheet, paste0("L", row))))
    blended <- xl_pricing_tool_cat_onlevel_number(sheet, "O", row, "Blended", required = FALSE)
    if (!nzchar(province) && !nzchar(peril)) {
      blank_run <- blank_run + 1L
      if (blank_run >= 20L) break
      next
    }
    blank_run <- 0L
    if (nzchar(province) && nzchar(peril) && is.finite(blended) && blended != 0) {
      rows[[length(rows) + 1L]] <- data.frame(
        Province = province,
        ProvinceKey = xl_pricing_tool_cat_onlevel_key(province),
        Peril = peril,
        stringsAsFactors = FALSE
      )
    }
  }
  if (length(rows) == 0) {
    data.frame(Province = character(), ProvinceKey = character(), Peril = character(), stringsAsFactors = FALSE)
  } else {
    out <- do.call(rbind, rows)
    rownames(out) <- NULL
    out
  }
}

xl_pricing_tool_cat_onlevel_calculate <- function(events, aggregate, gnpi, cpi, keyzone) {
  rows <- lapply(seq_len(nrow(events)), function(i) {
    event <- events[i, , drop = FALSE]
    result <- xl_pricing_tool_cat_onlevel_calculate_row(event, aggregate, gnpi, cpi, keyzone)
    data.frame(
      from_Expo = result$from,
      to_Expo = result$to,
      warning = paste(result$warnings, collapse = "; "),
      stringsAsFactors = FALSE
    )
  })
  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

xl_pricing_tool_cat_onlevel_calculate_row <- function(event, aggregate, gnpi, cpi, keyzone) {
  method <- xl_pricing_tool_cat_onlevel_method(event$Method[[1]])
  warnings <- character()

  if (!is.finite(event$In_Prior_Current[[1]]) || event$In_Prior_Current[[1]] == 0) {
    return(list(from = NA_real_, to = NA_real_, warnings = character()))
  }

  if (method %in% c("no onlevel", "user defined")) {
    warning <- if (method == "user defined") "User-defined method treated as No On-level in this version" else character()
    return(list(from = 1, to = 1, warnings = warning))
  }

  if (method == "gnpi") {
    from_lookup <- xl_pricing_tool_cat_onlevel_lookup_year(gnpi, "GNPI", event$from_year[[1]])
    to_lookup <- xl_pricing_tool_cat_onlevel_lookup_year(gnpi, "GNPI", event$to_year[[1]])
    warnings <- c(warnings, from_lookup$warning, to_lookup$warning)
    return(list(from = from_lookup$value, to = to_lookup$value, warnings = warnings[nzchar(warnings)]))
  }

  if (method == "cpi") {
    as_at <- xl_pricing_tool_cat_onlevel_asat(event$AsAt[[1]])
    column <- if (as_at == "Current") "Current" else if (as_at == "Prior") "Prior" else NA_character_
    if (is.na(column)) {
      return(list(from = 1, to = NA_real_, warnings = sprintf("Unsupported AsAt for CPI: %s", event$AsAt[[1]])))
    }
    lookup <- xl_pricing_tool_cat_onlevel_lookup_year(cpi, column, event$from_year[[1]])
    warnings <- c(warnings, lookup$warning)
    return(list(from = 1, to = lookup$value, warnings = warnings[nzchar(warnings)]))
  }

  if (!method %in% c("province", "nationwide", "keyzone")) {
    return(list(from = NA_real_, to = NA_real_, warnings = sprintf("Unsupported on-level method: %s", event$Method[[1]])))
  }

  provinces <- switch(
    method,
    province = xl_pricing_tool_cat_onlevel_split_provinces(event$Affected_Provinces[[1]]),
    nationwide = unique(aggregate$ProvinceKey[aggregate$Peril == event$Peril[[1]]]),
    keyzone = unique(keyzone$ProvinceKey[keyzone$Peril == event$Peril[[1]]]),
    character()
  )

  if (length(provinces) == 0) {
    return(list(from = NA_real_, to = NA_real_, warnings = sprintf("No provinces resolved for method %s", event$Method[[1]])))
  }

  known <- unique(aggregate$ProvinceKey[aggregate$Peril == event$Peril[[1]]])
  unrecognized <- setdiff(provinces, known)
  if (length(unrecognized) > 0) {
    warnings <- c(warnings, sprintf("Unrecognized province(s): %s", paste(unrecognized, collapse = ", ")))
  }
  provinces <- intersect(provinces, known)
  if (length(provinces) == 0) {
    return(list(from = NA_real_, to = NA_real_, warnings = warnings))
  }

  from_expo <- xl_pricing_tool_cat_onlevel_sum_exposure(aggregate, event$Peril[[1]], provinces, event$from_year[[1]])
  to_expo <- xl_pricing_tool_cat_onlevel_sum_exposure(aggregate, event$Peril[[1]], provinces, event$to_year[[1]])
  warnings <- c(warnings, from_expo$warning, to_expo$warning)
  list(from = from_expo$value, to = to_expo$value, warnings = warnings[nzchar(warnings)])
}

xl_pricing_tool_cat_onlevel_sum_exposure <- function(aggregate, peril, province_keys, year) {
  available_years <- sort(unique(aggregate$Year[aggregate$Peril == peril & aggregate$ProvinceKey %in% province_keys]))
  resolved <- xl_pricing_tool_cat_onlevel_resolve_year(available_years, year)
  value <- sum(aggregate$Exposure[
    aggregate$Peril == peril &
      aggregate$ProvinceKey %in% province_keys &
      aggregate$Year == resolved$year
  ], na.rm = TRUE)
  if (!is.finite(value) || value <= 0) {
    resolved$warning <- paste(c(resolved$warning, sprintf("No positive aggregate exposure for %s year %s", peril, resolved$year)), collapse = "; ")
  }
  list(value = value, warning = resolved$warning)
}

xl_pricing_tool_cat_onlevel_lookup_year <- function(data, value_col, requested_year) {
  available_years <- sort(unique(data$Year[is.finite(data[[value_col]])]))
  resolved <- xl_pricing_tool_cat_onlevel_resolve_year(available_years, requested_year)
  value <- data[[value_col]][data$Year == resolved$year][[1]]
  list(value = value, warning = resolved$warning)
}

xl_pricing_tool_cat_onlevel_resolve_year <- function(available_years, requested_year) {
  if (length(available_years) == 0) {
    return(list(year = NA_integer_, warning = "No available years in source table"))
  }
  if (requested_year %in% available_years) {
    return(list(year = requested_year, warning = ""))
  }
  if (requested_year < min(available_years)) {
    year <- min(available_years)
  } else if (requested_year > max(available_years)) {
    year <- max(available_years)
  } else {
    year <- max(available_years[available_years < requested_year])
  }
  list(year = year, warning = sprintf("Year %s not available; used %s", requested_year, year))
}

xl_pricing_tool_cat_onlevel_split_provinces <- function(value) {
  value <- trimws(value)
  if (!nzchar(value) || tolower(value) == "all") {
    return(character())
  }
  parts <- unlist(strsplit(value, "[,;/]+"))
  unique(xl_pricing_tool_cat_onlevel_key(parts[nzchar(trimws(parts))]))
}

xl_pricing_tool_cat_onlevel_method <- function(value) {
  key <- tolower(trimws(value))
  key <- gsub("[-_]+", " ", key)
  key <- gsub("\\s+", " ", key)
  if (key %in% c("no on level", "no onlevel", "no on-level", "none")) return("no onlevel")
  if (key %in% c("user defined", "user-defined")) return("user defined")
  key
}

xl_pricing_tool_cat_onlevel_asat <- function(value) {
  key <- tolower(trimws(value))
  if (key %in% c("prior", "p")) return("Prior")
  if (key %in% c("current", "curr", "c")) return("Current")
  value
}

xl_pricing_tool_cat_onlevel_key <- function(value) {
  key <- trimws(value)
  key <- gsub("\\s+", " ", key)
  key
}

xl_pricing_tool_cat_onlevel_number <- function(sheet, col, row, field_name, required = TRUE) {
  raw_value <- trimws(xl_pricing_tool_cell_value(sheet, paste0(col, row)))
  if (!nzchar(raw_value) || raw_value %in% c("#N/A", "#VALUE!", "#DIV/0!")) {
    if (required) {
      stop(sprintf("CatOnLevel %s cell %s%s is blank or invalid.", field_name, col, row), call. = FALSE)
    }
    return(NA_real_)
  }
  value <- suppressWarnings(as.numeric(gsub(",", "", raw_value, fixed = TRUE)))
  if (is.na(value)) {
    if (required) {
      stop(sprintf("CatOnLevel %s cell %s%s is not numeric: %s", field_name, col, row, raw_value), call. = FALSE)
    }
    return(NA_real_)
  }
  value
}
