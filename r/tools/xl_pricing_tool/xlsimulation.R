xl_pricing_tool_xlsimulation_gather <- function(workbook_path, output_dir = NULL, context = list()) {
  xl_pricing_tool_validate(workbook_path, context)

  sim <- xl_pricing_tool_xlsimulation_gather_context(workbook_path, output_dir, context)
  inputs <- xl_pricing_tool_xlsimulation_read_gather_contract(sim)

  coverage_handoff <- xl_pricing_tool_xlsimulation_coverage_handoff(inputs$layers, inputs$active_causes, inputs$coverage)
  breakdown_handoff <- xl_pricing_tool_xlsimulation_breakdown_empty_handoff(inputs$layers, inputs$active_causes)

  utils::write.table(
    coverage_handoff,
    file = file.path(sim$sim_dir, "xlsimulation_coverage_matrix.csv"),
    sep = ",",
    row.names = FALSE,
    col.names = FALSE,
    na = "",
    quote = TRUE
  )
  utils::write.table(
    breakdown_handoff,
    file = file.path(sim$sim_dir, "xlsimulation_layer_breakdown.csv"),
    sep = ",",
    row.names = FALSE,
    col.names = FALSE,
    na = "",
    quote = TRUE
  )

  summary <- data.frame(
    Key = c(
      "Active sheet",
      "Active layers",
      "Active causes",
      "Inactive causes skipped",
      "Coverage matrix output",
      "Breakdown header output"
    ),
    Value = c(
      sim$sheet_name,
      paste(inputs$layers$LayerID, collapse = ", "),
      paste(inputs$active_causes$CauseID, collapse = ", "),
      paste(inputs$inactive_causes$CauseID, collapse = ", "),
      file.path(sim$sim_dir, "xlsimulation_coverage_matrix.csv"),
      file.path(sim$sim_dir, "xlsimulation_layer_breakdown.csv")
    ),
    stringsAsFactors = FALSE
  )
  utils::write.csv(file = file.path(sim$sim_dir, "xlsimulation_gather_summary.csv"), summary, row.names = FALSE, na = "")

  paste(
    "XLSimulation gather completed.",
    sprintf("Output folder: %s", sim$output_dir),
    sprintf("Active sheet: %s", sim$sheet_name),
    sprintf("Active layers: %s", nrow(inputs$layers)),
    sprintf("Active causes: %s", paste(inputs$active_causes$CauseID, collapse = ", ")),
    sprintf("Inactive causes skipped: %s", paste(inputs$inactive_causes$CauseID, collapse = ", ")),
    sprintf("Coverage matrix output: %s", file.path(sim$sim_dir, "xlsimulation_coverage_matrix.csv")),
    sprintf("Breakdown header output: %s", file.path(sim$sim_dir, "xlsimulation_layer_breakdown.csv")),
    "Source loss blocks are not read during gather.",
    "CSV outputs are handled by the VBA client after BERT returns.",
    sep = vb_newline()
  )
}

xl_pricing_tool_xlsimulation_update <- function(workbook_path, output_dir = NULL, context = list()) {
  xl_pricing_tool_validate(workbook_path, context)

  sim <- xl_pricing_tool_xlsimulation_update_context(workbook_path, output_dir, context)
  inputs <- xl_pricing_tool_xlsimulation_read_contract(sim, validate_sources = TRUE)

  if (!is.na(inputs$settings$RandomSeed)) {
    set.seed(inputs$settings$RandomSeed)
  }

  cause_events <- list()
  cause_oep <- list()
  for (i in seq_len(nrow(inputs$active_causes))) {
    cause <- inputs$active_causes[i, , drop = FALSE]
    source <- inputs$sources[[cause$CauseID]]
    events <- xl_pricing_tool_xlsimulation_simulate_cause(cause, source, inputs$settings)
    cause_events[[cause$CauseID]] <- events
    cause_oep[[cause$CauseID]] <- xl_pricing_tool_xlsimulation_annual_oep(events, inputs$settings$SimulationYears)
  }

  layer_results <- xl_pricing_tool_xlsimulation_layer_results(
    layers = inputs$layers,
    active_causes = inputs$active_causes,
    coverage = inputs$coverage,
    cause_events = cause_events,
    settings = inputs$settings
  )

  oep_handoff <- xl_pricing_tool_xlsimulation_oep_handoff(
    causes = inputs$causes,
    cause_oep = cause_oep,
    return_periods = xl_pricing_tool_xlsimulation_oep_return_periods(sim)
  )

  utils::write.table(
    layer_results$layer_output,
    file = file.path(sim$sim_dir, "xlsimulation_layer_output.csv"),
    sep = ",",
    row.names = FALSE,
    col.names = FALSE,
    na = "",
    quote = TRUE
  )
  utils::write.table(
    layer_results$layer_breakdown,
    file = file.path(sim$sim_dir, "xlsimulation_layer_breakdown.csv"),
    sep = ",",
    row.names = FALSE,
    col.names = FALSE,
    na = "",
    quote = TRUE
  )
  utils::write.table(
    oep_handoff,
    file = file.path(sim$sim_dir, "xlsimulation_oep_output.csv"),
    sep = ",",
    row.names = FALSE,
    col.names = FALSE,
    na = "",
    quote = TRUE
  )

  paste(
    "XLSimulation update completed.",
    sprintf("Output folder: %s", sim$output_dir),
    sprintf("Active sheet: %s", sim$sheet_name),
    sprintf("Simulation years: %s", inputs$settings$SimulationYears),
    sprintf("Active causes: %s", paste(inputs$active_causes$CauseID, collapse = ", ")),
    sprintf("Layer output: %s", file.path(sim$sim_dir, "xlsimulation_layer_output.csv")),
    sprintf("Layer breakdown: %s", file.path(sim$sim_dir, "xlsimulation_layer_breakdown.csv")),
    sprintf("OEP output: %s", file.path(sim$sim_dir, "xlsimulation_oep_output.csv")),
    "CSV outputs are handled by the VBA client after BERT returns.",
    sep = vb_newline()
  )
}

xl_pricing_tool_xlsimulation_build <- function(workbook_path, output_dir = NULL, context = list()) {
  "No build action defined for XLSimulation. V1 build is reserved for a future pricing handoff/export step."
}

xl_pricing_tool_xlsimulation_context <- function(workbook_path, output_dir = NULL, context = list()) {
  if (is.null(output_dir) || identical(output_dir, "")) {
    output_dir <- if (!is.null(context$output_dir)) {
      context$output_dir
    } else {
      btk_default_run_dir("xl_pricing_tool", workbook_path)
    }
  }

  sheet_name <- "Sim_Variations"
  if (!is.null(context$active_sheet) && nzchar(context$active_sheet)) {
    sheet_name <- context$active_sheet
  }

  sim_dir <- file.path(output_dir, "xlsimulation")
  dir.create(sim_dir, recursive = TRUE, showWarnings = FALSE)

  list(
    workbook_path = workbook_path,
    output_dir = output_dir,
    sim_dir = sim_dir,
    sheet_name = sheet_name,
    cells = xl_pricing_tool_xlsimulation_load_cells(workbook_path, sheet_name)
  )
}

xl_pricing_tool_xlsimulation_gather_context <- function(workbook_path, output_dir = NULL, context = list()) {
  if (is.null(output_dir) || identical(output_dir, "")) {
    output_dir <- if (!is.null(context$output_dir)) {
      context$output_dir
    } else {
      btk_default_run_dir("xl_pricing_tool", workbook_path)
    }
  }

  sheet_name <- "Sim_Variations"
  if (!is.null(context$active_sheet) && nzchar(context$active_sheet)) {
    sheet_name <- context$active_sheet
  }

  sim_dir <- file.path(output_dir, "xlsimulation")
  dir.create(sim_dir, recursive = TRUE, showWarnings = FALSE)

  list(
    workbook_path = workbook_path,
    output_dir = output_dir,
    sim_dir = sim_dir,
    sheet_name = sheet_name,
    cells = xl_pricing_tool_xlsimulation_load_cells_for_ranges(
      workbook_path = workbook_path,
      sheet_name = sheet_name,
      ranges = c("B26:H41", "B49:J61", "M49:V61")
    )
  )
}

xl_pricing_tool_xlsimulation_update_context <- function(workbook_path, output_dir = NULL, context = list()) {
  sim <- xl_pricing_tool_xlsimulation_base_context(workbook_path, output_dir, context)
  config_ranges <- c("B13:D18", "B26:H41", "B49:J61", "M49:V61", "B109:B121")
  sim$cells <- xl_pricing_tool_xlsimulation_load_cells_for_ranges(
    workbook_path = workbook_path,
    sheet_name = sim$sheet_name,
    ranges = config_ranges
  )

  causes <- xl_pricing_tool_xlsimulation_read_causes(sim)
  active_causes <- causes[toupper(causes$Active) == "Y", , drop = FALSE]
  if (nrow(active_causes) == 0) {
    stop("XLSimulation has no active loss causes in B49:J61.", call. = FALSE)
  }
  xl_pricing_tool_xlsimulation_validate_cause_declarations(active_causes)

  sim$source_manifest <- xl_pricing_tool_xlsimulation_load_source_manifest(sim)
  manifest_cause_ids <- if (is.null(sim$source_manifest)) character() else sim$source_manifest$CauseID
  causes_needing_workbook <- active_causes[!active_causes$CauseID %in% manifest_cause_ids, , drop = FALSE]
  source_ranges <- unique(trimws(causes_needing_workbook$Location[nzchar(trimws(causes_needing_workbook$Location))]))
  if (length(source_ranges) > 0) {
    source_cells <- xl_pricing_tool_xlsimulation_load_source_cells_for_ranges(
      workbook_path = workbook_path,
      sheet_name = sim$sheet_name,
      ranges = source_ranges
    )
    sim$cells <- xl_pricing_tool_xlsimulation_merge_cells(sim$cells, source_cells)
  }

  sim
}

xl_pricing_tool_xlsimulation_base_context <- function(workbook_path, output_dir = NULL, context = list()) {
  if (is.null(output_dir) || identical(output_dir, "")) {
    output_dir <- if (!is.null(context$output_dir)) {
      context$output_dir
    } else {
      btk_default_run_dir("xl_pricing_tool", workbook_path)
    }
  }

  sheet_name <- "Sim_Variations"
  if (!is.null(context$active_sheet) && nzchar(context$active_sheet)) {
    sheet_name <- context$active_sheet
  }

  sim_dir <- file.path(output_dir, "xlsimulation")
  dir.create(sim_dir, recursive = TRUE, showWarnings = FALSE)

  list(
    workbook_path = workbook_path,
    output_dir = output_dir,
    sim_dir = sim_dir,
    sheet_name = sheet_name
  )
}

xl_pricing_tool_xlsimulation_read_gather_contract <- function(sim) {
  layers <- xl_pricing_tool_xlsimulation_read_layers(sim)
  causes <- xl_pricing_tool_xlsimulation_read_causes(sim)
  active_causes <- causes[toupper(causes$Active) == "Y", , drop = FALSE]
  inactive_causes <- causes[toupper(causes$Active) != "Y", , drop = FALSE]

  if (nrow(active_causes) == 0) {
    stop("XLSimulation has no active loss causes in B49:J61.", call. = FALSE)
  }

  missing_family <- active_causes$CauseID[!nzchar(active_causes$CauseFamily)]
  if (length(missing_family) > 0) {
    stop(sprintf("XLSimulation active cause(s) are missing CauseFamily: %s", paste(missing_family, collapse = ", ")), call. = FALSE)
  }
  xl_pricing_tool_xlsimulation_validate_cause_declarations(active_causes)

  coverage <- xl_pricing_tool_xlsimulation_read_coverage(sim, layers, active_causes)

  list(
    layers = layers,
    causes = causes,
    active_causes = active_causes,
    inactive_causes = inactive_causes,
    coverage = coverage
  )
}

xl_pricing_tool_xlsimulation_read_contract <- function(sim, validate_sources = TRUE) {
  settings <- xl_pricing_tool_xlsimulation_read_settings(sim)
  layers <- xl_pricing_tool_xlsimulation_read_layers(sim)
  causes <- xl_pricing_tool_xlsimulation_read_causes(sim)
  active_causes <- causes[toupper(causes$Active) == "Y", , drop = FALSE]
  inactive_causes <- causes[toupper(causes$Active) != "Y", , drop = FALSE]
  coverage <- xl_pricing_tool_xlsimulation_read_coverage(sim, layers, active_causes)

  if (nrow(active_causes) == 0) {
    stop("XLSimulation has no active loss causes in B49:J61.", call. = FALSE)
  }
  xl_pricing_tool_xlsimulation_validate_cause_declarations(active_causes)

  sources <- list()
  source_summaries <- character()
  if (isTRUE(validate_sources)) {
    for (i in seq_len(nrow(active_causes))) {
      cause <- active_causes[i, , drop = FALSE]
      source <- xl_pricing_tool_xlsimulation_read_source(sim, cause)
      sources[[cause$CauseID]] <- source
      source_summaries <- c(source_summaries, sprintf("%s: %s rows", cause$CauseID, nrow(source$data)))
    }
  }

  list(
    settings = settings,
    layers = layers,
    causes = causes,
    active_causes = active_causes,
    inactive_causes = inactive_causes,
    coverage = coverage,
    sources = sources,
    source_summaries = source_summaries
  )
}

xl_pricing_tool_xlsimulation_read_settings <- function(sim) {
  raw <- xl_pricing_tool_xlsimulation_read_range(sim, "B13:D18", col_names = FALSE)
  if (nrow(raw) < 2 || ncol(raw) < 2) {
    stop("XLSimulation settings table B13:D18 is missing.", call. = FALSE)
  }

  keys <- trimws(as.character(raw[-1, 1]))
  values <- raw[-1, 2]
  names(values) <- xl_pricing_tool_xlsimulation_key(keys)

  years <- xl_pricing_tool_xlsimulation_number(values[["simulationyears"]], "Simulation Years")
  seed <- suppressWarnings(as.integer(values[["randomseed"]]))
  return_period <- xl_pricing_tool_xlsimulation_number(values[["returnperiod"]], "Return Period")
  risk_measure <- toupper(trimws(as.character(values[["riskmeasure"]])))
  scaling <- xl_pricing_tool_xlsimulation_number(values[["catscalingfactor"]], "CAT Scaling Factor")

  if (!risk_measure %in% c("VAR", "TVAR")) {
    stop("Risk Measure must be VaR or TVaR.", call. = FALSE)
  }
  if (years <= 0 || years != floor(years)) {
    stop("Simulation Years must be a positive integer.", call. = FALSE)
  }
  if (return_period <= 1) {
    stop("Return Period must be greater than 1.", call. = FALSE)
  }

  list(
    SimulationYears = as.integer(years),
    RandomSeed = seed,
    ReturnPeriod = return_period,
    RiskMeasure = risk_measure,
    CatScalingFactor = scaling
  )
}

xl_pricing_tool_xlsimulation_read_layers <- function(sim) {
  raw <- xl_pricing_tool_xlsimulation_read_range(sim, "B26:H41", col_names = FALSE)
  expected <- c("LayerID", "LayerName", "Limit", "Deductible", "# of Reinstatements", "ReinstatementPct", "InputROL")
  headers <- trimws(as.character(raw[1, ]))
  if (!all(headers == expected)) {
    stop("XLSimulation layer table must have expected headers in B26:H26.", call. = FALSE)
  }

  raw <- raw[-1, , drop = FALSE]
  names(raw) <- expected
  raw <- raw[nzchar(trimws(as.character(raw$LayerID))), , drop = FALSE]
  if (nrow(raw) == 0) {
    stop("XLSimulation layer table has no active layer rows.", call. = FALSE)
  }

  out <- data.frame(
    LayerID = as.character(raw$LayerID),
    LayerName = as.character(raw$LayerName),
    Limit = xl_pricing_tool_xlsimulation_numeric_vector(raw$Limit, "Limit"),
    Deductible = xl_pricing_tool_xlsimulation_numeric_vector(raw$Deductible, "Deductible"),
    Reinstatements = xl_pricing_tool_xlsimulation_numeric_vector(raw[["# of Reinstatements"]], "# of Reinstatements"),
    ReinstatementPct = xl_pricing_tool_xlsimulation_numeric_vector(raw$ReinstatementPct, "ReinstatementPct"),
    InputROL = xl_pricing_tool_xlsimulation_numeric_vector(raw$InputROL, "InputROL"),
    stringsAsFactors = FALSE
  )

  if (any(out$Limit < 0) || any(out$Deductible < 0) || any(out$Reinstatements < 0) || any(out$ReinstatementPct < 0)) {
    stop("XLSimulation layer values cannot be negative.", call. = FALSE)
  }

  out
}

xl_pricing_tool_xlsimulation_read_causes <- function(sim) {
  raw <- xl_pricing_tool_xlsimulation_read_range(sim, "B49:J61", col_names = FALSE)
  expected <- c("CauseID", "CauseFamily", "ModelSource", "Peril", "InputMode", "FilePath", "Active", "Notes", "Location")
  headers <- trimws(as.character(raw[1, ]))
  if (!all(headers == expected)) {
    stop("XLSimulation loss cause declarations must have expected headers in B49:J49.", call. = FALSE)
  }

  raw <- raw[-1, , drop = FALSE]
  names(raw) <- expected
  raw <- raw[nzchar(trimws(as.character(raw$CauseID))), , drop = FALSE]
  raw[] <- lapply(raw, function(x) trimws(as.character(x)))
  rownames(raw) <- NULL
  raw
}

xl_pricing_tool_xlsimulation_read_coverage <- function(sim, layers, active_causes) {
  raw <- xl_pricing_tool_xlsimulation_read_range(sim, "M49:V61", col_names = FALSE)
  if (nrow(raw) == 0 || ncol(raw) == 0) {
    stop("XLSimulation coverage matrix M49:V61 is missing.", call. = FALSE)
  }

  headers <- trimws(as.character(raw[1, ]))
  existing_causes <- headers[-1][nzchar(headers[-1])]
  rows <- raw[-1, , drop = FALSE]
  layer_ids <- trimws(as.character(rows[[1]]))
  rows <- rows[nzchar(layer_ids), , drop = FALSE]
  layer_ids <- layer_ids[nzchar(layer_ids)]

  out <- matrix(FALSE, nrow = nrow(layers), ncol = nrow(active_causes))
  rownames(out) <- layers$LayerID
  colnames(out) <- active_causes$CauseID

  for (i in seq_len(nrow(layers))) {
    row_index <- match(layers$LayerID[[i]], layer_ids)
    for (j in seq_len(nrow(active_causes))) {
      col_index <- match(active_causes$CauseID[[j]], existing_causes)
      if (!is.na(row_index) && !is.na(col_index)) {
        value <- toupper(trimws(as.character(rows[row_index, col_index + 1])))
        out[i, j] <- value %in% c("Y", "YES", "1", "TRUE")
      } else {
        out[i, j] <- TRUE
      }
    }
  }

  out
}

xl_pricing_tool_xlsimulation_read_source <- function(sim, cause) {
  manifest_source <- xl_pricing_tool_xlsimulation_read_source_from_manifest(sim, cause)
  if (!is.null(manifest_source)) {
    return(manifest_source)
  }

  parsed <- xl_pricing_tool_xlsimulation_parse_location(cause$Location)
  if (is.null(parsed)) {
    stop(sprintf("XLSimulation cause %s has malformed Location: %s", cause$CauseID, cause$Location), call. = FALSE)
  }

  raw <- xl_pricing_tool_xlsimulation_read_range(sim, parsed$range, col_names = FALSE)
  if (nrow(raw) < 2) {
    stop(sprintf("XLSimulation cause %s source block has no data rows.", cause$CauseID), call. = FALSE)
  }

  headers <- trimws(as.character(raw[1, ]))
  data <- raw[-1, , drop = FALSE]
  names(data) <- make.names(headers, unique = TRUE)
  data <- xl_pricing_tool_xlsimulation_drop_blank_rows(data)

  xl_pricing_tool_xlsimulation_validate_source(cause, headers, data)
  list(headers = headers, data = data, range = parsed$range)
}

xl_pricing_tool_xlsimulation_load_source_manifest <- function(sim) {
  manifest_path <- file.path(sim$sim_dir, "source_inputs", "xlsimulation_source_manifest.csv")
  if (!file.exists(manifest_path)) {
    return(NULL)
  }

  manifest <- utils::read.csv(
    manifest_path,
    stringsAsFactors = FALSE,
    check.names = FALSE,
    colClasses = "character",
    na.strings = character()
  )
  required <- c("CauseID", "CauseFamily", "ModelSource", "Location", "CsvPath")
  missing <- setdiff(required, names(manifest))
  if (length(missing) > 0) {
    stop(sprintf("XLSimulation source manifest is missing: %s.", paste(missing, collapse = ", ")), call. = FALSE)
  }

  manifest[] <- lapply(manifest, function(x) trimws(as.character(x)))
  manifest <- manifest[nzchar(manifest$CauseID) & nzchar(manifest$CsvPath), , drop = FALSE]
  rownames(manifest) <- NULL
  manifest
}

xl_pricing_tool_xlsimulation_read_source_from_manifest <- function(sim, cause) {
  manifest <- sim$source_manifest
  if (is.null(manifest) || nrow(manifest) == 0) {
    return(NULL)
  }

  row_index <- match(cause$CauseID, manifest$CauseID)
  if (is.na(row_index)) {
    return(NULL)
  }

  csv_path <- manifest$CsvPath[[row_index]]
  if (!file.exists(csv_path)) {
    stop(sprintf("XLSimulation source CSV for %s was not found: %s", cause$CauseID, csv_path), call. = FALSE)
  }

  raw <- utils::read.csv(
    csv_path,
    header = FALSE,
    stringsAsFactors = FALSE,
    check.names = FALSE,
    colClasses = "character",
    na.strings = character(),
    blank.lines.skip = FALSE
  )
  raw <- xl_pricing_tool_xlsimulation_trim_empty_edges(raw)
  if (nrow(raw) < 2) {
    stop(sprintf("XLSimulation cause %s source CSV has no data rows.", cause$CauseID), call. = FALSE)
  }

  headers <- trimws(as.character(raw[1, ]))
  data <- raw[-1, , drop = FALSE]
  names(data) <- make.names(headers, unique = TRUE)
  data <- xl_pricing_tool_xlsimulation_drop_blank_rows(data)

  xl_pricing_tool_xlsimulation_validate_source(cause, headers, data)
  list(headers = headers, data = data, range = manifest$Location[[row_index]], csv_path = csv_path)
}

xl_pricing_tool_xlsimulation_validate_source <- function(cause, headers, data) {
  family <- xl_pricing_tool_xlsimulation_family(cause$CauseFamily)
  model <- toupper(trimws(cause$ModelSource))

  if (identical(family, "ModeledCAT") && model %in% c("CTR", "TS")) {
    expected <- c("EventID", "Year", "Loss")
    if (!all(headers == expected)) {
      stop(sprintf("%s must have headers: %s.", cause$CauseID, paste(expected, collapse = ", ")), call. = FALSE)
    }
    if (nrow(data) == 0) stop(sprintf("%s has no event rows.", cause$CauseID), call. = FALSE)
  } else if (identical(family, "ModeledCAT") && identical(model, "RMS")) {
    normalized <- gsub("^Evend\\.ID$", "Event.ID", make.names(headers))
    expected <- c("Event.ID", "Event.Rate", "Mean.Loss", "Std.Dev.Ind", "Std.Dev.Corr", "Exposure")
    if (!all(normalized == expected)) {
      stop(sprintf("%s must have RMS headers: Event ID, Event Rate, Mean Loss, Std Dev Ind, Std Dev Corr, Exposure.", cause$CauseID), call. = FALSE)
    }
    if (nrow(data) == 0) stop(sprintf("%s has no RMS event rows.", cause$CauseID), call. = FALSE)
  } else if (identical(family, "FS")) {
    keys <- trimws(as.character(data[[1]]))
    required <- c("Mean Frequency", "Frequency Type", "Severity Family", "Loss Cap", "Minimum Loss")
    missing <- setdiff(required, keys)
    if (length(missing) > 0) {
      stop(sprintf("%s FS table is missing: %s.", cause$CauseID, paste(missing, collapse = ", ")), call. = FALSE)
    }
  } else if (identical(family, "CDF")) {
    keys <- trimws(as.character(data[[1]]))
    required <- c("Mean Frequency", "Frequency Type", "Interpolation")
    missing <- setdiff(required, keys)
    if (length(missing) > 0) {
      stop(sprintf("%s CDF table is missing: %s.", cause$CauseID, paste(missing, collapse = ", ")), call. = FALSE)
    }
    if (!"Percentile" %in% keys || !"Loss Severity" %in% trimws(as.character(data[[2]]))) {
      stop(sprintf("%s CDF table must include Percentile/Loss Severity rows.", cause$CauseID), call. = FALSE)
    }
  } else {
    stop(sprintf("Unsupported XLSimulation cause family/model for %s: %s/%s", cause$CauseID, family, model), call. = FALSE)
  }

  invisible(TRUE)
}

xl_pricing_tool_xlsimulation_simulate_cause <- function(cause, source, settings) {
  family <- xl_pricing_tool_xlsimulation_family(cause$CauseFamily)
  model <- toupper(trimws(cause$ModelSource))

  if (identical(family, "ModeledCAT") && model %in% c("CTR", "TS")) {
    events <- xl_pricing_tool_xlsimulation_simulate_elt(source$data, settings$SimulationYears)
  } else if (identical(family, "ModeledCAT") && identical(model, "RMS")) {
    events <- xl_pricing_tool_xlsimulation_simulate_rms(source$data, settings$SimulationYears)
  } else if (identical(family, "FS")) {
    events <- xl_pricing_tool_xlsimulation_simulate_fs(source$data, settings$SimulationYears)
  } else if (identical(family, "CDF")) {
    events <- xl_pricing_tool_xlsimulation_simulate_cdf(source$data, settings$SimulationYears)
  } else {
    stop(sprintf("Unsupported XLSimulation cause: %s", cause$CauseID), call. = FALSE)
  }

  events$CauseID <- cause$CauseID
  if (identical(family, "ModeledCAT")) {
    events$Loss <- events$Loss * settings$CatScalingFactor
  }
  events[, c("Year", "CauseID", "EventID", "Loss"), drop = FALSE]
}

xl_pricing_tool_xlsimulation_simulate_elt <- function(data, simulation_years) {
  names(data) <- c("EventID", "Year", "Loss")
  data$Year <- suppressWarnings(as.integer(data$Year))
  data$Loss <- suppressWarnings(as.numeric(data$Loss))
  data <- data[is.finite(data$Year) & is.finite(data$Loss) & data$Loss > 0, , drop = FALSE]
  if (nrow(data) == 0) {
    return(xl_pricing_tool_xlsimulation_empty_events())
  }

  data$EventID <- as.character(data$EventID)
  source_years <- unique(data$Year)
  source_year_count <- length(source_years)
  source_year_map <- stats::setNames(seq_along(source_years), as.character(source_years))

  if (simulation_years <= source_year_count) {
    selected_years <- source_years[seq_len(simulation_years)]
    keep <- data$Year %in% selected_years
    if (!any(keep)) {
      return(xl_pricing_tool_xlsimulation_empty_events())
    }
    return(data.frame(
      Year = unname(source_year_map[as.character(data$Year[keep])]),
      EventID = data$EventID[keep],
      Loss = data$Loss[keep],
      stringsAsFactors = FALSE
    ))
  }

  full_cycles <- simulation_years %/% source_year_count
  remainder <- simulation_years %% source_year_count
  rows <- vector("list", full_cycles + as.integer(remainder > 0L))
  for (cycle in seq_len(full_cycles)) {
    rows[[cycle]] <- data.frame(
      Year = unname(source_year_map[as.character(data$Year)]) + (cycle - 1L) * source_year_count,
      EventID = data$EventID,
      Loss = data$Loss,
      stringsAsFactors = FALSE
    )
  }
  if (remainder > 0L) {
    selected_years <- source_years[seq_len(remainder)]
    keep <- data$Year %in% selected_years
    rows[[length(rows)]] <- data.frame(
      Year = unname(source_year_map[as.character(data$Year[keep])]) + full_cycles * source_year_count,
      EventID = data$EventID[keep],
      Loss = data$Loss[keep],
      stringsAsFactors = FALSE
    )
  }

  out <- do.call(rbind, rows)
  if (is.null(out)) xl_pricing_tool_xlsimulation_empty_events() else out
}

xl_pricing_tool_xlsimulation_simulate_rms <- function(data, simulation_years) {
  names(data) <- c("EventID", "EventRate", "MeanLoss", "StdDevInd", "StdDevCorr", "Exposure")
  data$EventRate <- suppressWarnings(as.numeric(data$EventRate))
  data$MeanLoss <- suppressWarnings(as.numeric(data$MeanLoss))
  data$StdDevInd <- suppressWarnings(as.numeric(data$StdDevInd))
  data <- data[is.finite(data$EventRate) & data$EventRate > 0 & is.finite(data$MeanLoss) & data$MeanLoss > 0, , drop = FALSE]
  if (nrow(data) == 0) return(xl_pricing_tool_xlsimulation_empty_events())

  lambda <- sum(data$EventRate, na.rm = TRUE)
  counts <- stats::rpois(simulation_years, lambda = lambda)
  total <- sum(counts)
  if (total == 0) return(xl_pricing_tool_xlsimulation_empty_events())

  selected <- xl_pricing_tool_xlsimulation_sample_weighted_cdf(data$EventRate, total)
  selected_data <- data[selected, , drop = FALSE]
  losses <- xl_pricing_tool_xlsimulation_rlnorm_from_mean_sd(selected_data$MeanLoss, selected_data$StdDevInd)
  data.frame(
    Year = rep(seq_len(simulation_years), counts),
    EventID = paste0("RMS_", seq_len(total)),
    Loss = losses,
    stringsAsFactors = FALSE
  )
}

xl_pricing_tool_xlsimulation_simulate_fs <- function(data, simulation_years) {
  params <- xl_pricing_tool_xlsimulation_parameter_list(data)
  frequency <- xl_pricing_tool_xlsimulation_param_number(params, "Mean Frequency", required = TRUE)
  counts <- stats::rpois(simulation_years, lambda = frequency)
  total <- sum(counts)
  if (total == 0) return(xl_pricing_tool_xlsimulation_empty_events())

  family <- tolower(params[["Severity Family"]])
  losses <- xl_pricing_tool_xlsimulation_sample_severity_family(family, params, total)
  losses <- xl_pricing_tool_xlsimulation_apply_bounds(losses, params)
  data.frame(
    Year = rep(seq_len(simulation_years), counts),
    EventID = paste0("FS_", seq_len(total)),
    Loss = losses,
    stringsAsFactors = FALSE
  )
}

xl_pricing_tool_xlsimulation_simulate_cdf <- function(data, simulation_years) {
  params <- xl_pricing_tool_xlsimulation_parameter_list(data)
  frequency <- xl_pricing_tool_xlsimulation_param_number(params, "Mean Frequency", required = TRUE)
  counts <- stats::rpois(simulation_years, lambda = frequency)
  total <- sum(counts)
  if (total == 0) return(xl_pricing_tool_xlsimulation_empty_events())

  cdf <- xl_pricing_tool_xlsimulation_cdf_table(data)
  losses <- stats::approx(cdf$Percentile, cdf$LossSeverity, xout = stats::runif(total), rule = 2, ties = "ordered")$y
  losses <- xl_pricing_tool_xlsimulation_apply_bounds(losses, params)
  data.frame(
    Year = rep(seq_len(simulation_years), counts),
    EventID = paste0("CDF_", seq_len(total)),
    Loss = losses,
    stringsAsFactors = FALSE
  )
}

xl_pricing_tool_xlsimulation_layer_results <- function(layers, active_causes, coverage, cause_events, settings) {
  output_rows <- list()
  breakdown_rows <- list()
  header <- c(
    "LayerID", "LayerName", "IncludedCauses", "Limit", "Deductible", "Reinstatements", "InputROL",
    "InputPremium", "ReinstatementFactor", "ExpectedLoss_AEP", "StdDev_AEP", "SelectedVaR_OEP",
    "SelectedTVaR_OEP", "SelectedRiskMetric_OEP", "CapitalRequired", "LossCostROL", "InputLossRatio",
    "InputSDMultiple", "InputROE"
  )
  output_rows[[1]] <- header
  breakdown_rows[[1]] <- c("LayerID", active_causes$CauseID)

  for (i in seq_len(nrow(layers))) {
    layer <- layers[i, , drop = FALSE]
    covered_causes <- active_causes$CauseID[coverage[i, ]]
    by_cause <- matrix(0, nrow = settings$SimulationYears, ncol = length(covered_causes))
    colnames(by_cause) <- covered_causes
    oep <- numeric(settings$SimulationYears)

    for (j in seq_along(covered_causes)) {
      cause_id <- covered_causes[[j]]
      events <- cause_events[[cause_id]]
      if (is.null(events) || nrow(events) == 0) next
      layer_loss <- pmin(pmax(events$Loss - layer$Deductible, 0), layer$Limit)
      valid <- layer_loss > 0
      if (!any(valid)) next

      annual_sum <- rowsum(layer_loss[valid], group = events$Year[valid], reorder = FALSE)
      by_cause[as.integer(rownames(annual_sum)), j] <- annual_sum[, 1]
      annual_max <- stats::aggregate(layer_loss[valid], by = list(Year = events$Year[valid]), FUN = max)
      oep[annual_max$Year] <- pmax(oep[annual_max$Year], annual_max$x)
    }

    total_uncapped <- rowSums(by_cause)
    reinstatement_factor <- 1 + layer$Reinstatements * layer$ReinstatementPct
    annual_capacity <- layer$Limit * reinstatement_factor
    annual_aep <- pmin(total_uncapped, annual_capacity)
    cap_factor <- ifelse(total_uncapped > 0, annual_aep / total_uncapped, 0)
    capped_by_cause <- by_cause * cap_factor

    expected_loss <- mean(annual_aep)
    sd_loss <- stats::sd(annual_aep)
    selected_var <- xl_pricing_tool_xlsimulation_var(oep, settings$ReturnPeriod)
    selected_tvar <- xl_pricing_tool_xlsimulation_tvar(oep, selected_var)
    selected_metric <- if (identical(settings$RiskMeasure, "VAR")) selected_var else selected_tvar
    capital <- selected_metric - expected_loss
    input_premium <- layer$Limit * layer$InputROL * reinstatement_factor

    output_rows[[length(output_rows) + 1L]] <- c(
      layer$LayerID,
      layer$LayerName,
      paste(covered_causes, collapse = ", "),
      layer$Limit,
      layer$Deductible,
      layer$Reinstatements,
      layer$InputROL,
      input_premium,
      reinstatement_factor,
      expected_loss,
      sd_loss,
      selected_var,
      selected_tvar,
      selected_metric,
      capital,
      xl_pricing_tool_xlsimulation_safe_div(expected_loss, layer$Limit),
      xl_pricing_tool_xlsimulation_safe_div(expected_loss, input_premium),
      xl_pricing_tool_xlsimulation_safe_div(input_premium - expected_loss, sd_loss),
      xl_pricing_tool_xlsimulation_safe_div(input_premium - expected_loss, capital)
    )

    breakdown_rows[[length(breakdown_rows) + 1L]] <- c(
      layer$LayerID,
      vapply(active_causes$CauseID, function(cause_id) {
        if (cause_id %in% colnames(capped_by_cause)) mean(capped_by_cause[, cause_id]) else 0
      }, numeric(1))
    )
  }

  list(
    layer_output = as.data.frame(do.call(rbind, output_rows), stringsAsFactors = FALSE),
    layer_breakdown = as.data.frame(do.call(rbind, breakdown_rows), stringsAsFactors = FALSE)
  )
}

xl_pricing_tool_xlsimulation_oep_handoff <- function(causes, cause_oep, return_periods) {
  block_starts <- c(1, 5, 9, 13, 17, 24, 31, 34, 37, 41, 44, 47)
  out <- matrix("", nrow = 15, ncol = 48)

  max_blocks <- min(nrow(causes), length(block_starts))
  for (i in seq_len(max_blocks)) {
    col <- block_starts[[i]]
    cause_id <- causes$CauseID[[i]]
    out[1, col] <- "Simulated OEP"
    out[2, col] <- "Return Period"
    out[2, col + 1] <- "OEP"
    out[3:15, col] <- return_periods
    values <- rep(0, length(return_periods))
    if (!is.null(cause_oep[[cause_id]])) {
      values <- vapply(return_periods, function(rp) xl_pricing_tool_xlsimulation_var(cause_oep[[cause_id]], rp), numeric(1))
    }
    out[3:15, col + 1] <- values
  }

  as.data.frame(out, stringsAsFactors = FALSE)
}

xl_pricing_tool_xlsimulation_coverage_handoff <- function(layers, active_causes, coverage) {
  out <- matrix("", nrow = nrow(layers) + 1L, ncol = nrow(active_causes) + 1L)
  out[1, ] <- c("LayerID", active_causes$CauseID)
  out[-1, 1] <- layers$LayerID
  for (i in seq_len(nrow(layers))) {
    for (j in seq_len(nrow(active_causes))) {
      out[i + 1L, j + 1L] <- if (coverage[i, j]) "Y" else "N"
    }
  }
  as.data.frame(out, stringsAsFactors = FALSE)
}

xl_pricing_tool_xlsimulation_breakdown_empty_handoff <- function(layers, active_causes) {
  out <- matrix("", nrow = nrow(layers) + 1L, ncol = nrow(active_causes) + 1L)
  out[1, ] <- c("LayerID", active_causes$CauseID)
  out[-1, 1] <- layers$LayerID
  as.data.frame(out, stringsAsFactors = FALSE)
}

xl_pricing_tool_xlsimulation_oep_return_periods <- function(sim) {
  raw <- xl_pricing_tool_xlsimulation_read_range(sim, "B109:B121", col_names = FALSE)
  out <- suppressWarnings(as.numeric(raw[[1]]))
  out[is.finite(out) & out > 1]
}

xl_pricing_tool_xlsimulation_read_range <- function(sim, range, col_names = FALSE) {
  parsed <- xl_pricing_tool_xlsimulation_parse_location(range)
  if (is.null(parsed)) {
    stop(sprintf("Invalid XLSimulation range: %s", range), call. = FALSE)
  }

  cells <- sim$cells[
    sim$cells$row >= parsed$start_row &
      sim$cells$row <= parsed$end_row &
      sim$cells$col >= parsed$start_col &
      sim$cells$col <= parsed$end_col,
    ,
    drop = FALSE
  ]

  if (nrow(cells) == 0) {
    return(data.frame())
  }

  end_row <- max(cells$row, parsed$start_row)
  values <- matrix("", nrow = end_row - parsed$start_row + 1L, ncol = parsed$end_col - parsed$start_col + 1L)
  values[cbind(cells$row - parsed$start_row + 1L, cells$col - parsed$start_col + 1L)] <- cells$value

  if (isTRUE(col_names)) {
    headers <- make.names(values[1, ], unique = TRUE)
    out <- as.data.frame(values[-1, , drop = FALSE], stringsAsFactors = FALSE)
    names(out) <- headers
  } else {
    out <- as.data.frame(values, stringsAsFactors = FALSE)
  }

  xl_pricing_tool_xlsimulation_trim_empty_edges(out)
}

xl_pricing_tool_xlsimulation_load_cells <- function(workbook_path, sheet_name) {
  temp_dir <- tempfile("xl_pricing_tool_xlsimulation_")
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
  sheet_file <- if (startsWith(target, "/")) sub("^/", "", target) else file.path("xl", target)
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
  if (identical(cell_tags[1], -1L)) {
    return(data.frame(row = integer(), col = integer(), value = character(), stringsAsFactors = FALSE))
  }

  rows <- integer(length(cell_tags))
  cols <- integer(length(cell_tags))
  values <- character(length(cell_tags))
  keep <- logical(length(cell_tags))

  for (i in seq_along(cell_tags)) {
    tag <- cell_tags[[i]]
    ref <- sub('.*\\br="([^"]+)".*', "\\1", tag)
    pos <- xl_pricing_tool_xlsimulation_cell_position(ref)
    if (is.null(pos)) next

    type <- if (grepl('\\bt="s"', tag)) "s" else if (grepl('\\bt="inlineStr"', tag)) "inlineStr" else ""
    value <- ""
    if (grepl("<v>", tag, fixed = TRUE)) {
      raw <- sub(".*<v>(.*?)</v>.*", "\\1", tag)
      if (identical(type, "s")) {
        idx <- suppressWarnings(as.integer(raw) + 1L)
        value <- if (!is.na(idx) && idx <= length(shared)) shared[idx] else raw
      } else {
        value <- xl_pricing_tool_xml_unescape(raw)
      }
    } else if (identical(type, "inlineStr") && grepl("<t", tag, fixed = TRUE)) {
      value <- sub(".*<t[^>]*>(.*?)</t>.*", "\\1", tag)
      value <- xl_pricing_tool_xml_unescape(value)
    }

    if (nzchar(value)) {
      rows[[i]] <- pos$row
      cols[[i]] <- pos$col
      values[[i]] <- value
      keep[[i]] <- TRUE
    }
  }

  data.frame(row = rows[keep], col = cols[keep], value = values[keep], stringsAsFactors = FALSE)
}

xl_pricing_tool_xlsimulation_load_cells_for_ranges <- function(workbook_path, sheet_name, ranges) {
  ranges <- unique(ranges[nzchar(trimws(ranges))])
  bounds <- lapply(ranges, xl_pricing_tool_xlsimulation_parse_location)
  if (any(vapply(bounds, is.null, logical(1)))) {
    stop("Invalid XLSimulation range configuration.", call. = FALSE)
  }

  package <- xl_pricing_tool_xlsimulation_sheet_package(workbook_path, sheet_name)
  on.exit(unlink(package$temp_dir, recursive = TRUE, force = TRUE), add = TRUE)
  shared <- xl_pricing_tool_xlsimulation_shared_strings(package$temp_dir, package$shared_xml)
  sheet_text <- paste(readLines(package$sheet_xml[1], warn = FALSE), collapse = "")

  rows_needed <- sort(unique(unlist(lapply(bounds, function(bound) bound$start_row:bound$end_row))))
  row_pattern <- sprintf("<row\\b[^>]*\\br=\"(%s)\"[^>]*>.*?</row>", paste(rows_needed, collapse = "|"))
  row_tags <- regmatches(sheet_text, gregexpr(row_pattern, sheet_text, perl = TRUE))[[1]]
  if (identical(row_tags[1], -1L)) {
    return(data.frame(row = integer(), col = integer(), value = character(), stringsAsFactors = FALSE))
  }

  cell_tags <- unlist(regmatches(row_tags, gregexpr("<c\\b[^>]*>.*?</c>", row_tags, perl = TRUE)), use.names = FALSE)
  if (length(cell_tags) == 0 || identical(cell_tags[1], -1L)) {
    return(data.frame(row = integer(), col = integer(), value = character(), stringsAsFactors = FALSE))
  }

  xl_pricing_tool_xlsimulation_cell_tags_to_frame(cell_tags, shared, bounds)
}

xl_pricing_tool_xlsimulation_load_source_cells_for_ranges <- function(workbook_path, sheet_name, ranges) {
  ranges <- unique(ranges[nzchar(trimws(ranges))])
  bounds <- lapply(ranges, xl_pricing_tool_xlsimulation_parse_location)
  if (any(vapply(bounds, is.null, logical(1)))) {
    stop("Invalid XLSimulation source range configuration.", call. = FALSE)
  }
  if (!requireNamespace("xml2", quietly = TRUE)) {
    stop("XLSimulation source range reader requires the R package 'xml2'.", call. = FALSE)
  }

  package <- xl_pricing_tool_xlsimulation_sheet_package(workbook_path, sheet_name)
  on.exit(unlink(package$temp_dir, recursive = TRUE, force = TRUE), add = TRUE)
  shared <- xl_pricing_tool_xlsimulation_shared_strings(package$temp_dir, package$shared_xml)
  xl_pricing_tool_xlsimulation_load_cells_for_ranges_xml2(package$sheet_xml[1], shared, bounds)
}

xl_pricing_tool_xlsimulation_load_cells_for_ranges_xml2 <- function(sheet_xml, shared, bounds) {
  doc <- xml2::read_xml(sheet_xml)
  ns <- xml2::xml_ns(doc)
  cells <- xml2::xml_find_all(doc, ".//d1:c[@r]", ns)
  if (length(cells) == 0) {
    return(data.frame(row = integer(), col = integer(), value = character(), stringsAsFactors = FALSE))
  }

  refs <- xml2::xml_attr(cells, "r")
  ref_cols <- sub("[0-9]+$", "", refs)
  ref_rows <- suppressWarnings(as.integer(sub("^[A-Z]+", "", refs)))
  unique_cols <- unique(ref_cols)
  col_lookup <- stats::setNames(vapply(unique_cols, xl_pricing_tool_xlsimulation_col_number, integer(1)), unique_cols)
  ref_col_nums <- unname(col_lookup[ref_cols])

  keep <- rep(FALSE, length(cells))
  for (bound in bounds) {
    keep <- keep |
      (ref_rows >= bound$start_row &
        ref_rows <= bound$end_row &
        ref_col_nums >= bound$start_col &
        ref_col_nums <= bound$end_col)
  }

  if (!any(keep)) {
    return(data.frame(row = integer(), col = integer(), value = character(), stringsAsFactors = FALSE))
  }

  selected <- cells[keep]
  types <- xml2::xml_attr(selected, "t")
  raw_values <- xml2::xml_text(xml2::xml_find_first(selected, "d1:v|d1:is/d1:t", ns))
  raw_values[is.na(raw_values)] <- ""

  values <- raw_values
  shared_idx <- !is.na(types) & types == "s" & nzchar(raw_values)
  if (any(shared_idx)) {
    idx <- suppressWarnings(as.integer(raw_values[shared_idx]) + 1L)
    values[shared_idx] <- ifelse(!is.na(idx) & idx <= length(shared), shared[idx], raw_values[shared_idx])
  }
  values <- xl_pricing_tool_xml_unescape(values)

  nonblank <- nzchar(values)
  data.frame(
    row = ref_rows[keep][nonblank],
    col = ref_col_nums[keep][nonblank],
    value = values[nonblank],
    stringsAsFactors = FALSE
  )
}

xl_pricing_tool_xlsimulation_merge_cells <- function(primary, secondary) {
  if (nrow(primary) == 0) return(secondary)
  if (nrow(secondary) == 0) return(primary)

  out <- rbind(primary, secondary)
  out <- out[!duplicated(paste(out$row, out$col, sep = ":")), , drop = FALSE]
  out[order(out$row, out$col), , drop = FALSE]
}

xl_pricing_tool_xlsimulation_sheet_package <- function(workbook_path, sheet_name) {
  temp_dir <- tempfile("xl_pricing_tool_xlsimulation_")
  dir.create(temp_dir, recursive = TRUE, showWarnings = FALSE)

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
    unlink(temp_dir, recursive = TRUE, force = TRUE)
    stop(sprintf("Workbook is missing sheet: %s", sheet_name), call. = FALSE)
  }

  rid <- sub('.*r:id="([^"]+)".*', "\\1", sheet_tag[1])
  rel_tag <- regmatches(rels, gregexpr("<Relationship\\b[^>]*>", rels, perl = TRUE))[[1]]
  rel_tag <- rel_tag[grepl(sprintf('Id="%s"', rid), rel_tag, fixed = TRUE)]
  if (length(rel_tag) == 0) {
    unlink(temp_dir, recursive = TRUE, force = TRUE)
    stop(sprintf("Workbook relationship is missing for sheet: %s", sheet_name), call. = FALSE)
  }

  target <- sub('.*Target="([^"]+)".*', "\\1", rel_tag[1])
  sheet_file <- if (startsWith(target, "/")) sub("^/", "", target) else file.path("xl", target)
  sheet_xml <- utils::unzip(workbook_path, files = sheet_file, exdir = temp_dir)

  structure(
    list(temp_dir = temp_dir, sheet_xml = sheet_xml, shared_xml = shared_xml),
    class = "xlsimulation_sheet_package"
  )
}

xl_pricing_tool_xlsimulation_shared_strings <- function(temp_dir, shared_xml) {
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
  shared
}

xl_pricing_tool_xlsimulation_cell_tags_to_frame <- function(cell_tags, shared, bounds = list()) {
  rows <- integer(length(cell_tags))
  cols <- integer(length(cell_tags))
  values <- character(length(cell_tags))
  keep <- logical(length(cell_tags))

  for (i in seq_along(cell_tags)) {
    tag <- cell_tags[[i]]
    ref <- sub('.*\\br="([^"]+)".*', "\\1", tag)
    pos <- xl_pricing_tool_xlsimulation_cell_position(ref)
    if (is.null(pos)) next
    if (length(bounds) > 0 && !xl_pricing_tool_xlsimulation_in_bounds(pos, bounds)) next

    type <- if (grepl('\\bt="s"', tag)) "s" else if (grepl('\\bt="inlineStr"', tag)) "inlineStr" else ""
    value <- ""
    if (grepl("<v>", tag, fixed = TRUE)) {
      raw <- sub(".*<v>(.*?)</v>.*", "\\1", tag)
      if (identical(type, "s")) {
        idx <- suppressWarnings(as.integer(raw) + 1L)
        value <- if (!is.na(idx) && idx <= length(shared)) shared[idx] else raw
      } else {
        value <- xl_pricing_tool_xml_unescape(raw)
      }
    } else if (identical(type, "inlineStr") && grepl("<t", tag, fixed = TRUE)) {
      value <- sub(".*<t[^>]*>(.*?)</t>.*", "\\1", tag)
      value <- xl_pricing_tool_xml_unescape(value)
    }

    if (nzchar(value)) {
      rows[[i]] <- pos$row
      cols[[i]] <- pos$col
      values[[i]] <- value
      keep[[i]] <- TRUE
    }
  }

  data.frame(row = rows[keep], col = cols[keep], value = values[keep], stringsAsFactors = FALSE)
}

xl_pricing_tool_xlsimulation_in_bounds <- function(pos, bounds) {
  any(vapply(bounds, function(bound) {
    pos$row >= bound$start_row &&
      pos$row <= bound$end_row &&
      pos$col >= bound$start_col &&
      pos$col <= bound$end_col
  }, logical(1)))
}

xl_pricing_tool_xlsimulation_row_in_bounds <- function(row, bounds) {
  is.finite(row) && any(vapply(bounds, function(bound) {
    row >= bound$start_row && row <= bound$end_row
  }, logical(1)))
}

xl_pricing_tool_xlsimulation_cell_position <- function(ref) {
  match <- regexec("^([A-Z]+)([0-9]+)$", ref)
  parts <- regmatches(ref, match)[[1]]
  if (length(parts) == 0) return(NULL)
  list(
    col = xl_pricing_tool_xlsimulation_col_number(parts[[2]]),
    row = as.integer(parts[[3]])
  )
}

xl_pricing_tool_xlsimulation_parse_location <- function(location) {
  location <- gsub("\\$", "", trimws(location))
  location <- gsub("\\s+", "", location)
  m <- regexec("^([A-Z]+)([0-9]+):([A-Z]+)([0-9]+)$", location)
  parts <- regmatches(location, m)[[1]]
  if (length(parts) == 0) return(NULL)
  list(
    range = location,
    start_col = xl_pricing_tool_xlsimulation_col_number(parts[[2]]),
    start_row = as.integer(parts[[3]]),
    end_col = xl_pricing_tool_xlsimulation_col_number(parts[[4]]),
    end_row = as.integer(parts[[5]])
  )
}

xl_pricing_tool_xlsimulation_trim_empty_edges <- function(data) {
  if (nrow(data) == 0 || ncol(data) == 0) return(data)
  row_keep <- apply(data, 1, function(row) any(nzchar(trimws(as.character(row)))))
  col_keep <- apply(data, 2, function(col) any(nzchar(trimws(as.character(col)))))
  if (!any(row_keep) || !any(col_keep)) {
    return(data.frame())
  }
  data[row_keep, col_keep, drop = FALSE]
}

xl_pricing_tool_xlsimulation_col_number <- function(col) {
  chars <- strsplit(col, "", fixed = TRUE)[[1]]
  out <- 0L
  for (ch in chars) {
    out <- out * 26L + match(ch, LETTERS)
  }
  out
}

xl_pricing_tool_xlsimulation_validate_cause_declarations <- function(causes) {
  duplicate_ids <- unique(causes$CauseID[duplicated(causes$CauseID)])
  if (length(duplicate_ids) > 0) {
    stop(sprintf("XLSimulation CauseID values must be unique. Duplicates: %s", paste(duplicate_ids, collapse = ", ")), call. = FALSE)
  }

  normalized_family <- xl_pricing_tool_xlsimulation_family(causes$CauseFamily)
  normalized_model <- toupper(trimws(causes$ModelSource))

  bad_family <- causes$CauseID[is.na(normalized_family)]
  if (length(bad_family) > 0) {
    stop(sprintf("XLSimulation active cause(s) have unsupported CauseFamily: %s", paste(bad_family, collapse = ", ")), call. = FALSE)
  }

  modeled <- normalized_family == "ModeledCAT"
  bad_modeled <- causes$CauseID[modeled & !normalized_model %in% c("CTR", "TS", "RMS")]
  if (length(bad_modeled) > 0) {
    stop(sprintf("ModeledCAT cause(s) must use ModelSource CTR, TS, or RMS: %s", paste(bad_modeled, collapse = ", ")), call. = FALSE)
  }

  bad_fs_model <- causes$CauseID[normalized_family %in% c("FS", "CDF") & nzchar(normalized_model)]
  if (length(bad_fs_model) > 0) {
    stop(sprintf("FS/CDF cause(s) should leave ModelSource blank: %s", paste(bad_fs_model, collapse = ", ")), call. = FALSE)
  }

  if (sum(modeled) > 6) {
    stop("XLSimulation supports at most 6 active ModeledCAT causes in V1.", call. = FALSE)
  }
  if (sum(normalized_family == "FS") > 3) {
    stop("XLSimulation supports at most 3 active FS causes in V1.", call. = FALSE)
  }
  if (sum(normalized_family == "CDF") > 3) {
    stop("XLSimulation supports at most 3 active CDF causes in V1.", call. = FALSE)
  }

  invisible(TRUE)
}

xl_pricing_tool_xlsimulation_family <- function(value) {
  key <- tolower(trimws(value))
  out <- rep(NA_character_, length(key))
  out[key %in% c("modeledcat", "modeled cat", "cat", "catmodel", "cat model")] <- "ModeledCAT"
  out[key %in% c("fs", "frequency severity", "frequency-severity", "frequencyseverity")] <- "FS"
  out[key %in% c("cdf", "severity cdf", "severitycdf")] <- "CDF"
  out
}

xl_pricing_tool_xlsimulation_drop_blank_rows <- function(data) {
  keep <- apply(data, 1, function(row) any(nzchar(trimws(as.character(row)))))
  data[keep, , drop = FALSE]
}

xl_pricing_tool_xlsimulation_key <- function(value) {
  gsub("[^a-z0-9]", "", tolower(trimws(as.character(value))))
}

xl_pricing_tool_xlsimulation_number <- function(value, field_name) {
  out <- suppressWarnings(as.numeric(value))
  if (is.na(out)) {
    stop(sprintf("XLSimulation setting '%s' must be numeric.", field_name), call. = FALSE)
  }
  out
}

xl_pricing_tool_xlsimulation_numeric_vector <- function(value, field_name) {
  out <- suppressWarnings(as.numeric(value))
  if (any(is.na(out))) {
    stop(sprintf("XLSimulation field '%s' contains nonnumeric values.", field_name), call. = FALSE)
  }
  out
}

xl_pricing_tool_xlsimulation_empty_events <- function() {
  data.frame(Year = integer(), EventID = character(), Loss = numeric(), stringsAsFactors = FALSE)
}

xl_pricing_tool_xlsimulation_annual_oep <- function(events, simulation_years) {
  out <- numeric(simulation_years)
  if (nrow(events) == 0) return(out)
  annual <- stats::aggregate(events$Loss, by = list(Year = events$Year), FUN = max)
  out[annual$Year] <- annual$x
  out
}

xl_pricing_tool_xlsimulation_rlnorm_from_mean_sd <- function(mean, sd) {
  sd <- ifelse(is.na(sd) | sd < 0, 0, sd)
  zero_sd <- sd == 0
  out <- mean
  if (any(!zero_sd)) {
    sigma2 <- log(1 + (sd[!zero_sd]^2 / mean[!zero_sd]^2))
    out[!zero_sd] <- stats::rlnorm(sum(!zero_sd), meanlog = log(mean[!zero_sd]) - sigma2 / 2, sdlog = sqrt(sigma2))
  }
  out
}

xl_pricing_tool_xlsimulation_sample_weighted_cdf <- function(weights, n) {
  weights <- suppressWarnings(as.numeric(weights))
  weights[!is.finite(weights) | weights < 0] <- 0
  total_weight <- sum(weights)
  if (!is.finite(total_weight) || total_weight <= 0) {
    stop("XLSimulation weighted CDF has no positive weights.", call. = FALSE)
  }

  findInterval(stats::runif(n) * total_weight, cumsum(weights)) + 1L
}

xl_pricing_tool_xlsimulation_parameter_list <- function(data) {
  keys <- trimws(as.character(data[[1]]))
  values <- trimws(as.character(data[[2]]))
  params <- as.list(values)
  names(params) <- keys
  params[nzchar(names(params))]
}

xl_pricing_tool_xlsimulation_param_number <- function(params, key, required = FALSE) {
  value <- params[[key]]
  if (is.null(value) || !nzchar(value)) {
    if (isTRUE(required)) stop(sprintf("Missing XLSimulation parameter: %s", key), call. = FALSE)
    return(NA_real_)
  }
  out <- suppressWarnings(as.numeric(value))
  if (is.na(out) && isTRUE(required)) {
    stop(sprintf("XLSimulation parameter '%s' must be numeric.", key), call. = FALSE)
  }
  out
}

xl_pricing_tool_xlsimulation_sample_severity_family <- function(family, params, n) {
  if (identical(family, "pareto")) {
    alpha <- xl_pricing_tool_xlsimulation_param_number(params, "alpha", required = TRUE)
    scale <- xl_pricing_tool_xlsimulation_param_number(params, "Minimum Loss", required = FALSE)
    if (is.na(scale) || scale <= 0) scale <- 1
    scale / (1 - stats::runif(n))^(1 / alpha)
  } else if (identical(family, "lognormal")) {
    stats::rlnorm(
      n,
      meanlog = xl_pricing_tool_xlsimulation_param_number(params, "meanlog", required = TRUE),
      sdlog = xl_pricing_tool_xlsimulation_param_number(params, "sigma", required = TRUE)
    )
  } else if (identical(family, "gamma")) {
    stats::rgamma(
      n,
      shape = xl_pricing_tool_xlsimulation_param_number(params, "shape", required = TRUE),
      rate = xl_pricing_tool_xlsimulation_param_number(params, "rate", required = TRUE)
    )
  } else if (identical(family, "loggamma")) {
    exp(stats::rgamma(
      n,
      shape = xl_pricing_tool_xlsimulation_param_number(params, "shape", required = TRUE),
      rate = xl_pricing_tool_xlsimulation_param_number(params, "rate", required = TRUE)
    ))
  } else if (identical(family, "weibull")) {
    stats::rweibull(
      n,
      shape = xl_pricing_tool_xlsimulation_param_number(params, "shape", required = TRUE),
      scale = xl_pricing_tool_xlsimulation_param_number(params, "scale", required = TRUE)
    )
  } else if (identical(family, "loglogistic")) {
    shape <- xl_pricing_tool_xlsimulation_param_number(params, "shape", required = TRUE)
    scale <- xl_pricing_tool_xlsimulation_param_number(params, "scale", required = TRUE)
    u <- stats::runif(n)
    scale * (u / (1 - u))^(1 / shape)
  } else {
    stop(sprintf("Unsupported severity family: %s", family), call. = FALSE)
  }
}

xl_pricing_tool_xlsimulation_apply_bounds <- function(losses, params) {
  min_loss <- xl_pricing_tool_xlsimulation_param_number(params, "Minimum Loss", required = FALSE)
  cap <- xl_pricing_tool_xlsimulation_param_number(params, "Loss Cap", required = FALSE)
  if (!is.na(min_loss)) losses <- pmax(losses, min_loss)
  if (!is.na(cap) && cap > 0) losses <- pmin(losses, cap)
  losses
}

xl_pricing_tool_xlsimulation_cdf_table <- function(data) {
  keys <- trimws(as.character(data[[1]]))
  start <- which(keys == "Percentile")
  if (length(start) == 0) {
    stop("CDF table is missing Percentile/Loss Severity header.", call. = FALSE)
  }
  cdf <- data[(start[[1]] + 1L):nrow(data), 1:2, drop = FALSE]
  out <- data.frame(
    Percentile = suppressWarnings(as.numeric(cdf[[1]])),
    LossSeverity = suppressWarnings(as.numeric(cdf[[2]])),
    stringsAsFactors = FALSE
  )
  out <- out[is.finite(out$Percentile) & is.finite(out$LossSeverity), , drop = FALSE]
  if (nrow(out) == 0) stop("CDF table has no numeric percentile/loss rows.", call. = FALSE)
  if (max(out$Percentile, na.rm = TRUE) > 1) out$Percentile <- out$Percentile / 100
  out <- out[order(out$Percentile), , drop = FALSE]
  out
}

xl_pricing_tool_xlsimulation_var <- function(values, return_period) {
  if (length(values) == 0) return(0)
  as.numeric(stats::quantile(values, probs = 1 - 1 / return_period, na.rm = TRUE, names = FALSE, type = 8))
}

xl_pricing_tool_xlsimulation_tvar <- function(values, var_value) {
  tail <- values[values >= var_value]
  if (length(tail) == 0) return(var_value)
  mean(tail, na.rm = TRUE)
}

xl_pricing_tool_xlsimulation_safe_div <- function(numerator, denominator) {
  if (is.na(denominator) || denominator == 0) return(NA_real_)
  numerator / denominator
}
