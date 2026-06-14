xl_pricing_tool_profile_update <- function(workbook_path, output_dir = NULL, context = list()) {
  xl_pricing_tool_validate(workbook_path, context)

  if (is.null(output_dir) || identical(output_dir, "")) {
    output_dir <- if (!is.null(context$output_dir)) {
      context$output_dir
    } else {
      btk_default_run_dir("xl_pricing_tool", workbook_path)
    }
  }

  profile_dir <- file.path(output_dir, "profile")
  dir.create(profile_dir, recursive = TRUE, showWarnings = FALSE)

  sheet <- xl_pricing_tool_read_sheet(workbook_path, "Input_Profile")
  profile_input <- xl_pricing_tool_profile_read_input(sheet)
  layers <- xl_pricing_tool_profile_read_layers(sheet)
  params <- xl_pricing_tool_profile_read_params(sheet)
  band_cuts <- xl_pricing_tool_profile_read_band_cuts(sheet)
  facets <- xl_pricing_tool_profile_read_facets(sheet, profile_input$LOB)
  profile_output <- xl_pricing_tool_profile_calculate(profile_input, layers, params)
  handoff <- xl_pricing_tool_profile_handoff(profile_output$lob_level, profile_output$layer_level)
  chart_data <- xl_pricing_tool_profile_chart_data(profile_input, band_cuts, facets)

  output_file <- file.path(profile_dir, "profile_output.csv")
  chart_data_file <- file.path(profile_dir, "profile_si_composition_data.csv")
  chart_file <- file.path(profile_dir, "profile_si_composition_chart.png")
  utils::write.table(
    handoff,
    file = output_file,
    sep = ",",
    row.names = FALSE,
    col.names = TRUE,
    na = "",
    quote = TRUE
  )
  utils::write.csv(chart_data, chart_data_file, row.names = FALSE, na = "")
  xl_pricing_tool_render_profile_chart(chart_data, chart_file)

  paste(
    "Profile update completed.",
    sprintf("Output folder: %s", output_dir),
    sprintf("Input rows: %s", nrow(profile_input)),
    sprintf("Layer rows: %s", nrow(layers)),
    sprintf("LOB-level output rows: %s", nrow(profile_output$lob_level)),
    sprintf("Layer-level output rows: %s", nrow(profile_output$layer_level)),
    sprintf("Profile output: %s", output_file),
    sprintf("Profile chart data: %s", chart_data_file),
    sprintf("Profile chart image: %s", chart_file),
    "CSV output and chart image are handled by the VBA client after BERT returns.",
    sep = vb_newline()
  )
}

xl_pricing_tool_profile_gather <- function(workbook_path, output_dir = NULL, context = list()) {
  "No gather action defined for Input_Profile because Gather From is blank."
}

xl_pricing_tool_profile_build <- function(workbook_path, output_dir = NULL, context = list()) {
  "Profile build is not implemented yet. This stage will generate the simulated loss table after the schema and curve mechanics are approved."
}

xl_pricing_tool_profile_read_input <- function(sheet) {
  headers <- vapply(2:9, function(col) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), 12)))
  }, character(1))
  expected <- c("AsAt", "UY", "LOB", "Nrisk", "SI", "Premium", "Incurred", "Remarks")
  if (!all(headers == expected)) {
    stop("Input_Profile risk profile table must have AsAt, UY, LOB, Nrisk, SI, Premium, Incurred, Remarks headers in B12:I12.", call. = FALSE)
  }

  rows <- list()
  blank_run <- 0
  for (row in 13:(13 + 5000)) {
    as_at <- trimws(xl_pricing_tool_cell_value(sheet, paste0("B", row)))
    lob <- trimws(xl_pricing_tool_cell_value(sheet, paste0("D", row)))

    if (!nzchar(as_at) && !nzchar(lob)) {
      blank_run <- blank_run + 1
      if (blank_run >= 20) break
      next
    }
    blank_run <- 0

    if (!nzchar(as_at) || !nzchar(lob)) {
      stop(sprintf("Input_Profile row %s has blank AsAt or LOB.", row), call. = FALSE)
    }

    as_at <- xl_pricing_tool_profile_normalize_asat(as_at, row)
    nrisk <- xl_pricing_tool_profile_number(sheet, "E", row, "Nrisk", required = TRUE)
    si <- xl_pricing_tool_profile_number(sheet, "F", row, "SI", required = TRUE)
    premium <- xl_pricing_tool_profile_number(sheet, "G", row, "Premium", required = FALSE)
    incurred <- xl_pricing_tool_profile_number(sheet, "H", row, "Incurred", required = FALSE)

    if (nrisk <= 0 && si != 0) {
      stop(sprintf("Input_Profile row %s has nonzero SI but Nrisk is not positive.", row), call. = FALSE)
    }
    if (nrisk < 0) {
      stop(sprintf("Input_Profile row %s has negative Nrisk.", row), call. = FALSE)
    }
    if (si < 0) {
      stop(sprintf("Input_Profile row %s has negative SI.", row), call. = FALSE)
    }

    rows[[length(rows) + 1L]] <- data.frame(
      AsAt = as_at,
      UY = trimws(xl_pricing_tool_cell_value(sheet, paste0("C", row))),
      LOB = lob,
      Nrisk = nrisk,
      SI = si,
      Premium = premium,
      Incurred = incurred,
      Remarks = trimws(xl_pricing_tool_cell_value(sheet, paste0("I", row))),
      stringsAsFactors = FALSE
    )
  }

  if (length(rows) == 0) {
    stop("Input_Profile risk profile table has no rows.", call. = FALSE)
  }

  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

xl_pricing_tool_profile_read_layers <- function(sheet) {
  headers <- vapply(24:29, function(col) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), 13)))
  }, character(1))
  expected <- c("LayerID", "Limit", "Deductible", "LayerName", "Prior_LOBs", "Current_LOBs")
  if (!all(headers == expected)) {
    stop("Input_Profile layer table must have LayerID, Limit, Deductible, LayerName, Prior_LOBs, Current_LOBs headers in X13:AC13.", call. = FALSE)
  }

  rows <- list()
  blank_run <- 0
  for (row in 15:(15 + 500)) {
    layer_id <- trimws(xl_pricing_tool_cell_value(sheet, paste0("X", row)))
    layer_name <- trimws(xl_pricing_tool_cell_value(sheet, paste0("AA", row)))

    if (!nzchar(layer_id) && !nzchar(layer_name)) {
      blank_run <- blank_run + 1
      if (blank_run >= 20) break
      next
    }
    blank_run <- 0

    if (!nzchar(layer_id) || !nzchar(layer_name)) {
      stop(sprintf("Input_Profile layer row %s has blank LayerID or LayerName.", row), call. = FALSE)
    }

    limit <- xl_pricing_tool_profile_number(sheet, "Y", row, "Limit", required = TRUE)
    deductible <- xl_pricing_tool_profile_number(sheet, "Z", row, "Deductible", required = TRUE)
    if (limit < 0 || deductible < 0) {
      stop(sprintf("Input_Profile layer row %s has negative Limit or Deductible.", row), call. = FALSE)
    }

    rows[[length(rows) + 1L]] <- data.frame(
      LayerID = layer_id,
      Limit = limit,
      Deductible = deductible,
      LayerName = layer_name,
      Prior_LOBs = trimws(xl_pricing_tool_cell_value(sheet, paste0("AB", row))),
      Current_LOBs = trimws(xl_pricing_tool_cell_value(sheet, paste0("AC", row))),
      stringsAsFactors = FALSE
    )
  }

  if (length(rows) == 0) {
    stop("Input_Profile layer table has no rows.", call. = FALSE)
  }

  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

xl_pricing_tool_profile_read_params <- function(sheet) {
  rows <- list()
  for (row in c(13:22, 24:33)) {
    lob <- trimws(xl_pricing_tool_cell_value(sheet, paste0("M", row)))
    loss_cause <- trimws(xl_pricing_tool_cell_value(sheet, paste0("N", row)))
    if (!nzchar(lob) || !nzchar(loss_cause) || identical(tolower(lob), "lob")) {
      next
    }

    as_at <- if (grepl("^P_Expo_", loss_cause, ignore.case = TRUE)) {
      "Prior"
    } else if (grepl("^C_Expo_", loss_cause, ignore.case = TRUE)) {
      "Current"
    } else {
      next
    }

    subject_prem <- xl_pricing_tool_profile_number(sheet, "P", row, "SubjectPrem", required = FALSE)
    actual_prem <- xl_pricing_tool_profile_number(sheet, "Q", row, "ActualPrem", required = FALSE)
    if (actual_prem < 0 || subject_prem < 0) {
      stop(sprintf("Input_Profile parameter row %s has negative SubjectPrem or ActualPrem.", row), call. = FALSE)
    }

    rows[[length(rows) + 1L]] <- data.frame(
      AsAt = as_at,
      LOB = lob,
      SubjectPrem = subject_prem,
      ActualPrem = actual_prem,
      stringsAsFactors = FALSE
    )
  }

  if (length(rows) == 0) {
    stop("Input_Profile exposure rating parameters have no prior/current LOB rows.", call. = FALSE)
  }

  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

xl_pricing_tool_profile_calculate <- function(profile_input, layers, params) {
  rows <- list()
  for (layer_index in seq_len(nrow(layers))) {
    layer <- layers[layer_index, , drop = FALSE]
    for (as_at in c("Current", "Prior")) {
      coverage_text <- if (identical(as_at, "Prior")) layer$Prior_LOBs else layer$Current_LOBs
      profile_rows <- profile_input[profile_input$AsAt == as_at, , drop = FALSE]
      included <- vapply(profile_rows$LOB, xl_pricing_tool_profile_lob_included, logical(1), coverage_text = coverage_text)
      profile_rows <- profile_rows[included, , drop = FALSE]
      if (nrow(profile_rows) == 0) {
        next
      }

      avg_si <- ifelse(profile_rows$Nrisk > 0, profile_rows$SI / profile_rows$Nrisk, 0)
      row_exposure <- profile_rows$Nrisk * pmin(pmax(avg_si - layer$Deductible, 0), layer$Limit)
      for (i in seq_len(nrow(profile_rows))) {
        rows[[length(rows) + 1L]] <- data.frame(
          AsAt = as_at,
          LOB = profile_rows$LOB[[i]],
          LayerName = layer$LayerName,
          expo_in_layer = row_exposure[[i]],
          stringsAsFactors = FALSE
        )
      }
    }
  }

  if (length(rows) == 0) {
    stop("Profile update produced no exposure-in-layer rows. Check layer coverage LOBs and risk profile input.", call. = FALSE)
  }

  detail <- do.call(rbind, rows)
  lob_level <- aggregate(
    expo_in_layer ~ AsAt + LOB + LayerName,
    detail,
    sum,
    na.rm = TRUE
  )
  lob_level <- merge(lob_level, params, by = c("AsAt", "LOB"), all.x = TRUE, sort = FALSE)

  missing_params <- lob_level[is.na(lob_level$ActualPrem) | is.na(lob_level$SubjectPrem), c("AsAt", "LOB"), drop = FALSE]
  if (nrow(missing_params) > 0) {
    missing_text <- paste(unique(sprintf("%s/%s", missing_params$AsAt, missing_params$LOB)), collapse = ", ")
    stop(sprintf("Input_Profile exposure rating parameters are missing for: %s", missing_text), call. = FALSE)
  }

  lob_level$adj_expo_in_layer <- ifelse(
    lob_level$ActualPrem > 0,
    lob_level$expo_in_layer * lob_level$SubjectPrem / lob_level$ActualPrem,
    0
  )
  lob_level <- lob_level[, c("AsAt", "LOB", "LayerName", "expo_in_layer", "adj_expo_in_layer")]
  lob_level <- lob_level[order(match(lob_level$AsAt, c("Current", "Prior")), lob_level$LOB, lob_level$LayerName), , drop = FALSE]
  rownames(lob_level) <- NULL

  layer_level <- aggregate(
    cbind(expo_in_layer, adj_expo_in_layer) ~ AsAt + LayerName,
    lob_level,
    sum,
    na.rm = TRUE
  )
  covered <- aggregate(
    LOB ~ AsAt + LayerName,
    lob_level,
    function(x) paste(unique(x), collapse = ", ")
  )
  layer_level <- merge(layer_level, covered, by = c("AsAt", "LayerName"), all.x = TRUE, sort = FALSE)
  names(layer_level)[names(layer_level) == "LOB"] <- "covered_lob"
  layer_level$tempid <- paste(layer_level$LayerName, layer_level$AsAt, sep = "_")
  layer_level <- layer_level[, c("AsAt", "LayerName", "expo_in_layer", "adj_expo_in_layer", "covered_lob", "tempid")]
  layer_level <- layer_level[order(match(layer_level$AsAt, c("Current", "Prior")), layer_level$LayerName), , drop = FALSE]
  rownames(layer_level) <- NULL

  list(lob_level = lob_level, layer_level = layer_level)
}

xl_pricing_tool_profile_handoff <- function(lob_level, layer_level) {
  max_rows <- max(nrow(lob_level), nrow(layer_level))
  blank <- rep("", max_rows)

  pad_col <- function(values) {
    values <- as.character(values)
    c(values, rep("", max_rows - length(values)))
  }

  out <- data.frame(
    AsAt = pad_col(lob_level$AsAt),
    LOB = pad_col(lob_level$LOB),
    LayerName = pad_col(lob_level$LayerName),
    expo_in_layer = pad_col(format(lob_level$expo_in_layer, scientific = FALSE, trim = TRUE)),
    adj_expo_in_layer = pad_col(format(lob_level$adj_expo_in_layer, scientific = FALSE, trim = TRUE)),
    blank,
    AsAt_rollup = pad_col(layer_level$AsAt),
    LayerName_rollup = pad_col(layer_level$LayerName),
    expo_in_layer_rollup = pad_col(format(layer_level$expo_in_layer, scientific = FALSE, trim = TRUE)),
    adj_expo_in_layer_rollup = pad_col(format(layer_level$adj_expo_in_layer, scientific = FALSE, trim = TRUE)),
    covered_lob = pad_col(layer_level$covered_lob),
    tempid = pad_col(layer_level$tempid),
    check.names = FALSE,
    stringsAsFactors = FALSE
  )
  names(out) <- c(
    "AsAt",
    "LOB",
    "LayerName",
    "expo_in_layer",
    "adj_expo_in_layer",
    "",
    "AsAt",
    "LayerName",
    "expo_in_layer",
    "adj_expo_in_layer",
    "covered_lob",
    "tempid"
  )
  out
}

xl_pricing_tool_profile_read_band_cuts <- function(sheet) {
  rows <- 37:44
  values <- vapply(rows, function(row) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0("K", row)))
  }, character(1))
  names(values) <- paste0("K", rows)

  label_rows <- tolower(values) %in% c("band cut", "band cuts", "cut", "cuts")
  values <- values[!label_rows]
  values <- values[nzchar(values)]

  cuts <- suppressWarnings(as.numeric(gsub(",", "", values, fixed = TRUE)))
  if (any(is.na(cuts))) {
    bad <- values[is.na(cuts)]
    stop(sprintf("Input_Profile band specification K37:K44 must be numeric after the Band Cut label. Invalid value(s): %s", paste(bad, collapse = ", ")), call. = FALSE)
  }
  if (any(cuts < 0)) {
    stop("Input_Profile band specification K37:K44 cannot contain negative values.", call. = FALSE)
  }

  cuts <- sort(unique(cuts))
  if (length(cuts) < 2) {
    stop("Input_Profile band specification K37:K44 must contain at least two unique numeric cut points.", call. = FALSE)
  }
  cuts
}

xl_pricing_tool_profile_read_facets <- function(sheet, available_lobs) {
  raw_values <- vapply(47:55, function(row) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0("K", row)))
  }, character(1))
  raw_values <- raw_values[nzchar(raw_values)]
  raw_values <- raw_values[tolower(raw_values) != "lob"]

  if (length(raw_values) == 0) {
    return(data.frame(Facet = "All LOBs", LOB = unique(available_lobs), stringsAsFactors = FALSE))
  }

  rows <- list()
  for (facet in raw_values) {
    tokens <- xl_pricing_tool_profile_split_lobs(facet)
    if (length(tokens) == 0) {
      next
    }

    matched <- xl_pricing_tool_profile_match_lobs(tokens, available_lobs)
    if (length(matched) == 0) {
      stop(sprintf("Input_Profile facet specification '%s' in K47:K55 does not match any LOB in the risk profile input.", facet), call. = FALSE)
    }

    for (lob in matched) {
      rows[[length(rows) + 1L]] <- data.frame(Facet = facet, LOB = lob, stringsAsFactors = FALSE)
    }
  }

  if (length(rows) == 0) {
    return(data.frame(Facet = "All LOBs", LOB = unique(available_lobs), stringsAsFactors = FALSE))
  }

  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

xl_pricing_tool_profile_chart_data <- function(profile_input, band_cuts, facets) {
  profile_input$Avg_SI <- ifelse(profile_input$Nrisk > 0, profile_input$SI / profile_input$Nrisk, 0)
  profile_input$Band <- xl_pricing_tool_profile_band_labels(profile_input$Avg_SI, band_cuts)

  rows <- list()
  for (facet in unique(facets$Facet)) {
    facet_lobs <- facets$LOB[facets$Facet == facet]
    facet_input <- profile_input[tolower(profile_input$LOB) %in% tolower(facet_lobs), , drop = FALSE]
    if (nrow(facet_input) == 0) {
      next
    }

    totals <- aggregate(SI ~ Facet + AsAt + Band, transform(facet_input, Facet = facet), sum, na.rm = TRUE)
    totals$TotalSI <- ave(totals$SI, totals$Facet, totals$AsAt, FUN = sum)
    totals$Percent <- ifelse(totals$TotalSI > 0, totals$SI / totals$TotalSI, NA_real_)
    rows[[length(rows) + 1L]] <- totals
  }

  if (length(rows) == 0) {
    stop("Profile chart produced no data. Check the facet LOB specifications in K47:K55.", call. = FALSE)
  }

  out <- do.call(rbind, rows)
  out$AsAt <- factor(out$AsAt, levels = c("Prior", "Current"))
  out$Band <- factor(out$Band, levels = xl_pricing_tool_profile_band_label_levels(band_cuts))
  out[order(out$Facet, out$AsAt, out$Band), c("Facet", "AsAt", "Band", "SI", "Percent"), drop = FALSE]
}

xl_pricing_tool_render_profile_chart <- function(chart_data, output_file) {
  if (!requireNamespace("ggplot2", quietly = TRUE)) {
    stop("Package 'ggplot2' is required to render Profile chart images.", call. = FALSE)
  }
  if (nrow(chart_data) == 0) {
    stop("Profile chart data is empty.", call. = FALSE)
  }

  ggplot2 <- asNamespace("ggplot2")
  plot <- ggplot2$ggplot(chart_data, ggplot2$aes(x = AsAt, y = Percent, fill = Band)) +
    ggplot2$geom_bar(stat = "identity", width = 0.68, color = "white") +
    ggplot2$facet_wrap(stats::as.formula("~ Facet")) +
    ggplot2$scale_y_continuous(
      labels = function(x) paste0(round(x * 100), "%"),
      limits = c(0, 1),
      expand = c(0, 0)
    ) +
    ggplot2$scale_fill_brewer(palette = "Blues", direction = 1) +
    ggplot2$labs(
      title = "Sum Insured Composition by Risk Band",
      subtitle = "Percent of SI by prior/current profile and selected LOB facets",
      x = NULL,
      y = NULL,
      fill = "Avg SI Band"
    ) +
    ggplot2$theme_minimal(base_size = 9) +
    ggplot2$theme(
      plot.title = ggplot2$element_text(face = "bold", color = "#22313F", size = 12),
      plot.subtitle = ggplot2$element_text(color = "#52606D", size = 9),
      panel.grid.major.x = ggplot2$element_blank(),
      panel.grid.minor = ggplot2$element_blank(),
      panel.grid.major.y = ggplot2$element_line(color = "#E6EDF3"),
      axis.text = ggplot2$element_text(color = "#52606D"),
      strip.text = ggplot2$element_text(face = "bold", color = "#22313F"),
      legend.position = "bottom"
    )

  ggplot2$ggsave(output_file, plot = plot, width = 8.6, height = 4.8, dpi = 160, bg = "white")
}

xl_pricing_tool_profile_band_labels <- function(values, cuts) {
  labels <- xl_pricing_tool_profile_band_label_levels(cuts)
  idx <- findInterval(values, cuts, rightmost.closed = FALSE, all.inside = FALSE)
  idx <- pmax(1, pmin(idx, length(labels)))
  labels[idx]
}

xl_pricing_tool_profile_band_label_levels <- function(cuts) {
  labels <- character(length(cuts))
  for (i in seq_along(cuts)) {
    if (i < length(cuts)) {
      labels[[i]] <- sprintf("%s - %s", xl_pricing_tool_profile_format_band(cuts[[i]]), xl_pricing_tool_profile_format_band(cuts[[i + 1L]]))
    } else {
      labels[[i]] <- sprintf("%s+", xl_pricing_tool_profile_format_band(cuts[[i]]))
    }
  }
  labels
}

xl_pricing_tool_profile_format_band <- function(value) {
  if (abs(value) >= 1e9) {
    return(sprintf("%sB", format(round(value / 1e9, 1), trim = TRUE, nsmall = 0)))
  }
  if (abs(value) >= 1e6) {
    return(sprintf("%sM", format(round(value / 1e6, 1), trim = TRUE, nsmall = 0)))
  }
  format(round(value, 0), big.mark = ",", scientific = FALSE, trim = TRUE)
}

xl_pricing_tool_profile_lob_included <- function(lob, coverage_text) {
  coverage <- trimws(as.character(coverage_text))
  if (!nzchar(coverage)) {
    return(TRUE)
  }

  tokens <- xl_pricing_tool_profile_split_lobs(coverage)
  if (length(tokens) == 0) {
    return(TRUE)
  }

  tolower(trimws(lob)) %in% tolower(tokens)
}

xl_pricing_tool_profile_split_lobs <- function(value) {
  tokens <- unlist(strsplit(trimws(as.character(value)), "[,;/]+", perl = TRUE), use.names = FALSE)
  tokens <- trimws(tokens)
  tokens[nzchar(tokens)]
}

xl_pricing_tool_profile_match_lobs <- function(tokens, available_lobs) {
  matched <- character()
  for (token in tokens) {
    exact <- available_lobs[tolower(available_lobs) == tolower(token)]
    if (length(exact) > 0) {
      matched <- c(matched, exact)
      next
    }

    normalized_token <- gsub("[^a-z0-9]", "", tolower(token))
    normalized_lobs <- gsub("[^a-z0-9]", "", tolower(available_lobs))
    normalized <- available_lobs[normalized_lobs == normalized_token]
    if (length(normalized) > 0) {
      matched <- c(matched, normalized)
    }
  }
  unique(matched)
}

xl_pricing_tool_profile_number <- function(sheet, col, row, field_name, required = TRUE) {
  raw_value <- trimws(xl_pricing_tool_cell_value(sheet, paste0(col, row)))
  if (!nzchar(raw_value)) {
    if (required) {
      stop(sprintf("Input_Profile %s cell %s%s is blank.", field_name, col, row), call. = FALSE)
    }
    return(0)
  }

  value <- suppressWarnings(as.numeric(gsub(",", "", raw_value, fixed = TRUE)))
  if (is.na(value)) {
    stop(sprintf("Input_Profile %s cell %s%s is not numeric: %s", field_name, col, row, raw_value), call. = FALSE)
  }
  value
}

xl_pricing_tool_profile_normalize_asat <- function(value, row) {
  key <- tolower(trimws(value))
  if (key %in% c("prior", "p")) {
    return("Prior")
  }
  if (key %in% c("current", "curr", "c")) {
    return("Current")
  }
  stop(sprintf("Input_Profile row %s has unsupported AsAt '%s'. Supported values: Prior, Current.", row, value), call. = FALSE)
}
