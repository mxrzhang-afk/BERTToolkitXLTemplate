xl_pricing_tool_riskfit_update <- function(workbook_path, output_dir = NULL, context = list()) {
  xl_pricing_tool_validate(workbook_path, context)

  if (is.null(output_dir) || identical(output_dir, "")) {
    output_dir <- if (!is.null(context$output_dir)) {
      context$output_dir
    } else {
      btk_default_run_dir("xl_pricing_tool", workbook_path)
    }
  }

  riskfit_dir <- file.path(output_dir, "riskfit")
  dir.create(riskfit_dir, recursive = TRUE, showWarnings = FALSE)

  sheet <- xl_pricing_tool_read_sheet(workbook_path, "Risk Loss Fitting")
  loss_input <- xl_pricing_tool_riskfit_read_loss_input(sheet)
  chart_config <- xl_pricing_tool_riskfit_read_loss_comparison_config(sheet)
  layers <- xl_pricing_tool_riskfit_read_loss_layers(sheet)
  chart_data <- xl_pricing_tool_riskfit_loss_comparison_data(loss_input, chart_config)

  chart_data_file <- file.path(riskfit_dir, "riskfit_loss_comparison_data.csv")
  final_chart_file <- file.path(riskfit_dir, "riskfit_final_incurred_loss_chart.png")
  actual_chart_file <- file.path(riskfit_dir, "riskfit_actual_incurred_loss_chart.png")

  utils::write.csv(chart_data, chart_data_file, row.names = FALSE, na = "")
  final_chart_data <- chart_data[chart_data$ChartType == "Final_Incurred", , drop = FALSE]
  actual_chart_data <- chart_data[chart_data$ChartType == "Actual_Incurred", , drop = FALSE]

  xl_pricing_tool_render_riskfit_loss_comparison_chart(
    chart_data = final_chart_data,
    layers = layers,
    chart_config = chart_config,
    subtitle = chart_config$adjusted_subtitle,
    output_file = final_chart_file
  )
  if (chart_config$show_actual) {
    xl_pricing_tool_render_riskfit_loss_comparison_chart(
      chart_data = actual_chart_data,
      layers = layers,
      chart_config = chart_config,
      subtitle = chart_config$actual_subtitle,
      output_file = actual_chart_file
    )
  }

  paste(
    "RiskFit update completed.",
    sprintf("Output folder: %s", output_dir),
    sprintf("Input rows: %s", nrow(loss_input)),
    sprintf("Chart rows: %s", nrow(chart_data)),
    sprintf("Layer rows: %s", nrow(layers)),
    sprintf("Loss comparison chart data: %s", chart_data_file),
    sprintf("Final incurred chart image: %s", final_chart_file),
    if (chart_config$show_actual) sprintf("Actual incurred chart image: %s", actual_chart_file) else "Actual incurred chart image: skipped because showActual is FALSE.",
    "RiskFit chart image is handled by the VBA client after BERT returns.",
    sep = vb_newline()
  )
}

xl_pricing_tool_riskfit_gather <- function(workbook_path, output_dir = NULL, context = list()) {
  "No gather action defined for Risk Loss Fitting because Gather From is blank."
}

xl_pricing_tool_riskfit_build <- function(workbook_path, output_dir = NULL, context = list()) {
  "No build action defined for Risk Loss Fitting."
}

xl_pricing_tool_riskfit_read_loss_input <- function(sheet) {
  expected <- c(
    "AsAt",
    "ClaimID",
    "LOB",
    "Insured",
    "AY",
    "TreatyYear",
    "Actual_Incurred",
    "OnLevel_Incurred",
    "ReturnPeriod",
    "OriginalRate",
    "GNPI_Scale",
    "RP_Scale",
    "FinalRate",
    "Sev_Trend",
    "Final_Incurred"
  )
  headers <- vapply(2:16, function(col) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), 12)))
  }, character(1))
  if (!all(headers == expected)) {
    stop("Risk Loss Fitting loss table must have expected headers in B12:P12.", call. = FALSE)
  }

  rows <- list()
  blank_run <- 0
  for (row in 13:(13 + 5000)) {
    as_at <- trimws(xl_pricing_tool_cell_value(sheet, paste0("B", row)))
    claim_id <- trimws(xl_pricing_tool_cell_value(sheet, paste0("C", row)))
    insured <- trimws(xl_pricing_tool_cell_value(sheet, paste0("E", row)))

    if (!nzchar(as_at) && !nzchar(claim_id) && !nzchar(insured)) {
      blank_run <- blank_run + 1
      if (blank_run >= 20) break
      next
    }
    blank_run <- 0

    if (!nzchar(as_at)) {
      stop(sprintf("Risk Loss Fitting row %s has blank AsAt.", row), call. = FALSE)
    }
    as_at <- xl_pricing_tool_riskfit_normalize_asat(as_at, row)

    treaty_year <- xl_pricing_tool_riskfit_number(sheet, "G", row, "TreatyYear", required = TRUE)
    actual_incurred <- xl_pricing_tool_riskfit_number(sheet, "H", row, "Actual_Incurred", required = FALSE)
    final_incurred <- xl_pricing_tool_riskfit_number(sheet, "P", row, "Final_Incurred", required = FALSE)

    rows[[length(rows) + 1L]] <- data.frame(
      AsAt = as_at,
      ClaimID = claim_id,
      LOB = trimws(xl_pricing_tool_cell_value(sheet, paste0("D", row))),
      Insured = insured,
      TreatyYear = as.integer(treaty_year),
      Actual_Incurred = actual_incurred,
      Final_Incurred = final_incurred,
      stringsAsFactors = FALSE
    )
  }

  if (length(rows) == 0) {
    stop("Risk Loss Fitting loss table has no input rows.", call. = FALSE)
  }

  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

xl_pricing_tool_riskfit_read_loss_comparison_config <- function(sheet) {
  min_loss <- xl_pricing_tool_riskfit_number(sheet, "T", 56, "loss comparison minimum loss", required = TRUE)
  start_year <- xl_pricing_tool_riskfit_number(sheet, "T", 57, "loss comparison start year", required = TRUE)
  show_actual <- xl_pricing_tool_riskfit_bool(sheet, "T", 58, default = TRUE)
  title <- trimws(xl_pricing_tool_cell_value(sheet, "Z56"))
  adjusted_subtitle <- trimws(xl_pricing_tool_cell_value(sheet, "Z57"))
  actual_subtitle <- trimws(xl_pricing_tool_cell_value(sheet, "Z58"))

  if (!nzchar(title)) {
    stop("Risk Loss Fitting loss comparison chart title in Z56 cannot be blank.", call. = FALSE)
  }
  if (!nzchar(adjusted_subtitle)) {
    stop("Risk Loss Fitting adjusted loss subtitle in Z57 cannot be blank.", call. = FALSE)
  }
  if (show_actual && !nzchar(actual_subtitle)) {
    stop("Risk Loss Fitting actual loss subtitle in Z58 cannot be blank when showActual is TRUE.", call. = FALSE)
  }
  if (min_loss < 0) {
    stop("Risk Loss Fitting loss comparison minimum loss cannot be negative.", call. = FALSE)
  }
  if (start_year < 1900) {
    stop("Risk Loss Fitting loss comparison start year must be a valid year.", call. = FALSE)
  }

  list(
    min_loss = min_loss,
    start_year = as.integer(start_year),
    show_actual = show_actual,
    title = title,
    adjusted_subtitle = adjusted_subtitle,
    actual_subtitle = actual_subtitle
  )
}

xl_pricing_tool_riskfit_read_loss_layers <- function(sheet) {
  limit_header <- trimws(xl_pricing_tool_cell_value(sheet, "S61"))
  ded_header <- trimws(xl_pricing_tool_cell_value(sheet, "T61"))
  if (tolower(limit_header) != "limit" || tolower(ded_header) != "ded") {
    stop("Risk Loss Fitting layer structure must have Limit and Ded headers in S61:T61.", call. = FALSE)
  }

  rows <- list()
  blank_run <- 0
  for (row in 62:(62 + 200)) {
    limit_raw <- trimws(xl_pricing_tool_cell_value(sheet, paste0("S", row)))
    ded_raw <- trimws(xl_pricing_tool_cell_value(sheet, paste0("T", row)))
    if (!nzchar(limit_raw) && !nzchar(ded_raw)) {
      blank_run <- blank_run + 1
      if (blank_run >= 10) break
      next
    }
    blank_run <- 0

    limit <- xl_pricing_tool_riskfit_number(sheet, "S", row, "layer Limit", required = TRUE)
    ded <- xl_pricing_tool_riskfit_number(sheet, "T", row, "layer Ded", required = TRUE)
    if (limit <= 0 || ded < 0) {
      stop(sprintf("Risk Loss Fitting layer row %s must have positive Limit and nonnegative Ded.", row), call. = FALSE)
    }

    rows[[length(rows) + 1L]] <- data.frame(
      Ded = ded,
      Limit = limit,
      Lower = ded,
      Upper = ded + limit,
      stringsAsFactors = FALSE
    )
  }

  if (length(rows) == 0) {
    stop("Risk Loss Fitting layer structure S61:T... must contain at least one layer.", call. = FALSE)
  }

  out <- do.call(rbind, rows)
  out <- out[order(out$Lower, out$Upper), , drop = FALSE]
  out$FillColor <- rep(c("#F4F7FA", "#E9F1F7"), length.out = nrow(out))
  rownames(out) <- NULL
  out
}

xl_pricing_tool_riskfit_loss_comparison_data <- function(loss_input, chart_config) {
  base <- loss_input[
    loss_input$TreatyYear >= chart_config$start_year &
      loss_input$AsAt %in% c("Prior", "Current"),
    ,
    drop = FALSE
  ]

  rows <- list()
  adjusted <- base[base$Final_Incurred >= chart_config$min_loss, , drop = FALSE]
  if (nrow(adjusted) > 0) {
    rows[[length(rows) + 1L]] <- data.frame(
      ChartType = rep("Final_Incurred", nrow(adjusted)),
      AsAt = adjusted$AsAt,
      TreatyYear = adjusted$TreatyYear,
      ClaimID = adjusted$ClaimID,
      Insured = adjusted$Insured,
      Loss = adjusted$Final_Incurred,
      stringsAsFactors = FALSE
    )
  }

  if (chart_config$show_actual) {
    actual <- base[base$Actual_Incurred >= chart_config$min_loss, , drop = FALSE]
    if (nrow(actual) > 0) {
      rows[[length(rows) + 1L]] <- data.frame(
        ChartType = rep("Actual_Incurred", nrow(actual)),
        AsAt = actual$AsAt,
        TreatyYear = actual$TreatyYear,
        ClaimID = actual$ClaimID,
        Insured = actual$Insured,
        Loss = actual$Actual_Incurred,
        stringsAsFactors = FALSE
      )
    }
  }

  if (length(rows) == 0) {
    stop("Risk Loss Fitting loss comparison chart has no losses after applying StartYear/min filters.", call. = FALSE)
  }

  out <- do.call(rbind, rows)
  out$AsAt <- factor(out$AsAt, levels = c("Prior", "Current"))
  out
}

xl_pricing_tool_render_riskfit_loss_comparison_chart <- function(chart_data, layers, chart_config, subtitle, output_file) {
  if (!requireNamespace("ggplot2", quietly = TRUE)) {
    stop("Package 'ggplot2' is required to render RiskFit chart images.", call. = FALSE)
  }
  if (nrow(chart_data) == 0) {
    stop(sprintf("Risk Loss Fitting chart '%s' has no rows after filters.", subtitle), call. = FALSE)
  }

  ggplot2 <- asNamespace("ggplot2")
  y_max <- max(layers$Upper, na.rm = TRUE)
  chart_data <- xl_pricing_tool_riskfit_prepare_paired_losses(chart_data)
  bg_data <- xl_pricing_tool_riskfit_layer_background(chart_data, layers)

  plot <- ggplot2$ggplot(chart_data, ggplot2$aes(x = LossID, y = Loss, fill = AsAt))
  if (nrow(bg_data) > 0) {
    plot <- plot +
      ggplot2$geom_tile(
        data = bg_data,
        ggplot2$aes(x = LossID, y = LayerMid, height = LayerHeight, group = LayerID),
        inherit.aes = FALSE,
        fill = bg_data$FillColor,
        width = 1,
        alpha = 0.74
      )
  }

  plot <- plot +
    ggplot2$geom_col(
      position = ggplot2$position_dodge(width = 0.76),
      width = 0.68,
      alpha = 0.92
    ) +
    ggplot2$facet_grid(. ~ TreatyYear, scales = "free_x", space = "free_x") +
    ggplot2$scale_fill_manual(values = c(Prior = "#6F9BB3", Current = "#FF8B3D")) +
    ggplot2$scale_y_continuous(
      limits = c(0, y_max),
      labels = xl_pricing_tool_riskfit_amount_labels,
      expand = c(0, 0)
    ) +
    ggplot2$labs(
      title = chart_config$title,
      subtitle = subtitle,
      x = NULL,
      y = NULL,
      fill = NULL
    ) +
    ggplot2$theme_minimal(base_size = 9) +
    ggplot2$theme(
      plot.title = ggplot2$element_text(face = "bold", color = "#22313F", size = 13),
      plot.subtitle = ggplot2$element_text(color = "#52606D", size = 9),
      panel.grid.major.x = ggplot2$element_blank(),
      panel.grid.minor = ggplot2$element_blank(),
      panel.grid.major.y = ggplot2$element_line(color = "#FFFFFF", size = 0.3),
      axis.text = ggplot2$element_text(color = "#52606D"),
      axis.text.x = ggplot2$element_text(angle = 90, hjust = 1, vjust = 0.5),
      strip.text = ggplot2$element_text(face = "bold", color = "#22313F", hjust = 0),
      legend.position = "bottom",
      plot.margin = ggplot2$margin(8, 14, 8, 8)
    )

  ggplot2$ggsave(
    output_file,
    plot = plot,
    width = 9.8,
    height = 3.8,
    dpi = 160,
    bg = "white"
  )
}

xl_pricing_tool_riskfit_prepare_paired_losses <- function(chart_data) {
  chart_data$ClaimKey <- ifelse(nzchar(chart_data$ClaimID), chart_data$ClaimID, chart_data$Insured)
  chart_data$ClaimKey <- ifelse(nzchar(chart_data$ClaimKey), chart_data$ClaimKey, paste0("loss_", seq_len(nrow(chart_data))))
  chart_data$LabelBase <- ifelse(nzchar(chart_data$Insured), chart_data$Insured, chart_data$ClaimKey)

  aggregated <- stats::aggregate(
    Loss ~ TreatyYear + ClaimKey + LabelBase + AsAt,
    data = chart_data,
    FUN = max
  )
  order_data <- stats::aggregate(
    Loss ~ TreatyYear + ClaimKey + LabelBase,
    data = aggregated,
    FUN = max
  )
  order_data <- order_data[order(order_data$TreatyYear, order_data$Loss, order_data$ClaimKey), , drop = FALSE]
  order_data$XKey <- paste(order_data$TreatyYear, order_data$ClaimKey, sep = "__")
  order_data$DisplayLabel <- xl_pricing_tool_riskfit_truncate_labels(order_data$LabelBase)
  order_data$DisplayLabel <- make.unique(order_data$DisplayLabel, sep = " ")

  as_at_levels <- c("Prior", "Current")
  rows <- vector("list", nrow(order_data) * length(as_at_levels))
  k <- 0L
  for (i in seq_len(nrow(order_data))) {
    for (as_at in as_at_levels) {
      k <- k + 1L
      value <- aggregated$Loss[
        aggregated$TreatyYear == order_data$TreatyYear[[i]] &
          aggregated$ClaimKey == order_data$ClaimKey[[i]] &
          as.character(aggregated$AsAt) == as_at
      ]
      rows[[k]] <- data.frame(
        TreatyYear = order_data$TreatyYear[[i]],
        LossID = order_data$DisplayLabel[[i]],
        AsAt = as_at,
        Loss = if (length(value) == 0) 0 else max(value, na.rm = TRUE),
        stringsAsFactors = FALSE
      )
    }
  }

  out <- do.call(rbind, rows)
  out$TreatyYear <- factor(out$TreatyYear, levels = sort(unique(out$TreatyYear)))
  out$LossID <- factor(out$LossID, levels = order_data$DisplayLabel)
  out$AsAt <- factor(out$AsAt, levels = as_at_levels)
  out
}

xl_pricing_tool_riskfit_layer_background <- function(chart_data, layers) {
  x_keys <- unique(chart_data[, c("TreatyYear", "LossID"), drop = FALSE])
  if (nrow(x_keys) == 0 || nrow(layers) == 0) {
    return(data.frame())
  }

  rows <- vector("list", nrow(x_keys) * nrow(layers))
  k <- 0L
  for (i in seq_len(nrow(x_keys))) {
    for (j in seq_len(nrow(layers))) {
      k <- k + 1L
      rows[[k]] <- data.frame(
        TreatyYear = x_keys$TreatyYear[[i]],
        LossID = x_keys$LossID[[i]],
        LayerID = j,
        LayerMid = (layers$Lower[[j]] + layers$Upper[[j]]) / 2,
        LayerHeight = layers$Upper[[j]] - layers$Lower[[j]],
        FillColor = layers$FillColor[[j]],
        stringsAsFactors = FALSE
      )
    }
  }

  out <- do.call(rbind, rows)
  out$TreatyYear <- factor(out$TreatyYear, levels = levels(chart_data$TreatyYear))
  out$LossID <- factor(out$LossID, levels = levels(chart_data$LossID))
  out
}

xl_pricing_tool_riskfit_loss_labels <- function(chart_data) {
  labels <- ifelse(
    nzchar(chart_data$Insured),
    chart_data$Insured,
    chart_data$ClaimID
  )
  make.unique(xl_pricing_tool_riskfit_truncate_labels(labels), sep = " ")
}

xl_pricing_tool_riskfit_truncate_labels <- function(labels) {
  vapply(labels, function(label) {
    label <- trimws(label)
    if (nchar(label) <= 18) {
      return(label)
    }
    paste0(substr(label, 1, 16), "...")
  }, character(1))
}

xl_pricing_tool_riskfit_number <- function(sheet, col, row, field_name, required = TRUE) {
  raw_value <- trimws(xl_pricing_tool_cell_value(sheet, paste0(col, row)))
  if (!nzchar(raw_value)) {
    if (required) {
      stop(sprintf("Risk Loss Fitting %s cell %s%s is blank.", field_name, col, row), call. = FALSE)
    }
    return(0)
  }

  value <- suppressWarnings(as.numeric(gsub(",", "", raw_value, fixed = TRUE)))
  if (is.na(value)) {
    stop(sprintf("Risk Loss Fitting %s cell %s%s is not numeric: %s", field_name, col, row, raw_value), call. = FALSE)
  }
  value
}

xl_pricing_tool_riskfit_bool <- function(sheet, col, row, default = TRUE) {
  raw_value <- tolower(trimws(xl_pricing_tool_cell_value(sheet, paste0(col, row))))
  if (!nzchar(raw_value)) {
    return(default)
  }
  if (raw_value %in% c("1", "true", "yes", "y")) {
    return(TRUE)
  }
  if (raw_value %in% c("0", "false", "no", "n")) {
    return(FALSE)
  }
  stop(sprintf("Risk Loss Fitting boolean cell %s%s must be TRUE/FALSE or 1/0.", col, row), call. = FALSE)
}

xl_pricing_tool_riskfit_normalize_asat <- function(value, row) {
  key <- tolower(trimws(value))
  if (key %in% c("prior", "p")) {
    return("Prior")
  }
  if (key %in% c("current", "curr", "c")) {
    return("Current")
  }
  stop(sprintf("Risk Loss Fitting row %s has unsupported AsAt '%s'. Supported values: Prior, Current.", row, value), call. = FALSE)
}

xl_pricing_tool_riskfit_amount_labels <- function(values) {
  vapply(values, function(value) {
    if (is.na(value)) {
      return("")
    }
    if (abs(value) >= 1e9) {
      return(sprintf("%sB", format(round(value / 1e9, 1), trim = TRUE, nsmall = 0)))
    }
    sprintf("%sM", format(round(value / 1e6, 0), trim = TRUE, big.mark = ",", scientific = FALSE))
  }, character(1))
}
