xl_pricing_tool_gnpi_update <- function(workbook_path, output_dir = NULL, context = list()) {
  xl_pricing_tool_validate(workbook_path, context)

  if (is.null(output_dir) || identical(output_dir, "")) {
    output_dir <- if (!is.null(context$output_dir)) {
      context$output_dir
    } else {
      btk_default_run_dir("xl_pricing_tool", workbook_path)
    }
  }

  gnpi_dir <- file.path(output_dir, "gnpi")
  dir.create(gnpi_dir, recursive = TRUE, showWarnings = FALSE)

  sheet <- xl_pricing_tool_read_sheet(workbook_path, "Input_GNPI")
  marker <- xl_pricing_tool_cell_value(sheet, "A1")
  gather_from <- xl_pricing_tool_cell_value(sheet, "C6")
  start_year <- suppressWarnings(as.integer(xl_pricing_tool_cell_value(sheet, "L19")))
  if (is.na(start_year)) {
    start_year <- min(suppressWarnings(as.integer(xl_pricing_tool_column_values(sheet, "C", 13))), na.rm = TRUE)
  }
  chart_options <- xl_pricing_tool_gnpi_chart_options(sheet)

  lob_sequence <- xl_pricing_tool_read_vertical_values(sheet, "L", 39)
  if (length(lob_sequence) == 0) {
    lob_sequence <- xl_pricing_tool_read_vertical_values(sheet, "L", 36)
  }
  if (length(lob_sequence) == 0) {
    lob_sequence <- xl_pricing_tool_read_vertical_values(sheet, "L", 41)
  }
  gnpi <- xl_pricing_tool_read_table(
    sheet = sheet,
    header_row = 12,
    start_col = 2,
    end_col = 9,
    first_data_row = 13,
    key_cols = c("LOB", "TreatyYear", "Type", "GNPI_Original", "Adjustment")
  )

  xl_pricing_tool_validate_gnpi_table(gnpi)

  gnpi$TreatyYear <- suppressWarnings(as.integer(gnpi$TreatyYear))
  for (col in c("GNPI_Original", "Adjustment", "GNPI", "Latest_Current", "Latest_Prior")) {
    gnpi[[col]] <- suppressWarnings(as.numeric(gnpi[[col]]))
  }

  gnpi$GNPI_Calculated <- ifelse(
    is.na(gnpi$GNPI),
    gnpi$GNPI_Original * ifelse(is.na(gnpi$Adjustment), 1, gnpi$Adjustment),
    gnpi$GNPI
  )

  renewal_year <- xl_pricing_tool_renewal_year(workbook_path)
  level_data <- xl_pricing_tool_gnpi_level_data(gnpi, start_year, renewal_year)
  composition_data <- xl_pricing_tool_gnpi_composition_data(gnpi, start_year, lob_sequence)
  config_data <- xl_pricing_tool_gnpi_config_data(
    marker = marker,
    gather_from = gather_from,
    start_year = start_year,
    renewal_year = renewal_year,
    chart1_mode = "clustered",
    chart2_mode = "stacked_percent",
    chart_options = chart_options
  )

  level_file <- file.path(gnpi_dir, "gnpi_level_data.csv")
  composition_file <- file.path(gnpi_dir, "gnpi_lob_comparison_data.csv")
  config_file <- file.path(gnpi_dir, "gnpi_chart_config.csv")
  level_chart_file <- file.path(gnpi_dir, "gnpi_level_chart.png")
  composition_chart_file <- file.path(gnpi_dir, "gnpi_lob_comparison_chart.png")

  utils::write.csv(level_data, level_file, row.names = FALSE, na = "")
  utils::write.csv(composition_data, composition_file, row.names = FALSE, na = "")
  utils::write.csv(config_data, config_file, row.names = FALSE, na = "")
  xl_pricing_tool_render_gnpi_charts(
    level_data = level_data,
    composition_data = composition_data,
    level_chart_file = level_chart_file,
    composition_chart_file = composition_chart_file,
    start_year = start_year,
    renewal_year = renewal_year,
    chart_options = chart_options
  )

  paste(
    "GNPI update completed.",
    sprintf("Output folder: %s", output_dir),
    sprintf("Input rows: %s", nrow(gnpi)),
    sprintf("Start year: %s", start_year),
    sprintf("Renewal year: %s", renewal_year),
    sprintf("Chart 1 data: %s", level_file),
    sprintf("Chart 2 data: %s", composition_file),
    sprintf("Chart config: %s", config_file),
    sprintf("Chart 1 image: %s", level_chart_file),
    sprintf("Chart 2 image: %s", composition_chart_file),
    "R chart images are handled by the VBA client after BERT returns.",
    sep = vb_newline()
  )
}

xl_pricing_tool_gnpi_gather <- function(workbook_path, output_dir = NULL, context = list()) {
  "No gather action defined for Input_GNPI because Gather From is blank."
}

xl_pricing_tool_gnpi_build <- function(workbook_path, output_dir = NULL, context = list()) {
  "No build output defined for Input_GNPI."
}

xl_pricing_tool_validate_gnpi_table <- function(gnpi) {
  required <- c("LOB", "TreatyYear", "Type", "GNPI_Original", "Adjustment", "GNPI", "Latest_Current", "Latest_Prior")
  missing <- setdiff(required, names(gnpi))
  if (length(missing) > 0) {
    stop(sprintf("Input_GNPI is missing required column(s): %s", paste(missing, collapse = ", ")), call. = FALSE)
  }

  valid_types <- c("Actual", "Revised", "Estimate")
  bad_types <- unique(gnpi$Type[!is.na(gnpi$Type) & !gnpi$Type %in% valid_types])
  if (length(bad_types) > 0) {
    stop(sprintf("Input_GNPI has unsupported Type value(s): %s", paste(bad_types, collapse = ", ")), call. = FALSE)
  }

  invisible(TRUE)
}

xl_pricing_tool_gnpi_level_data <- function(gnpi, start_year, renewal_year) {
  type_levels <- c("Actual", "Revised", "Estimate")
  years <- sort(unique(gnpi$TreatyYear[is.finite(gnpi$TreatyYear) & gnpi$TreatyYear >= start_year]))
  if (length(years) == 0) {
    return(data.frame())
  }

  rows <- list()
  for (year in years) {
    for (type in type_levels) {
      value <- sum(gnpi$GNPI_Calculated[gnpi$TreatyYear == year & gnpi$Type == type], na.rm = TRUE)
      if (value > 0) {
        rows[[length(rows) + 1L]] <- data.frame(
          TreatyYear = year,
          Type = type,
          Label = paste(year, type, sep = "\n"),
          GNPI = value,
          Section = ifelse(year == renewal_year, "Latest Current", ifelse(year == renewal_year - 1, "Latest Prior", "History")),
          stringsAsFactors = FALSE
        )
      }
    }
  }

  if (length(rows) == 0) {
    return(data.frame())
  }
  do.call(rbind, rows)
}

xl_pricing_tool_gnpi_composition_data <- function(gnpi, start_year, lob_sequence = character()) {
  type_levels <- c("Actual", "Revised", "Estimate")
  lobs <- unique(gnpi$LOB[!is.na(gnpi$LOB) & nzchar(gnpi$LOB)])
  if (length(lob_sequence) > 0) {
    lobs <- c(lob_sequence[lob_sequence %in% lobs], setdiff(lobs, lob_sequence))
  }

  years <- sort(unique(gnpi$TreatyYear[is.finite(gnpi$TreatyYear) & gnpi$TreatyYear >= start_year]))
  rows <- list()
  for (year in years) {
    for (type in type_levels) {
      section_has_data <- any(gnpi$TreatyYear == year & gnpi$Type == type & gnpi$GNPI_Calculated > 0, na.rm = TRUE)
      if (!section_has_data) {
        next
      }
      values <- vapply(lobs, function(lob) {
        sum(gnpi$GNPI_Calculated[gnpi$TreatyYear == year & gnpi$Type == type & gnpi$LOB == lob], na.rm = TRUE)
      }, numeric(1))
      total <- sum(values, na.rm = TRUE)
      for (i in seq_along(lobs)) {
        rows[[length(rows) + 1L]] <- data.frame(
          TreatyYear = year,
          Type = type,
          Label = paste(year, type, sep = "\n"),
          LOB = lobs[[i]],
          GNPI = values[[i]],
          Percent = ifelse(total == 0, NA_real_, values[[i]] / total),
          stringsAsFactors = FALSE
        )
      }
    }
  }

  if (length(rows) == 0) {
    return(data.frame())
  }
  do.call(rbind, rows)
}

xl_pricing_tool_gnpi_config_data <- function(marker, gather_from, start_year, renewal_year, chart1_mode, chart2_mode, chart_options) {
  data.frame(
    Key = c(
      "Marker",
      "Gather From",
      "Start Year",
      "Renewal Year",
      "Chart 1 Mode",
      "Chart 2 Mode",
      "Actual Color",
      "Revised Color",
      "Estimate Color",
      "Latest Prior Color",
      "Latest Current Color",
      "Chart 1 Title",
      "Chart 2 Title",
      "Chart Width",
      "Chart Height",
      "Font Scale"
    ),
    Value = c(
      marker,
      gather_from,
      start_year,
      renewal_year,
      chart1_mode,
      chart2_mode,
      chart_options$actual_color,
      chart_options$revised_color,
      chart_options$estimate_color,
      "#64748B",
      "#166534",
      chart_options$level_title,
      chart_options$composition_title,
      chart_options$width,
      chart_options$height,
      chart_options$font_scale
    )
  )
}

xl_pricing_tool_gnpi_chart_options <- function(sheet) {
  option <- function(cell, default) {
    value <- trimws(xl_pricing_tool_cell_value(sheet, cell))
    if (nzchar(value)) value else default
  }
  number_option <- function(cell, default) {
    value <- suppressWarnings(as.numeric(option(cell, "")))
    if (is.na(value) || value <= 0) default else value
  }

  list(
    level_title = option("L20", "GNPI by Treaty Year and Type"),
    composition_title = option("L21", "LOB Composition by Treaty Year"),
    actual_color = option("L22", "#4F9FBC"),
    revised_color = option("L23", "#3A7F72"),
    estimate_color = option("L24", "#8EC5D6"),
    label_color = option("L25", "#1F2933"),
    grid_color = option("L26", "#E5E7EB"),
    width = number_option("L27", 1216),
    height = number_option("L28", 624),
    font_scale = number_option("L29", 1)
  )
}

xl_pricing_tool_render_gnpi_charts <- function(level_data, composition_data, level_chart_file, composition_chart_file, start_year, renewal_year, chart_options) {
  xl_pricing_tool_render_gnpi_level_chart(
    level_data = level_data,
    output_file = level_chart_file,
    start_year = start_year,
    renewal_year = renewal_year,
    chart_options = chart_options
  )
  xl_pricing_tool_render_gnpi_comparison_chart(
    composition_data = composition_data,
    output_file = composition_chart_file,
    chart_options = chart_options
  )
}

xl_pricing_tool_render_gnpi_level_chart <- function(level_data, output_file, start_year, renewal_year, chart_options) {
  if (nrow(level_data) == 0) {
    stop("No GNPI level data is available for chart rendering.", call. = FALSE)
  }

  type_levels <- c("Actual", "Revised", "Estimate")
  values <- level_data$GNPI
  y_max <- max(values, na.rm = TRUE)
  if (!is.finite(y_max) || y_max <= 0) {
    y_max <- 1
  }
  scale <- xl_pricing_tool_gnpi_chart_scale(y_max)
  values <- values / scale
  y_max <- y_max / scale
  colors <- c(
    Actual = chart_options$actual_color,
    Revised = chart_options$revised_color,
    Estimate = chart_options$estimate_color
  )
  bar_colors <- colors[level_data$Type]
  bar_colors[is.na(bar_colors)] <- "#8EC5D6"

  grDevices::png(output_file, width = chart_options$width, height = chart_options$height, res = 160, bg = "white")
  old_par <- graphics::par(no.readonly = TRUE)
  on.exit({
    graphics::par(old_par)
    grDevices::dev.off()
  }, add = TRUE)

  graphics::par(mar = c(6.7, 5.0, 4.6, 1.8), family = "sans", cex = chart_options$font_scale)
  mids <- graphics::barplot(
    values,
    col = bar_colors,
    border = "white",
    ylim = c(0, y_max * 1.45),
    names.arg = level_data$Label,
    ylab = "",
    las = 2,
    cex.names = 0.68,
    axes = FALSE
  )
  y_ticks <- graphics::axTicks(2)
  graphics::axis(2, at = y_ticks, las = 1, col.axis = "#52606D", labels = paste0(format(round(y_ticks, 0), big.mark = ",", scientific = FALSE, trim = TRUE), "M"))
  graphics::grid(nx = NA, ny = NULL, col = chart_options$grid_color, lty = 1)
  graphics::box(col = "#CBD2D9")
  graphics::title(
    main = chart_options$level_title,
    sub = sprintf("Each bar is one year/type section from %s; latest prior/current years outlined", start_year),
    col.main = chart_options$label_color,
    col.sub = "#52606D",
    font.main = 2,
    cex.main = 1.25,
    cex.sub = 0.82
  )
  graphics::legend(
    "topleft",
    legend = type_levels,
    fill = colors[type_levels],
    border = NA,
    bty = "n",
    horiz = TRUE,
    cex = 0.82
  )
  graphics::text(mids, values, labels = xl_pricing_tool_format_amount(values), pos = 3, cex = 0.55, col = chart_options$label_color)

  prior_idx <- which(level_data$TreatyYear == renewal_year - 1)
  current_idx <- which(level_data$TreatyYear == renewal_year)
  if (length(prior_idx) > 0) {
    graphics::rect(mids[prior_idx] - 0.46, 0, mids[prior_idx] + 0.46, values[prior_idx] * 1.12, border = "#64748B", lwd = 2)
  }
  if (length(current_idx) > 0) {
    graphics::rect(mids[current_idx] - 0.46, 0, mids[current_idx] + 0.46, values[current_idx] * 1.12, border = "#166534", lwd = 2)
  }
}

xl_pricing_tool_render_gnpi_comparison_chart <- function(composition_data, output_file, chart_options) {
  if (nrow(composition_data) == 0) {
    stop("No GNPI LOB comparison data is available for chart rendering.", call. = FALSE)
  }

  sections <- unique(composition_data$Label)
  lobs <- unique(composition_data$LOB)
  chart_matrix <- vapply(sections, function(section) {
    values <- composition_data$Percent[composition_data$Label == section]
    names(values) <- composition_data$LOB[composition_data$Label == section]
    values[lobs]
  }, numeric(length(lobs)))
  rownames(chart_matrix) <- lobs
  chart_matrix[is.na(chart_matrix)] <- 0
  lob_colors <- xl_pricing_tool_lob_colors(length(lobs))

  grDevices::png(output_file, width = chart_options$width, height = chart_options$height, res = 160, bg = "white")
  old_par <- graphics::par(no.readonly = TRUE)
  on.exit({
    graphics::par(old_par)
    grDevices::dev.off()
  }, add = TRUE)

  graphics::par(mar = c(5.2, 5.0, 4.4, 1.8), family = "sans", cex = chart_options$font_scale)
  mids <- graphics::barplot(
    chart_matrix,
    beside = FALSE,
    col = lob_colors,
    border = "white",
    ylim = c(0, 1.15),
    names.arg = sections,
    ylab = "",
    las = 1,
    cex.names = 0.86,
    axes = FALSE
  )
  y_ticks <- seq(0, 1, by = 0.25)
  graphics::axis(2, at = y_ticks, las = 1, col.axis = "#52606D", labels = paste0(round(y_ticks * 100), "%"))
  graphics::grid(nx = NA, ny = NULL, col = chart_options$grid_color, lty = 1)
  graphics::box(col = "#CBD2D9")
  graphics::title(
    main = chart_options$composition_title,
    sub = "Stacked percentage composition by year/type section from configured start year",
    col.main = chart_options$label_color,
    col.sub = "#52606D",
    font.main = 2,
    cex.main = 1.25,
    cex.sub = 0.82
  )
  graphics::legend(
    "topleft",
    legend = rownames(chart_matrix),
    fill = lob_colors,
    border = NA,
    bty = "n",
    horiz = TRUE,
    cex = 0.82
  )
}

xl_pricing_tool_format_amount <- function(value) {
  ifelse(
    is.na(value),
    "",
    paste0(format(round(value, 1), big.mark = ",", trim = TRUE), "M")
  )
}

xl_pricing_tool_gnpi_chart_scale <- function(max_value) {
  if (is.finite(max_value) && max_value > 100000) {
    1000000
  } else {
    1
  }
}

xl_pricing_tool_lob_colors <- function(n) {
  palette <- c("#A7C7DC", "#3F7FBF", "#7EAA92", "#D9A441", "#C96C5A", "#8E7CC3", "#6FA8DC", "#93C47D")
  rep(palette, length.out = n)
}

xl_pricing_tool_renewal_year <- function(workbook_path) {
  sheet <- xl_pricing_tool_read_sheet(workbook_path, "Control_Module")
  value <- suppressWarnings(as.integer(xl_pricing_tool_cell_value(sheet, "C13")))
  if (is.na(value)) {
    stop("Unable to determine renewal year from Control_Module!C13.", call. = FALSE)
  }
  value
}
