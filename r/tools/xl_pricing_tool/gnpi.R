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
  start_year <- xl_pricing_tool_gnpi_start_year(sheet)
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
  composition_data <- xl_pricing_tool_gnpi_composition_data(gnpi, start_year, renewal_year, lob_sequence)
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
          SortKey = xl_pricing_tool_gnpi_level_sort_key(year, type, renewal_year),
          stringsAsFactors = FALSE
        )
      }
    }
  }

  if (length(rows) == 0) {
    return(data.frame())
  }
  out <- do.call(rbind, rows)
  out[order(out$SortKey), , drop = FALSE]
}

xl_pricing_tool_gnpi_composition_data <- function(gnpi, start_year, renewal_year, lob_sequence = character()) {
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
          SortKey = xl_pricing_tool_gnpi_level_sort_key(year, type, renewal_year),
          stringsAsFactors = FALSE
        )
      }
    }
  }

  if (length(rows) == 0) {
    return(data.frame())
  }
  out <- do.call(rbind, rows)
  out[order(out$SortKey), , drop = FALSE]
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
      "Chart 1 Title",
      "Chart 1 Subtitle",
      "Chart 2 Title",
      "Chart 2 Subtitle",
      "Chart Width",
      "Chart Height",
      "Font Scale",
      "Label Font Size",
      "Axis Font Size",
      "Title Font Size",
      "LOB Palette"
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
      chart_options$level_title,
      chart_options$level_subtitle,
      chart_options$composition_title,
      chart_options$composition_subtitle,
      chart_options$width,
      chart_options$height,
      chart_options$font_scale,
      chart_options$label_font_size,
      chart_options$axis_font_size,
      chart_options$title_font_size,
      paste(chart_options$lob_colors, collapse = ",")
    )
  )
}

xl_pricing_tool_gnpi_start_year <- function(sheet) {
  value <- suppressWarnings(as.integer(xl_pricing_tool_cell_value(sheet, "L18")))
  if (!is.na(value)) {
    return(value)
  }

  value <- suppressWarnings(as.integer(xl_pricing_tool_cell_value(sheet, "L19")))
  if (!is.na(value)) {
    return(value)
  }

  value <- suppressWarnings(as.integer(xl_pricing_tool_find_setting_value(sheet, c("start year", "start_year"))))
  if (!is.na(value)) value else NA_integer_
}

xl_pricing_tool_gnpi_chart_options <- function(sheet) {
  option <- function(cell, default, labels = character()) {
    value <- trimws(xl_pricing_tool_cell_value(sheet, cell))
    if (!nzchar(value) && length(labels) > 0) {
      value <- xl_pricing_tool_find_setting_value(sheet, labels, search_cols = 27:35, search_rows = 10:80)
    }
    if (!nzchar(value) && length(labels) > 0) {
      value <- xl_pricing_tool_find_setting_value(sheet, labels)
    }
    if (nzchar(value)) value else default
  }
  number_option <- function(cell, default, labels = character()) {
    value <- suppressWarnings(as.numeric(option(cell, "", labels)))
    if (is.na(value) || value <= 0) default else value
  }

  list(
    level_title = option("L20", "GNPI by Treaty Year and Type", c("level chart title", "level plot title", "chart 1 title")),
    level_subtitle = option("L21", "Total GNPI by year and type", c("level chart subtitle", "level plot subtitle", "chart 1 subtitle")),
    composition_title = option("L22", "LOB Composition by Treaty Year", c("composition chart title", "stack chart title", "chart 2 title")),
    composition_subtitle = option("L23", "Stacked percentage composition by year/type section from configured start year", c("composition chart subtitle", "stack chart subtitle", "chart 2 subtitle")),
    actual_color = xl_pricing_tool_valid_color(option("L24", "#4F9FBC", c("actual color", "actual colour")), "#4F9FBC"),
    revised_color = xl_pricing_tool_valid_color(option("L25", "#3A7F72", c("revised color", "revised colour")), "#3A7F72"),
    estimate_color = xl_pricing_tool_valid_color(option("L26", "#8EC5D6", c("estimate color", "estimate colour")), "#8EC5D6"),
    label_color = xl_pricing_tool_valid_color(option("L27", "#1F2933", c("label color", "label colour", "font color", "font colour")), "#1F2933"),
    grid_color = xl_pricing_tool_valid_color(option("L28", "#E5E7EB", c("grid color", "grid colour")), "#E5E7EB"),
    width = number_option("L29", 1216, c("png width", "chart width")),
    height = number_option("L30", 624, c("png height", "chart height")),
    font_scale = number_option("L31", 1, c("font scale")),
    label_font_size = number_option("L32", 3.3, c("label font size", "data label font size")),
    axis_font_size = number_option("L33", 9, c("axis font size")),
    title_font_size = number_option("L34", 16, c("title font size")),
    lob_colors = xl_pricing_tool_gnpi_lob_palette(sheet)
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
  if (!requireNamespace("ggplot2", quietly = TRUE)) {
    stop("Package 'ggplot2' is required to render GNPI chart images.", call. = FALSE)
  }

  ggplot2 <- asNamespace("ggplot2")
  type_levels <- c("Actual", "Revised", "Estimate")
  y_max <- max(level_data$GNPI, na.rm = TRUE)
  if (!is.finite(y_max) || y_max <= 0) {
    y_max <- 1
  }
  scale <- xl_pricing_tool_gnpi_chart_scale(y_max)
  level_data$GNPI_Scaled <- level_data$GNPI / scale
  level_data$Label <- factor(level_data$Label, levels = unique(level_data$Label[order(level_data$SortKey)]))
  level_data$Type <- factor(level_data$Type, levels = type_levels)
  level_data$ValueLabel <- xl_pricing_tool_format_amount(level_data$GNPI_Scaled)
  colors <- c(
    Actual = chart_options$actual_color,
    Revised = chart_options$revised_color,
    Estimate = chart_options$estimate_color
  )
  level_data$FillColor <- unname(colors[as.character(level_data$Type)])
  level_data$FillColor[is.na(level_data$FillColor) | !nzchar(level_data$FillColor)] <- chart_options$estimate_color

  plot <- ggplot2$ggplot(level_data, ggplot2$aes_string(x = "Label", y = "GNPI_Scaled", fill = "FillColor")) +
    ggplot2$geom_bar(stat = "identity", width = 0.72, color = "white") +
    ggplot2$geom_text(
      ggplot2$aes_string(label = "ValueLabel"),
      vjust = -0.35,
      size = chart_options$label_font_size,
      color = chart_options$label_color
    ) +
    ggplot2$scale_fill_identity(guide = "legend", breaks = unname(colors), labels = names(colors)) +
    ggplot2$scale_y_continuous(
      labels = function(x) paste0(format(round(x, 0), big.mark = ",", scientific = FALSE, trim = TRUE), "M"),
      limits = c(0, max(level_data$GNPI_Scaled, na.rm = TRUE) * 1.22),
      expand = c(0, 0)
    ) +
    ggplot2$labs(
      title = chart_options$level_title,
      subtitle = chart_options$level_subtitle,
      x = NULL,
      y = NULL,
      fill = NULL
    ) +
    xl_pricing_tool_gnpi_theme(chart_options) +
    ggplot2$theme(axis.text.x = ggplot2$element_text(angle = 90, vjust = 0.5, hjust = 1))

  ggplot2$ggsave(output_file, plot = plot, width = chart_options$width / 160, height = chart_options$height / 160, dpi = 160, bg = "white")
}

xl_pricing_tool_render_gnpi_comparison_chart <- function(composition_data, output_file, chart_options) {
  if (nrow(composition_data) == 0) {
    stop("No GNPI LOB comparison data is available for chart rendering.", call. = FALSE)
  }
  if (!requireNamespace("ggplot2", quietly = TRUE)) {
    stop("Package 'ggplot2' is required to render GNPI chart images.", call. = FALSE)
  }

  ggplot2 <- asNamespace("ggplot2")
  composition_data$Label <- factor(composition_data$Label, levels = unique(composition_data$Label[order(composition_data$SortKey)]))
  composition_data$LOB <- factor(composition_data$LOB, levels = unique(composition_data$LOB))
  lob_colors <- xl_pricing_tool_lob_colors(nlevels(composition_data$LOB), chart_options$lob_colors)
  names(lob_colors) <- levels(composition_data$LOB)
  composition_data$FillColor <- unname(lob_colors[as.character(composition_data$LOB)])

  plot <- ggplot2$ggplot(composition_data, ggplot2$aes_string(x = "Label", y = "Percent", fill = "FillColor")) +
    ggplot2$geom_bar(stat = "identity", width = 0.72, color = "white") +
    ggplot2$scale_fill_identity(guide = "legend", breaks = unname(lob_colors), labels = names(lob_colors)) +
    ggplot2$scale_y_continuous(
      labels = function(x) paste0(round(x * 100), "%"),
      limits = c(0, 1),
      expand = c(0, 0)
    ) +
    ggplot2$labs(
      title = chart_options$composition_title,
      subtitle = chart_options$composition_subtitle,
      x = NULL,
      y = NULL,
      fill = NULL
    ) +
    xl_pricing_tool_gnpi_theme(chart_options)

  ggplot2$ggsave(output_file, plot = plot, width = chart_options$width / 160, height = chart_options$height / 160, dpi = 160, bg = "white")
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

xl_pricing_tool_lob_colors <- function(n, palette = character()) {
  if (length(palette) > 0) {
    palette <- vapply(palette[nzchar(palette)], xl_pricing_tool_valid_color, character(1), fallback = NA_character_)
    palette <- palette[!is.na(palette) & nzchar(palette)]
  }
  if (length(palette) == 0) {
    palette <- c("#A7C7DC", "#3F7FBF", "#7EAA92", "#D9A441", "#C96C5A", "#8E7CC3", "#6FA8DC", "#93C47D")
  }
  rep(palette, length.out = n)
}

xl_pricing_tool_gnpi_lob_palette <- function(sheet) {
  colors <- character()
  for (i in 1:12) {
    value <- xl_pricing_tool_find_setting_value(
      sheet,
      c(sprintf("lob color %s", i), sprintf("lob colour %s", i)),
      search_cols = 27:35,
      search_rows = 10:80
    )
    if (nzchar(value)) {
      colors <- c(colors, xl_pricing_tool_valid_color(value, ""))
    }
  }
  colors
}

xl_pricing_tool_valid_color <- function(value, fallback) {
  value <- trimws(as.character(value))
  value <- gsub("^['\"]|['\"]$", "", value)
  if (grepl("^[0-9A-Fa-f]{6}$", value)) {
    value <- paste0("#", value)
  }
  ok <- tryCatch({
    grDevices::col2rgb(value)
    TRUE
  }, error = function(e) FALSE)
  if (ok) value else fallback
}

xl_pricing_tool_gnpi_theme <- function(chart_options) {
  ggplot2 <- asNamespace("ggplot2")
  ggplot2$theme_minimal(base_size = chart_options$axis_font_size * chart_options$font_scale) +
    ggplot2$theme(
      plot.title = ggplot2$element_text(face = "bold", color = chart_options$label_color, size = chart_options$title_font_size * chart_options$font_scale),
      plot.subtitle = ggplot2$element_text(color = "#52606D", size = chart_options$axis_font_size * chart_options$font_scale),
      panel.grid.major.x = ggplot2$element_blank(),
      panel.grid.major.y = ggplot2$element_line(color = chart_options$grid_color),
      panel.grid.minor = ggplot2$element_blank(),
      axis.text = ggplot2$element_text(color = "#52606D", size = chart_options$axis_font_size * chart_options$font_scale),
      legend.position = "top",
      legend.justification = "left",
      legend.text = ggplot2$element_text(color = chart_options$label_color, size = chart_options$axis_font_size * chart_options$font_scale),
      plot.margin = grid::unit(c(12, 16, 12, 16), "pt")
    )
}

xl_pricing_tool_find_setting_value <- function(sheet, labels, search_cols = 10:26, search_rows = 1:120) {
  normalized_labels <- xl_pricing_tool_normalize_setting_label(labels)
  for (row in search_rows) {
    for (col in search_cols) {
      ref <- paste0(xl_pricing_tool_col_name(col), row)
      value <- xl_pricing_tool_cell_value(sheet, ref)
      if (!nzchar(value)) {
        next
      }
      if (xl_pricing_tool_normalize_setting_label(value) %in% normalized_labels) {
        for (offset in 1:4) {
          next_value <- trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col + offset), row)))
          if (nzchar(next_value)) {
            return(next_value)
          }
        }
      }
    }
  }
  ""
}

xl_pricing_tool_normalize_setting_label <- function(value) {
  value <- tolower(trimws(as.character(value)))
  value <- gsub("[_\\-]+", " ", value)
  value <- gsub("\\s+", " ", value)
  value
}

xl_pricing_tool_gnpi_level_sort_key <- function(year, type, renewal_year) {
  prior_year <- renewal_year - 1
  if (year < prior_year) {
    return(year * 10 + match(type, c("Actual", "Revised", "Estimate"), nomatch = 9))
  }
  if (year == prior_year && identical(type, "Estimate")) {
    return(prior_year * 10 + 1)
  }
  if (year == prior_year && type %in% c("Revised", "Actual")) {
    return(prior_year * 10 + 2 + match(type, c("Revised", "Actual"), nomatch = 2) / 10)
  }
  if (year == renewal_year && identical(type, "Estimate")) {
    return(renewal_year * 10 + 1)
  }
  year * 10 + 5 + match(type, c("Actual", "Revised", "Estimate"), nomatch = 9) / 10
}

xl_pricing_tool_renewal_year <- function(workbook_path) {
  sheet <- xl_pricing_tool_read_sheet(workbook_path, "Control_Module")
  value <- suppressWarnings(as.integer(xl_pricing_tool_cell_value(sheet, "C13")))
  if (is.na(value)) {
    stop("Unable to determine renewal year from Control_Module!C13.", call. = FALSE)
  }
  value
}
