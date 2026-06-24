xl_pricing_tool_relativepricing_gather <- function(workbook_path, output_dir = NULL, context = list()) {
  "No gather action defined for RelativePricing because workbook formulas own the current workflow."
}

xl_pricing_tool_relativepricing_update <- function(workbook_path, output_dir = NULL, context = list()) {
  xl_pricing_tool_validate(workbook_path, context, action = "relativepricing_update")

  if (is.null(output_dir) || identical(output_dir, "")) {
    output_dir <- if (!is.null(context$output_dir)) {
      context$output_dir
    } else {
      btk_default_run_dir("xl_pricing_tool", workbook_path)
    }
  }

  relativepricing_dir <- file.path(output_dir, "relativepricing")
  dir.create(relativepricing_dir, recursive = TRUE, showWarnings = FALSE)

  sheet <- xl_pricing_tool_action_sheet(workbook_path, context, "RelativePrice")
  input_layers <- xl_pricing_tool_read_sheet(workbook_path, "Input_Layers")
  checks <- c(
    xl_pricing_tool_relativepricing_check_marker(sheet),
    xl_pricing_tool_relativepricing_check_layout(sheet),
    xl_pricing_tool_relativepricing_check_input_layers(input_layers),
    xl_pricing_tool_relativepricing_check_names(workbook_path)
  )
  fitting_input <- xl_pricing_tool_relativepricing_read_fitting_input(sheet)
  fits <- xl_pricing_tool_relativepricing_fit_curves(fitting_input)

  parameters_file <- file.path(relativepricing_dir, "relativepricing_parameters.csv")
  if (nrow(fits$parameters) > 0) {
    utils::write.table(
      fits$parameters[, c("CurveCode", "a", "b"), drop = FALSE],
      file = parameters_file,
      sep = ",",
      row.names = FALSE,
      col.names = FALSE,
      na = "",
      quote = TRUE
    )
  } else {
    file.create(parameters_file)
  }

  plot_files <- if (nrow(fits$parameters) > 0) {
    xl_pricing_tool_relativepricing_render_plots(
      fitting_input = fitting_input,
      parameters = fits$parameters,
      output_dir = relativepricing_dir
    )
  } else {
    character()
  }

  paste(
    "RelativePricing update completed.",
    sprintf("Active sheet: %s", sheet$name),
    paste(checks, collapse = vb_newline()),
    sprintf("Input rows: %s", nrow(fitting_input)),
    sprintf("Fitted curve count: %s", nrow(fits$parameters)),
    if (length(fits$warnings) > 0) sprintf("Warnings: %s", paste(fits$warnings, collapse = "; ")) else "Warnings: none",
    sprintf("Output folder: %s", output_dir),
    sprintf("Parameter output: %s", parameters_file),
    sprintf("Plot images: %s", if (length(plot_files) > 0) paste(basename(plot_files), collapse = ", ") else "none"),
    "CSV output and plot images are handled by the VBA client after BERT returns.",
    sep = vb_newline()
  )
}

xl_pricing_tool_relativepricing_build <- function(workbook_path, output_dir = NULL, context = list()) {
  "No build output defined for RelativePricing."
}

xl_pricing_tool_relativepricing_read_fitting_input <- function(sheet) {
  rows <- list()
  for (row in 14:28) {
    curve_code <- trimws(xl_pricing_tool_cell_value(sheet, paste0("M", row)))
    rol <- suppressWarnings(as.numeric(trimws(xl_pricing_tool_cell_value(sheet, paste0("J", row)))))
    adjusted_layer_mp <- suppressWarnings(as.numeric(trimws(xl_pricing_tool_cell_value(sheet, paste0("L", row)))))
    layer <- trimws(xl_pricing_tool_cell_value(sheet, paste0("F", row)))

    if (!nzchar(curve_code) && !nzchar(layer) && is.na(rol) && is.na(adjusted_layer_mp)) {
      next
    }

    rows[[length(rows) + 1L]] <- data.frame(
      Row = row,
      Layer = layer,
      ROL = rol,
      AdjustedLayerMP = adjusted_layer_mp,
      CurveCode = curve_code,
      stringsAsFactors = FALSE
    )
  }

  if (length(rows) == 0) {
    stop("RelativePricing fitting base has no input rows in F14:O28.", call. = FALSE)
  }

  out <- do.call(rbind, rows)
  rownames(out) <- NULL
  out
}

xl_pricing_tool_relativepricing_fit_curves <- function(fitting_input) {
  curve_codes <- unique(fitting_input$CurveCode[nzchar(fitting_input$CurveCode)])
  curve_codes <- curve_codes[order(curve_codes)]
  parameter_rows <- list()
  warnings <- character()

  for (curve_code in curve_codes) {
    curve_rows <- fitting_input[
      fitting_input$CurveCode == curve_code &
        is.finite(fitting_input$ROL) &
        is.finite(fitting_input$AdjustedLayerMP) &
        fitting_input$ROL > 0 &
        fitting_input$AdjustedLayerMP > 0,
      ,
      drop = FALSE
    ]

    if (nrow(curve_rows) < 2) {
      warnings <- c(warnings, sprintf("Curve code %s skipped because it has fewer than 2 valid points.", curve_code))
      next
    }
    if (length(unique(curve_rows$AdjustedLayerMP)) < 2) {
      warnings <- c(warnings, sprintf("Curve code %s skipped because valid points do not span at least 2 x-values.", curve_code))
      next
    }

    fit <- stats::lm(log(ROL) ~ log(AdjustedLayerMP), data = curve_rows)
    coefficients <- stats::coef(fit)
    a <- exp(unname(coefficients[[1]]))
    b <- unname(coefficients[[2]])

    parameter_rows[[length(parameter_rows) + 1L]] <- data.frame(
      CurveCode = curve_code,
      a = a,
      b = b,
      Points = nrow(curve_rows),
      MinX = min(curve_rows$AdjustedLayerMP),
      MaxX = max(curve_rows$AdjustedLayerMP),
      stringsAsFactors = FALSE
    )
  }

  parameters <- if (length(parameter_rows) > 0) {
    do.call(rbind, parameter_rows)
  } else {
    data.frame(CurveCode = character(), a = numeric(), b = numeric(), Points = integer(), MinX = numeric(), MaxX = numeric(), stringsAsFactors = FALSE)
  }
  rownames(parameters) <- NULL
  list(parameters = parameters, warnings = warnings)
}

xl_pricing_tool_relativepricing_render_plots <- function(fitting_input, parameters, output_dir) {
  if (!requireNamespace("ggplot2", quietly = TRUE)) {
    stop("Package 'ggplot2' is required to render RelativePricing chart images.", call. = FALSE)
  }

  ggplot2 <- asNamespace("ggplot2")
  plot_files <- character(nrow(parameters))

  for (i in seq_len(nrow(parameters))) {
    parameter <- parameters[i, , drop = FALSE]
    curve_code <- parameter$CurveCode[[1]]
    points <- fitting_input[
      fitting_input$CurveCode == curve_code &
        is.finite(fitting_input$ROL) &
        is.finite(fitting_input$AdjustedLayerMP) &
        fitting_input$ROL > 0 &
        fitting_input$AdjustedLayerMP > 0,
      ,
      drop = FALSE
    ]
    curve_x <- seq(parameter$MinX[[1]], parameter$MaxX[[1]], length.out = 100)
    curve <- data.frame(
      AdjustedLayerMP = curve_x,
      FittedROL = parameter$a[[1]] * curve_x ^ parameter$b[[1]]
    )
    output_file <- file.path(output_dir, sprintf("relativepricing_curve_%s.png", xl_pricing_tool_relativepricing_safe_filename(curve_code)))

    plot <- ggplot2$ggplot() +
      ggplot2$geom_point(
        data = points,
        ggplot2$aes(x = AdjustedLayerMP, y = ROL),
        color = "#2F6F9F",
        size = 2.4
      ) +
      ggplot2$geom_line(
        data = curve,
        ggplot2$aes(x = AdjustedLayerMP, y = FittedROL),
        color = "#D95F02",
        size = 0.8
      ) +
      ggplot2$scale_y_continuous(labels = xl_pricing_tool_relativepricing_percent_labels) +
      ggplot2$labs(
        title = sprintf("Curve Code %s", curve_code),
        subtitle = sprintf("ROL = %.6g * MP ^ %.6g", parameter$a[[1]], parameter$b[[1]]),
        x = "(Adj.) Layer MP",
        y = "ROL"
      ) +
      ggplot2$theme_minimal(base_size = 9) +
      ggplot2$theme(
        plot.title = ggplot2$element_text(face = "bold", color = "#22313F", size = 11),
        plot.subtitle = ggplot2$element_text(color = "#52606D", size = 8),
        panel.grid.minor = ggplot2$element_blank(),
        panel.grid.major = ggplot2$element_line(color = "#EBEEF2", size = 0.25),
        axis.text = ggplot2$element_text(color = "#52606D"),
        plot.margin = ggplot2$margin(6, 8, 6, 8)
      )

    ggplot2$ggsave(output_file, plot = plot, width = 3.9, height = 2.75, dpi = 160, bg = "white")
    plot_files[[i]] <- output_file
  }

  plot_files
}

xl_pricing_tool_relativepricing_safe_filename <- function(value) {
  value <- trimws(as.character(value))
  value <- gsub("[\\\\/:*?\"<>| ]+", "_", value)
  value <- gsub("[^A-Za-z0-9_.-]", "_", value)
  if (!nzchar(value)) {
    value <- "curve"
  }
  value
}

xl_pricing_tool_relativepricing_percent_labels <- function(values) {
  paste0(format(round(values * 100, 2), trim = TRUE, scientific = FALSE), "%")
}

xl_pricing_tool_relativepricing_check_marker <- function(sheet) {
  marker <- xl_pricing_tool_normalize_marker(xl_pricing_tool_cell_value(sheet, "A1"))
  if (!identical(marker, "relativepricing") && !identical(marker, "relative pricing")) {
    stop("RelativePricing sheet must have <<RelativePricing>> in A1.", call. = FALSE)
  }

  "RelativePricing marker found in A1."
}

xl_pricing_tool_relativepricing_check_layout <- function(sheet) {
  fitting_headers <- vapply(6:15, function(col) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), 13)))
  }, character(1))
  expected_fitting <- c(
    "Layer",
    "Coverage",
    "Limit",
    "Deductible",
    "ROL",
    "Risk Adjustments",
    "(Adj.) Layer MP",
    "Curve Code",
    "Risk Adj. Basis",
    "Deviation"
  )
  if (!all(fitting_headers == expected_fitting)) {
    stop("RelativePricing fitting structure must have expected headers in F13:O13.", call. = FALSE)
  }

  pricing_headers <- vapply(6:15, function(col) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0(xl_pricing_tool_col_name(col), 40)))
  }, character(1))
  expected_pricing <- c(
    "Layer",
    "Coverage",
    "Limit",
    "Deductible",
    "Risk Adjustments",
    "(Adj.) Layer MP",
    "CurveCode",
    "Fitted ROL",
    "Deviation",
    "Adj. ROL"
  )
  if (!all(pricing_headers == expected_pricing)) {
    stop("RelativePricing pricing structure must have expected headers in F40:O40.", call. = FALSE)
  }

  layer_rows <- xl_pricing_tool_relativepricing_count_rows(sheet, "F", 14)
  pricing_rows <- xl_pricing_tool_relativepricing_count_rows(sheet, "F", 41)
  if (layer_rows == 0 || pricing_rows == 0) {
    stop("RelativePricing must have at least one fitting row and one pricing row.", call. = FALSE)
  }

  sprintf("RelativePricing layout valid: %s fitting row(s), %s pricing row(s).", layer_rows, pricing_rows)
}

xl_pricing_tool_relativepricing_check_input_layers <- function(sheet) {
  headers <- c(
    Layer = "B",
    Coverage = "C",
    Limit = "D",
    Deductible = "E",
    ROL = "I",
    GNPI = "J",
    Aggregates = "K",
    Keyzone = "L",
    LayerName = "N"
  )
  values <- vapply(headers, function(col) {
    trimws(xl_pricing_tool_cell_value(sheet, paste0(col, 12)))
  }, character(1))
  expected <- c(
    Layer = "Layer",
    Coverage = "Coverage",
    Limit = "Limit",
    Deductible = "Deductible",
    ROL = "ROL",
    GNPI = "GNPI",
    Aggregates = "Aggregates",
    Keyzone = "KZ Aggregates",
    LayerName = "Layer Name"
  )
  if (!all(values == expected)) {
    stop("Input_Layers must have expected RelativePricing fields in B12:E12, I12:L12, and N12.", call. = FALSE)
  }

  rows <- xl_pricing_tool_relativepricing_count_rows(sheet, "B", 13)
  if (rows == 0) {
    stop("Input_Layers must have at least one layer row for RelativePricing.", call. = FALSE)
  }

  sprintf("Input_Layers contract valid: %s layer row(s).", rows)
}

xl_pricing_tool_relativepricing_check_names <- function(workbook_path) {
  required <- c(
    "Layer_ID",
    "Layer_Coverage",
    "Layer_Limit",
    "Layer_Deductible",
    "Layer_ROL",
    "Layer_Keyzone"
  )
  names <- xl_pricing_tool_relativepricing_defined_names(workbook_path)
  missing <- required[!tolower(required) %in% tolower(names$Name)]
  if (length(missing) > 0) {
    stop(sprintf("RelativePricing is missing required defined name(s): %s", paste(missing, collapse = ", ")), call. = FALSE)
  }

  broken <- names[names$Name %in% required & grepl("#REF!|#NAME\\?|#N/A", names$Formula), , drop = FALSE]
  if (nrow(broken) > 0) {
    stop(sprintf("RelativePricing defined name(s) are broken: %s", paste(unique(broken$Name), collapse = ", ")), call. = FALSE)
  }

  sprintf("RelativePricing defined names valid: %s.", paste(required, collapse = ", "))
}

xl_pricing_tool_relativepricing_defined_names <- function(workbook_path) {
  temp_dir <- tempfile("xl_pricing_tool_names_")
  dir.create(temp_dir, recursive = TRUE, showWarnings = FALSE)
  on.exit(unlink(temp_dir, recursive = TRUE, force = TRUE), add = TRUE)

  workbook_xml <- utils::unzip(workbook_path, files = "xl/workbook.xml", exdir = temp_dir)
  if (length(workbook_xml) == 0 || !file.exists(workbook_xml[1])) {
    stop("Workbook package is missing xl/workbook.xml.", call. = FALSE)
  }

  xml <- paste(readLines(workbook_xml[1], warn = FALSE), collapse = "")
  defined_name_tags <- regmatches(xml, gregexpr("<definedName\\b[^>]*>.*?</definedName>", xml, perl = TRUE))[[1]]
  if (identical(defined_name_tags[1], -1L)) {
    return(data.frame(Name = character(), Formula = character(), stringsAsFactors = FALSE))
  }

  data.frame(
    Name = vapply(defined_name_tags, function(tag) {
      xl_pricing_tool_xml_unescape(sub('.*\\bname="([^"]+)".*', "\\1", tag))
    }, character(1)),
    Formula = vapply(defined_name_tags, function(tag) {
      xl_pricing_tool_xml_unescape(sub(".*>(.*?)</definedName>.*", "\\1", tag))
    }, character(1)),
    stringsAsFactors = FALSE
  )
}

xl_pricing_tool_relativepricing_count_rows <- function(sheet, col, first_row) {
  count <- 0
  blank_run <- 0
  for (row in first_row:(first_row + 500)) {
    value <- trimws(xl_pricing_tool_cell_value(sheet, paste0(col, row)))
    if (!nzchar(value)) {
      blank_run <- blank_run + 1
      if (blank_run >= 10) break
      next
    }
    blank_run <- 0
    count <- count + 1
  }

  count
}
