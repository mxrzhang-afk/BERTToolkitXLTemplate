china_exposure_map_validate <- function(workbook_path, context = list()) {
  btk_require_file(workbook_path, "workbook")
  btk_require_packages(c("readxl", "openxlsx"))

  sheets <- readxl::excel_sheets(workbook_path)
  required_sheets <- c("Control", "Parameters", "Input_Data", "Output")
  missing_sheets <- setdiff(required_sheets, sheets)
  if (length(missing_sheets) > 0) {
    stop(
      sprintf("Workbook is missing required sheet(s): %s", paste(missing_sheets, collapse = ", ")),
      call. = FALSE
    )
  }

  input_data <- btk_read_sheet(workbook_path, "Input_Data")
  btk_require_columns(
    input_data,
    c("Province", "Lat", "Lon", "EQ_New", "EQ_Old", "WS_New", "WS_Old"),
    "Input_Data"
  )

  "OK"
}

china_exposure_map_update <- function(workbook_path, output_dir = NULL, context = list()) {
  china_exposure_map_validate(workbook_path, context)

  btk_require_packages(c("sf", "ggplot2", "magick", "grid", "dplyr", "readxl"))

  if (is.null(output_dir) || identical(output_dir, "")) {
    output_dir <- if (!is.null(context$output_dir)) {
      context$output_dir
    } else {
      btk_default_run_dir("china_exposure_map", workbook_path)
    }
  }
  dir.create(output_dir, recursive = TRUE, showWarnings = FALSE)

  input_data <- as.data.frame(btk_read_sheet(workbook_path, "Input_Data"))
  parameters <- as.data.frame(readxl::read_excel(workbook_path, sheet = "Parameters", skip = 2))
  btk_require_columns(
    parameters,
    c("Metric", "Class", "Lower_Bound", "Upper_Bound", "Label", "Fill_Color", "Legend_Title", "Bar_Scale"),
    "Parameters"
  )

  for (col in c("Lat", "Lon", "EQ_New", "EQ_Old", "WS_New", "WS_Old")) {
    input_data[[col]] <- as.numeric(input_data[[col]])
  }

  eq_params <- china_exposure_map_metric_params(parameters, "EQ")
  ws_params <- china_exposure_map_metric_params(parameters, "WS")

  asset_dir <- china_exposure_map_asset_dir()
  china_shp <- file.path(asset_dir, "China_Province_4326.shp")
  border_shp <- file.path(asset_dir, "guojie.shp")
  btk_require_file(china_shp, "China province shapefile")
  btk_require_file(border_shp, "China border shapefile")

  old_wd <- getwd()
  on.exit(setwd(old_wd), add = TRUE)
  setwd(output_dir)

  china_map <- sf::st_read(china_shp, quiet = TRUE)
  china_map <- sf::st_transform(china_map, 4544)
  china_map_frame <- ggplot2::fortify(china_map)

  guojie <- sf::st_read(border_shp, quiet = TRUE)
  guojie <- sf::st_transform(guojie, 4544)

  dt <- china_exposure_map_polygon_data(china_map_frame)
  dt[substr(dt$ID, 1, 6) == "810000", "Province"] <- "Hongkong"
  dt[substr(dt$ID, 1, 6) == "820000", "Province"] <- "Macau"

  si <- china_exposure_map_project_points(input_data)
  dt <- dplyr::left_join(dt, si, by = c("Province" = "Province"))

  eq_file <- file.path(output_dir, "plot_EQ.png")
  ws_file <- file.path(output_dir, "plot_WS.png")

  china_exposure_map_render_metric(
    dt = dt,
    si = si,
    guojie = guojie,
    metric = "EQ",
    params = eq_params,
    output_file = eq_file,
    highlight_provinces = c(
      "Beijing", "Gansu", "Hainan", "Inner Mongolia", "Ningxia",
      "Shanxi", "Sichuan", "Tianjin", "Xinjiang", "Yunnan"
    ),
    yoy_colors = c("China SI 2025" = "#004DA8", "China SI 2024" = "#73B2FF")
  )

  china_exposure_map_render_metric(
    dt = dt,
    si = si,
    guojie = guojie,
    metric = "WS",
    params = ws_params,
    output_file = ws_file,
    highlight_provinces = c(
      "Fujian", "Guangdong", "Guangxi", "Hainan", "Jiangxi",
      "Shanghai", "Zhejiang"
    ),
    yoy_colors = c("China SI 2025" = "#A80000", "China SI 2024" = "#FFBEBE")
  )

  paste(
    "China exposure maps created.",
    sprintf("Output folder: %s", output_dir),
    sprintf("EQ: %s", eq_file),
    sprintf("WS: %s", ws_file),
    sep = vb_newline()
  )
}

china_exposure_map_run <- china_exposure_map_update

china_exposure_map_asset_dir <- function() {
  file.path(BTK_ROOT, "r", "tools", "china_exposure_map", "assets")
}

china_exposure_map_metric_params <- function(parameters, metric) {
  params <- parameters[parameters$Metric == metric, ]
  params <- params[
    !is.na(params$Class) &
      !is.na(params$Lower_Bound) &
      !is.na(params$Upper_Bound) &
      !is.na(params$Label) &
      !is.na(params$Fill_Color),
  ]
  if (nrow(params) == 0) {
    stop(sprintf("Parameters sheet has no rows for metric: %s", metric), call. = FALSE)
  }

  params$Class <- suppressWarnings(as.numeric(params$Class))
  params <- params[!is.na(params$Class), ]
  if (nrow(params) == 0) {
    stop(sprintf("Parameters sheet has no class-break rows for metric: %s", metric), call. = FALSE)
  }

  params <- params[order(params$Class), ]
  lower <- vapply(params$Lower_Bound, china_exposure_map_bound, numeric(1))
  upper <- vapply(params$Upper_Bound, china_exposure_map_bound, numeric(1))
  labels <- as.character(params$Label)

  if (length(c(lower[1], upper)) != length(labels) + 1) {
    stop(
      sprintf(
        "Parameters for %s must have one more break than label. Got %d breaks and %d labels.",
        metric,
        length(c(lower[1], upper)),
        length(labels)
      ),
      call. = FALSE
    )
  }

  list(
    breaks = c(lower[1], upper),
    labels = labels,
    colors = stats::setNames(as.character(params$Fill_Color), labels),
    legend = as.character(params$Legend_Title[1]),
    bar_scale = as.numeric(params$Bar_Scale[1])
  )
}

china_exposure_map_bound <- function(x) {
  x <- as.character(x)
  if (identical(x, "-Inf")) {
    return(-Inf)
  }
  if (identical(x, "Inf")) {
    return(Inf)
  }
  as.numeric(x)
}

china_exposure_map_polygon_data <- function(china_map_frame) {
  dt <- data.frame()
  for (i in seq_len(nrow(china_map_frame))) {
    dt_set <- as.data.frame(sf::st_coordinates(china_map_frame[i, ]))
    dt_set$Province <- china_map_frame$NAME[i]
    dt_set$ID <- NA_character_

    if (length(china_map_frame$geometry[[i]]) == 1) {
      dt_set$ID <- china_map_frame$ID[i]
    } else {
      for (j in seq_along(unique(dt_set$L2))) {
        dt_set[dt_set$L2 == j, "ID"] <- paste0(china_map_frame$ID[i], "-", j)
      }
    }

    dt <- rbind(dt, dt_set)
  }
  dt
}

china_exposure_map_project_points <- function(input_data) {
  si <- input_data
  rows <- which(is.finite(si$Lon) & is.finite(si$Lat))
  if (length(rows) == 0) {
    stop("Input_Data has no finite Lat/Lon rows.", call. = FALSE)
  }

  points <- sf::st_as_sf(si[rows, ], coords = c("Lon", "Lat"), crs = 4326, remove = FALSE)
  points <- sf::st_transform(points, 4544)
  coords <- as.data.frame(sf::st_coordinates(points))
  si$Lon[rows] <- coords$X
  si$Lat[rows] <- coords$Y
  si
}

china_exposure_map_render_metric <- function(
  dt,
  si,
  guojie,
  metric,
  params,
  output_file,
  highlight_provinces,
  yoy_colors
) {
  new_col <- paste0(metric, "_New")
  old_col <- paste0(metric, "_Old")
  class_col <- paste0(metric, "_New2")

  dt[[class_col]] <- cut(dt[[new_col]], breaks = params$breaks, labels = params$labels)
  highlight <- dt[dt$Province %in% highlight_provinces, ]

  combined_file <- sub("\\.png$", "_combined.png", output_file)
  cropped_file <- sub("\\.png$", "_cropped.png", output_file)
  yoy_file <- sub("\\.png$", "_yoy.png", output_file)

  p <- ggplot2::ggplot() +
    ggplot2::geom_polygon(
      data = dt,
      ggplot2::aes(x = X, y = Y, group = ID, fill = .data[[class_col]]),
      color = "#6E6E6E"
    ) +
    ggplot2::scale_fill_manual(values = params$colors, name = params$legend, drop = FALSE) +
    ggplot2::geom_sf(data = guojie, color = "#6E6E6E") +
    ggplot2::geom_polygon(
      data = highlight,
      ggplot2::aes(x = X, y = Y, group = ID),
      color = "#00FFC5",
      alpha = 0,
      linewidth = 1
    ) +
    china_exposure_map_theme(legend_title_size = 12, legend_text_size = 11, legend_position = c(0.25, 0.2))
  ggplot2::ggsave(combined_file, plot = p, width = 10, height = 7.5, units = "in", dpi = 300)

  img <- magick::image_read(combined_file)
  cropped_img <- magick::image_crop(img, "600x890+1670+1260")
  bordered_cropped_img <- magick::image_border(cropped_img, color = "black", geometry = "3x3")
  magick::image_write(bordered_cropped_img, cropped_file)
  img1 <- magick::image_read(cropped_file)

  label_rows <- seq_len(min(31, nrow(si)))
  p <- ggplot2::ggplot() +
    ggplot2::geom_polygon(
      data = dt,
      ggplot2::aes(x = X, y = Y, group = ID, fill = .data[[class_col]]),
      color = "#6E6E6E"
    ) +
    ggplot2::scale_fill_manual(values = params$colors, name = params$legend, drop = FALSE) +
    ggplot2::geom_polygon(
      data = highlight,
      ggplot2::aes(x = X, y = Y, group = ID),
      color = "#00FFC5",
      alpha = 0,
      linewidth = 1
    ) +
    ggplot2::geom_errorbar(
      data = si[label_rows, ],
      ggplot2::aes(
        x = Lon - 24500,
        ymin = Lat,
        ymax = Lat + params$bar_scale * .data[[new_col]],
        color = "China SI 2025"
      ),
      linewidth = 3,
      width = 0
    ) +
    ggplot2::geom_errorbar(
      data = si[label_rows, ],
      ggplot2::aes(
        x = Lon + 24500,
        ymin = Lat,
        ymax = Lat + params$bar_scale * .data[[old_col]],
        color = "China SI 2024"
      ),
      linewidth = 3,
      width = 0
    ) +
    ggplot2::scale_color_manual(name = "Exposure YOY", values = yoy_colors) +
    ggplot2::geom_text(
      data = si[label_rows, ],
      ggplot2::aes(x = Lon, y = Lat - 40000, label = Province),
      size = 2.5
    ) +
    china_exposure_map_theme(legend_title_size = 11, legend_text_size = 10, legend_position = c(0.1, 0.2))
  ggplot2::ggsave(yoy_file, plot = p, width = 10, height = 7.5, units = "in", dpi = 300)
  img2 <- magick::image_read(yoy_file)

  p <- ggplot2::ggplot() +
    ggplot2::theme_void() +
    ggplot2::coord_equal() +
    ggplot2::theme(plot.margin = ggplot2::margin(0, 0, 0, 0, "cm")) +
    ggplot2::annotation_custom(
      grid::rasterGrob(
        as.raster(img1),
        x = grid::unit(0.87, "npc"),
        y = grid::unit(0.35, "npc"),
        width = grid::unit(0.09, "npc"),
        height = grid::unit(0.1335, "npc")
      ),
      xmin = -Inf,
      xmax = Inf,
      ymin = -Inf,
      ymax = Inf
    ) +
    ggplot2::annotation_custom(
      grid::rasterGrob(
        as.raster(img2),
        x = grid::unit(0.5, "npc"),
        y = grid::unit(0.5, "npc"),
        width = grid::unit(0.64, "npc"),
        height = grid::unit(0.48, "npc")
      ),
      xmin = -Inf,
      xmax = Inf,
      ymin = -Inf,
      ymax = Inf
    )
  ggplot2::ggsave(output_file, plot = p, width = 10, height = 7.5, units = "in", dpi = 300)
}

china_exposure_map_theme <- function(legend_title_size, legend_text_size, legend_position) {
  ggplot2::theme(
    panel.grid = ggplot2::element_blank(),
    panel.background = ggplot2::element_blank(),
    axis.text = ggplot2::element_blank(),
    axis.ticks = ggplot2::element_blank(),
    axis.title = ggplot2::element_blank(),
    legend.title = ggplot2::element_text(size = legend_title_size, face = "bold"),
    legend.text = ggplot2::element_text(size = legend_text_size, face = "bold"),
    legend.position = legend_position
  )
}

vb_newline <- function() {
  "\r\n"
}
