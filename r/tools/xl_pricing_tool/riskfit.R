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

  sheet <- xl_pricing_tool_action_sheet(workbook_path, context, "Risk Loss Fitting")
  loss_input <- xl_pricing_tool_riskfit_read_loss_input(sheet)
  chart_config <- xl_pricing_tool_riskfit_read_loss_comparison_config(sheet)
  layers <- xl_pricing_tool_riskfit_read_loss_layers(sheet)
  chart_data <- xl_pricing_tool_riskfit_loss_comparison_data(loss_input, chart_config)
  severity_fits <- xl_pricing_tool_riskfit_fit_severity(loss_input, chart_config)

  chart_data_file <- file.path(riskfit_dir, "riskfit_loss_comparison_data.csv")
  fit_summary_file <- file.path(riskfit_dir, "riskfit_fit_summary.csv")
  fit_parameters_file <- file.path(riskfit_dir, "riskfit_fit_parameters.csv")
  final_chart_file <- file.path(riskfit_dir, "riskfit_final_incurred_loss_chart.png")
  actual_chart_file <- file.path(riskfit_dir, "riskfit_actual_incurred_loss_chart.png")

  utils::write.csv(chart_data, chart_data_file, row.names = FALSE, na = "")
  utils::write.csv(severity_fits$summary, fit_summary_file, row.names = FALSE, na = "")
  utils::write.csv(severity_fits$parameters, fit_parameters_file, row.names = FALSE, na = "")
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
  cdf_chart_files <- xl_pricing_tool_riskfit_render_cdf_charts(severity_fits, chart_config, riskfit_dir)

  paste(
    "RiskFit update completed.",
    sprintf("Output folder: %s", output_dir),
    sprintf("Input rows: %s", nrow(loss_input)),
    sprintf("Chart rows: %s", nrow(chart_data)),
    sprintf("Layer rows: %s", nrow(layers)),
    sprintf("Severity fit rows: %s", nrow(severity_fits$summary)),
    sprintf("Loss comparison chart data: %s", chart_data_file),
    sprintf("Severity fit summary: %s", fit_summary_file),
    sprintf("Severity fit parameters: %s", fit_parameters_file),
    sprintf("Final incurred chart image: %s", final_chart_file),
    if (chart_config$show_actual) sprintf("Actual incurred chart image: %s", actual_chart_file) else "Actual incurred chart image: skipped because showActual is FALSE.",
    sprintf("Severity CDF chart images: %s", paste(basename(cdf_chart_files), collapse = ", ")),
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
    alias_headers <- headers
    alias_headers[alias_headers == "EventID"] <- "ClaimID"
    alias_headers[alias_headers == "EventLabel"] <- "Insured"
    if (!all(alias_headers == expected)) {
      stop("Risk Loss Fitting loss table must have expected headers in B12:P12. ClaimID/EventID and Insured/EventLabel are both supported.", call. = FALSE)
    }
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
    final_rate <- xl_pricing_tool_riskfit_number(sheet, "N", row, "FinalRate", required = FALSE)
    final_incurred <- xl_pricing_tool_riskfit_number(sheet, "P", row, "Final_Incurred", required = FALSE)

    rows[[length(rows) + 1L]] <- data.frame(
      AsAt = as_at,
      ClaimID = claim_id,
      LOB = trimws(xl_pricing_tool_cell_value(sheet, paste0("D", row))),
      Insured = insured,
      TreatyYear = as.integer(treaty_year),
      Actual_Incurred = actual_incurred,
      FinalRate = final_rate,
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

xl_pricing_tool_riskfit_fit_severity <- function(loss_input, chart_config) {
  families <- xl_pricing_tool_riskfit_families()
  as_ats <- c("Prior", "Current")
  sample_data <- loss_input[
    loss_input$TreatyYear >= chart_config$start_year &
      loss_input$AsAt %in% as_ats &
      is.finite(loss_input$Final_Incurred) &
      is.finite(loss_input$FinalRate) &
      loss_input$Final_Incurred >= chart_config$min_loss &
      loss_input$FinalRate > 0,
    ,
    drop = FALSE
  ]

  summary_rows <- list()
  parameter_rows <- list()
  cdf_rows <- list()
  row_id <- 0L

  for (family in families$family) {
    for (as_at in as_ats) {
      row_id <- row_id + 1L
      family_info <- families[families$family == family, , drop = FALSE]
      sample <- sample_data[sample_data$AsAt == as_at, , drop = FALSE]
      fit <- xl_pricing_tool_riskfit_fit_family(
        family = family,
        loss = sample$Final_Incurred,
        weight = sample$FinalRate,
        threshold = chart_config$min_loss
      )

      summary_rows[[row_id]] <- data.frame(
        Family = family,
        FamilyLabel = family_info$label,
        AsAt = as_at,
        Status = fit$status,
        Message = fit$message,
        RawCount = fit$raw_count,
        TotalWeight = fit$total_weight,
        LogLik = fit$loglik,
        AIC = fit$aic,
        BIC = fit$bic,
        Mean = fit$mean,
        Q50 = xl_pricing_tool_riskfit_quantile_fit(fit, 0.5),
        Q90 = xl_pricing_tool_riskfit_quantile_fit(fit, 0.9),
        Q95 = xl_pricing_tool_riskfit_quantile_fit(fit, 0.95),
        Q99 = xl_pricing_tool_riskfit_quantile_fit(fit, 0.99),
        stringsAsFactors = FALSE
      )

      parameter_rows <- c(parameter_rows, xl_pricing_tool_riskfit_parameter_rows(fit, family_info$label, as_at))
      if (fit$status == "OK") {
        cdf_rows[[length(cdf_rows) + 1L]] <- xl_pricing_tool_riskfit_cdf_rows(
          sample = sample,
          fit = fit,
          family_label = family_info$label,
          as_at = as_at
        )
      }
    }
  }

  summary <- do.call(rbind, summary_rows)
  parameters <- if (length(parameter_rows) == 0) {
    data.frame(Family = character(), FamilyLabel = character(), AsAt = character(), Parameter = character(), Value = numeric(), Status = character(), stringsAsFactors = FALSE)
  } else {
    do.call(rbind, parameter_rows)
  }
  cdf <- if (length(cdf_rows) == 0) {
    data.frame(Family = character(), FamilyLabel = character(), AsAt = character(), Loss = numeric(), EmpiricalCDF = numeric(), FittedCDF = numeric(), stringsAsFactors = FALSE)
  } else {
    do.call(rbind, cdf_rows)
  }

  list(summary = summary, parameters = parameters, cdf = cdf)
}

xl_pricing_tool_riskfit_families <- function() {
  data.frame(
    family = c("pareto", "loggamma", "weibull", "lognormal", "gamma", "loglogistic"),
    label = c("Pareto", "Loggamma", "Weibull", "Lognormal", "Gamma", "Loglogistic"),
    stringsAsFactors = FALSE
  )
}

xl_pricing_tool_riskfit_fit_family <- function(family, loss, weight, threshold) {
  valid <- is.finite(loss) & is.finite(weight) & loss > 0 & weight > 0
  loss <- loss[valid]
  weight <- weight[valid]

  min_count <- if (family == "pareto") 2L else 3L
  if (length(loss) < min_count) {
    return(xl_pricing_tool_riskfit_failed_fit(family, length(loss), sum(weight), sprintf("Insufficient data: %s valid rows.", length(loss))))
  }

  fit <- tryCatch(
    switch(
      family,
      pareto = xl_pricing_tool_riskfit_fit_pareto(loss, weight, threshold),
      lognormal = xl_pricing_tool_riskfit_fit_lognormal(loss, weight),
      gamma = xl_pricing_tool_riskfit_fit_gamma(loss, weight),
      loggamma = xl_pricing_tool_riskfit_fit_loggamma(loss, weight),
      weibull = xl_pricing_tool_riskfit_fit_weibull(loss, weight),
      loglogistic = xl_pricing_tool_riskfit_fit_loglogistic(loss, weight),
      stop(sprintf("Unsupported severity family: %s", family), call. = FALSE)
    ),
    error = function(e) xl_pricing_tool_riskfit_failed_fit(family, length(loss), sum(weight), conditionMessage(e))
  )

  fit$raw_count <- length(loss)
  fit$total_weight <- sum(weight)
  fit
}

xl_pricing_tool_riskfit_failed_fit <- function(family, raw_count, total_weight, message) {
  list(
    family = family,
    status = "Failed",
    message = message,
    params = list(),
    loglik = NA_real_,
    aic = NA_real_,
    bic = NA_real_,
    mean = NA_real_,
    raw_count = raw_count,
    total_weight = total_weight,
    cdf = function(x) rep(NA_real_, length(x)),
    quantile = function(p) rep(NA_real_, length(p))
  )
}

xl_pricing_tool_riskfit_ok_fit <- function(family, params, loglik, raw_count, total_weight, cdf, quantile, mean = NA_real_) {
  k <- length(params)
  list(
    family = family,
    status = "OK",
    message = "",
    params = params,
    loglik = loglik,
    aic = 2 * k - 2 * loglik,
    bic = log(max(total_weight, 1)) * k - 2 * loglik,
    mean = mean,
    raw_count = raw_count,
    total_weight = total_weight,
    cdf = cdf,
    quantile = quantile
  )
}

xl_pricing_tool_riskfit_fit_pareto <- function(loss, weight, threshold) {
  xm <- threshold
  if (!is.finite(xm) || xm <= 0) {
    stop("Pareto threshold must be positive.", call. = FALSE)
  }
  keep <- loss >= xm
  loss <- loss[keep]
  weight <- weight[keep]
  if (length(loss) < 2) {
    stop("Pareto fit needs at least two losses at or above threshold.", call. = FALSE)
  }

  alpha <- sum(weight) / sum(weight * log(loss / xm))
  if (!is.finite(alpha) || alpha <= 0) {
    stop("Pareto alpha could not be estimated.", call. = FALSE)
  }
  loglik <- sum(weight * (log(alpha) + alpha * log(xm) - (alpha + 1) * log(loss)))
  mean <- if (alpha > 1) alpha * xm / (alpha - 1) else Inf

  xl_pricing_tool_riskfit_ok_fit(
    family = "pareto",
    params = list(threshold = xm, alpha = alpha),
    loglik = loglik,
    raw_count = length(loss),
    total_weight = sum(weight),
    mean = mean,
    cdf = function(x) ifelse(x < xm, 0, 1 - (xm / x)^alpha),
    quantile = function(p) xm / (1 - p)^(1 / alpha)
  )
}

xl_pricing_tool_riskfit_fit_lognormal <- function(loss, weight) {
  z <- log(loss)
  meanlog <- sum(weight * z) / sum(weight)
  sdlog <- sqrt(sum(weight * (z - meanlog)^2) / sum(weight))
  if (!is.finite(sdlog) || sdlog <= 0) {
    stop("Lognormal sigma must be positive.", call. = FALSE)
  }
  loglik <- sum(weight * stats::dlnorm(loss, meanlog = meanlog, sdlog = sdlog, log = TRUE))

  xl_pricing_tool_riskfit_ok_fit(
    family = "lognormal",
    params = list(meanlog = meanlog, sigma = sdlog),
    loglik = loglik,
    raw_count = length(loss),
    total_weight = sum(weight),
    mean = exp(meanlog + sdlog^2 / 2),
    cdf = function(x) stats::plnorm(x, meanlog = meanlog, sdlog = sdlog),
    quantile = function(p) stats::qlnorm(p, meanlog = meanlog, sdlog = sdlog)
  )
}

xl_pricing_tool_riskfit_fit_gamma <- function(loss, weight) {
  start <- xl_pricing_tool_riskfit_gamma_start(loss, weight)
  objective <- function(par) {
    shape <- exp(par[[1]])
    rate <- exp(par[[2]])
    -sum(weight * stats::dgamma(loss, shape = shape, rate = rate, log = TRUE))
  }
  opt <- xl_pricing_tool_riskfit_optim(objective, log(c(start$shape, start$rate)))
  shape <- exp(opt$par[[1]])
  rate <- exp(opt$par[[2]])
  loglik <- -opt$value

  xl_pricing_tool_riskfit_ok_fit(
    family = "gamma",
    params = list(shape = shape, rate = rate),
    loglik = loglik,
    raw_count = length(loss),
    total_weight = sum(weight),
    mean = shape / rate,
    cdf = function(x) stats::pgamma(x, shape = shape, rate = rate),
    quantile = function(p) stats::qgamma(p, shape = shape, rate = rate)
  )
}

xl_pricing_tool_riskfit_fit_loggamma <- function(loss, weight) {
  z <- log(loss)
  start <- xl_pricing_tool_riskfit_gamma_start(z, weight)
  objective <- function(par) {
    shape <- exp(par[[1]])
    rate <- exp(par[[2]])
    -sum(weight * (stats::dgamma(z, shape = shape, rate = rate, log = TRUE) - log(loss)))
  }
  opt <- xl_pricing_tool_riskfit_optim(objective, log(c(start$shape, start$rate)))
  shape <- exp(opt$par[[1]])
  rate <- exp(opt$par[[2]])
  loglik <- -opt$value

  xl_pricing_tool_riskfit_ok_fit(
    family = "loggamma",
    params = list(shape = shape, rate = rate),
    loglik = loglik,
    raw_count = length(loss),
    total_weight = sum(weight),
    mean = if (rate > 1) (rate / (rate - 1))^shape else Inf,
    cdf = function(x) stats::pgamma(log(x), shape = shape, rate = rate),
    quantile = function(p) exp(stats::qgamma(p, shape = shape, rate = rate))
  )
}

xl_pricing_tool_riskfit_fit_weibull <- function(loss, weight) {
  objective <- function(par) {
    shape <- exp(par[[1]])
    scale <- exp(par[[2]])
    -sum(weight * stats::dweibull(loss, shape = shape, scale = scale, log = TRUE))
  }
  opt <- xl_pricing_tool_riskfit_optim(objective, log(c(1.2, stats::median(loss))))
  shape <- exp(opt$par[[1]])
  scale <- exp(opt$par[[2]])
  loglik <- -opt$value

  xl_pricing_tool_riskfit_ok_fit(
    family = "weibull",
    params = list(shape = shape, scale = scale),
    loglik = loglik,
    raw_count = length(loss),
    total_weight = sum(weight),
    mean = scale * gamma(1 + 1 / shape),
    cdf = function(x) stats::pweibull(x, shape = shape, scale = scale),
    quantile = function(p) stats::qweibull(p, shape = shape, scale = scale)
  )
}

xl_pricing_tool_riskfit_fit_loglogistic <- function(loss, weight) {
  objective <- function(par) {
    shape <- exp(par[[1]])
    scale <- exp(par[[2]])
    z <- (loss / scale)^shape
    log_density <- log(shape) - log(scale) + (shape - 1) * log(loss / scale) - 2 * log1p(z)
    -sum(weight * log_density)
  }
  opt <- xl_pricing_tool_riskfit_optim(objective, log(c(2, stats::median(loss))))
  shape <- exp(opt$par[[1]])
  scale <- exp(opt$par[[2]])
  loglik <- -opt$value
  mean <- if (shape > 1) scale * pi / shape / sin(pi / shape) else Inf

  xl_pricing_tool_riskfit_ok_fit(
    family = "loglogistic",
    params = list(shape = shape, scale = scale),
    loglik = loglik,
    raw_count = length(loss),
    total_weight = sum(weight),
    mean = mean,
    cdf = function(x) 1 / (1 + (x / scale)^(-shape)),
    quantile = function(p) scale * (p / (1 - p))^(1 / shape)
  )
}

xl_pricing_tool_riskfit_gamma_start <- function(x, weight) {
  mean_x <- sum(weight * x) / sum(weight)
  var_x <- sum(weight * (x - mean_x)^2) / sum(weight)
  if (!is.finite(var_x) || var_x <= 0) {
    var_x <- mean_x^2
  }
  shape <- max(mean_x^2 / var_x, 0.05)
  rate <- max(mean_x / var_x, .Machine$double.eps)
  list(shape = shape, rate = rate)
}

xl_pricing_tool_riskfit_optim <- function(objective, par) {
  opt <- stats::optim(par = par, fn = objective, method = "Nelder-Mead", control = list(maxit = 5000))
  if (!is.finite(opt$value) || opt$convergence != 0) {
    opt <- stats::optim(par = par, fn = objective, method = "BFGS", control = list(maxit = 5000))
  }
  if (!is.finite(opt$value)) {
    stop("Numerical optimization failed.", call. = FALSE)
  }
  opt
}

xl_pricing_tool_riskfit_parameter_rows <- function(fit, family_label, as_at) {
  if (fit$status != "OK") {
    return(list(data.frame(
      Family = fit$family,
      FamilyLabel = family_label,
      AsAt = as_at,
      Parameter = "Status",
      Value = NA_real_,
      Status = fit$message,
      stringsAsFactors = FALSE
    )))
  }

  lapply(names(fit$params), function(param) {
    data.frame(
      Family = fit$family,
      FamilyLabel = family_label,
      AsAt = as_at,
      Parameter = param,
      Value = as.numeric(fit$params[[param]]),
      Status = fit$status,
      stringsAsFactors = FALSE
    )
  })
}

xl_pricing_tool_riskfit_cdf_rows <- function(sample, fit, family_label, as_at) {
  x <- sample$Final_Incurred
  w <- sample$FinalRate
  order_idx <- order(x)
  x <- x[order_idx]
  w <- w[order_idx]
  ecdf <- cumsum(w) / sum(w)
  data.frame(
    Family = fit$family,
    FamilyLabel = family_label,
    AsAt = as_at,
    Loss = x,
    EmpiricalCDF = ecdf,
    FittedCDF = pmin(pmax(fit$cdf(x), 0), 1),
    stringsAsFactors = FALSE
  )
}

xl_pricing_tool_riskfit_quantile_fit <- function(fit, p) {
  if (fit$status != "OK") {
    return(NA_real_)
  }
  value <- fit$quantile(p)
  ifelse(is.finite(value), value, NA_real_)
}

xl_pricing_tool_riskfit_render_cdf_charts <- function(severity_fits, chart_config, riskfit_dir) {
  if (!requireNamespace("ggplot2", quietly = TRUE)) {
    stop("Package 'ggplot2' is required to render RiskFit CDF chart images.", call. = FALSE)
  }

  files <- character()
  cdf_data <- severity_fits$cdf
  families <- xl_pricing_tool_riskfit_families()
  for (i in seq_len(nrow(families))) {
    family <- families$family[[i]]
    family_label <- families$label[[i]]
    family_data <- cdf_data[cdf_data$Family == family, , drop = FALSE]
    if (nrow(family_data) == 0) {
      next
    }

    output_file <- file.path(riskfit_dir, sprintf("riskfit_%s_cdf_chart.png", family))
    xl_pricing_tool_riskfit_render_one_cdf_chart(
      cdf_data = family_data,
      family_label = family_label,
      chart_config = chart_config,
      output_file = output_file
    )
    files <- c(files, output_file)
  }
  files
}

xl_pricing_tool_riskfit_render_one_cdf_chart <- function(cdf_data, family_label, chart_config, output_file) {
  ggplot2 <- asNamespace("ggplot2")
  line_data <- cdf_data[, c("Family", "FamilyLabel", "AsAt", "Loss", "FittedCDF"), drop = FALSE]
  names(line_data)[names(line_data) == "FittedCDF"] <- "CDF"
  line_data$Curve <- paste(line_data$AsAt, "Fitted")

  step_data <- cdf_data[, c("Family", "FamilyLabel", "AsAt", "Loss", "EmpiricalCDF"), drop = FALSE]
  names(step_data)[names(step_data) == "EmpiricalCDF"] <- "CDF"
  step_data$Curve <- paste(step_data$AsAt, "Empirical")

  plot_data <- rbind(step_data, line_data)
  plot_data$AsAt <- factor(plot_data$AsAt, levels = c("Prior", "Current"))
  plot_data$Curve <- factor(
    plot_data$Curve,
    levels = c("Prior Empirical", "Prior Fitted", "Current Empirical", "Current Fitted")
  )

  plot <- ggplot2$ggplot() +
    ggplot2$geom_step(
      data = step_data,
      ggplot2$aes(x = Loss, y = CDF, color = AsAt, linetype = "Empirical"),
      direction = "hv",
      alpha = 0.62,
      size = 0.7
    ) +
    ggplot2$geom_line(
      data = line_data,
      ggplot2$aes(x = Loss, y = CDF, color = AsAt, linetype = "Fitted"),
      size = 0.9
    ) +
    ggplot2$scale_color_manual(values = c(Prior = "#6F9BB3", Current = "#FF8B3D"), drop = FALSE) +
    ggplot2$scale_linetype_manual(values = c(Empirical = "solid", Fitted = "dashed")) +
    ggplot2$scale_x_continuous(labels = xl_pricing_tool_riskfit_amount_labels) +
    ggplot2$scale_y_continuous(labels = xl_pricing_tool_riskfit_percent_labels, limits = c(0, 1), expand = c(0, 0)) +
    ggplot2$labs(
      title = sprintf("%s Severity CDF Comparison", family_label),
      subtitle = sprintf("Weighted empirical CDF vs fitted CDF; min loss %s", xl_pricing_tool_riskfit_amount_labels(chart_config$min_loss)),
      x = NULL,
      y = NULL,
      color = NULL,
      linetype = NULL
    ) +
    ggplot2$theme_minimal(base_size = 9) +
    ggplot2$theme(
      plot.title = ggplot2$element_text(face = "bold", color = "#22313F", size = 11),
      plot.subtitle = ggplot2$element_text(color = "#52606D", size = 8),
      panel.grid.minor = ggplot2$element_blank(),
      panel.grid.major = ggplot2$element_line(color = "#EBEEF2", size = 0.25),
      axis.text = ggplot2$element_text(color = "#52606D"),
      legend.position = "top",
      plot.margin = ggplot2$margin(6, 8, 6, 8)
    )

  ggplot2$ggsave(
    filename = output_file,
    plot = plot,
    width = 4.65,
    height = 2.7,
    dpi = 160,
    bg = "white"
  )
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
  dimensions <- xl_pricing_tool_riskfit_loss_chart_dimensions(chart_data)

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
    width = dimensions$width,
    height = dimensions$height,
    dpi = 160,
    bg = "white"
  )
}

xl_pricing_tool_riskfit_loss_chart_dimensions <- function(chart_data) {
  claim_groups <- length(unique(chart_data$LossID))
  treaty_years <- length(unique(chart_data$TreatyYear))
  width <- 3.6 + 0.36 * claim_groups + 0.18 * treaty_years
  width <- min(max(width, 9.8), 18)
  height <- 3.8
  list(width = width, height = height)
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

xl_pricing_tool_riskfit_percent_labels <- function(values) {
  vapply(values, function(value) {
    if (is.na(value)) {
      return("")
    }
    sprintf("%s%%", format(round(value * 100, 0), trim = TRUE, scientific = FALSE))
  }, character(1))
}
