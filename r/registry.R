source(file.path(BTK_ROOT, "r", "utils", "paths.R"))
source(file.path(BTK_ROOT, "r", "utils", "logging.R"))
source(file.path(BTK_ROOT, "r", "utils", "validation.R"))
source(file.path(BTK_ROOT, "r", "utils", "excel_io.R"))

source(file.path(BTK_ROOT, "r", "tools", "china_exposure_map", "tool.R"))
source(file.path(BTK_ROOT, "r", "tools", "xl_pricing_tool", "tool.R"))
source(file.path(BTK_ROOT, "r", "tools", "xl_pricing_tool", "gnpi.R"))
source(file.path(BTK_ROOT, "r", "tools", "xl_pricing_tool", "agg.R"))
source(file.path(BTK_ROOT, "r", "tools", "xl_pricing_tool", "profile.R"))
source(file.path(BTK_ROOT, "r", "tools", "xl_pricing_tool", "riskfit.R"))
source(file.path(BTK_ROOT, "r", "tools", "xl_pricing_tool", "cat_onlevel.R"))
source(file.path(BTK_ROOT, "r", "tools", "xl_pricing_tool", "xlsimulation.R"))
source(file.path(BTK_ROOT, "r", "tools", "xl_pricing_tool", "relativepricing.R"))
source(file.path(BTK_ROOT, "r", "tools", "xl_pricing_tool", "rmstoyelt.R"))
source(file.path(BTK_ROOT, "r", "tools", "curve_fit_risk", "tool.R"))
source(file.path(BTK_ROOT, "r", "tools", "goecode_tool", "tool.R"))

toolkit_version <- function() {
  "0.1.0"
}

toolkit_registry <- function() {
  list(
    china_exposure_map = list(
      actions = list(
        update = china_exposure_map_update
      ),
      run = china_exposure_map_update,
      validate = china_exposure_map_validate,
      name = "China Exposure Map"
    ),
    xl_pricing_tool = list(
      actions = list(
        gather = xl_pricing_tool_gather,
        update = xl_pricing_tool_update,
        build = xl_pricing_tool_build,
        gnpi_gather = xl_pricing_tool_gnpi_gather,
        gnpi_update = xl_pricing_tool_gnpi_update,
        gnpi_build = xl_pricing_tool_gnpi_build,
        agg_gather = xl_pricing_tool_agg_gather,
        agg_update = xl_pricing_tool_agg_update,
        agg_build = xl_pricing_tool_agg_build,
        profile_gather = xl_pricing_tool_profile_gather,
        profile_update = xl_pricing_tool_profile_update,
        profile_build = xl_pricing_tool_profile_build,
        riskfit_gather = xl_pricing_tool_riskfit_gather,
        riskfit_update = xl_pricing_tool_riskfit_update,
        riskfit_build = xl_pricing_tool_riskfit_build,
        cat_onlevel_gather = xl_pricing_tool_cat_onlevel_gather,
        cat_onlevel_update = xl_pricing_tool_cat_onlevel_update,
        cat_onlevel_build = xl_pricing_tool_cat_onlevel_build,
        xlsimulation_gather = xl_pricing_tool_xlsimulation_gather,
        xlsimulation_update = xl_pricing_tool_xlsimulation_update,
        xlsimulation_build = xl_pricing_tool_xlsimulation_build,
        relativepricing_gather = xl_pricing_tool_relativepricing_gather,
        relativepricing_update = xl_pricing_tool_relativepricing_update,
        relativepricing_build = xl_pricing_tool_relativepricing_build,
        rmstoyelt_gather = xl_pricing_tool_rmstoyelt_gather,
        rmstoyelt_update = xl_pricing_tool_rmstoyelt_update,
        rmstoyelt_build = xl_pricing_tool_rmstoyelt_build
      ),
      run = xl_pricing_tool_update,
      validate = xl_pricing_tool_validate,
      name = "XL Pricing Tool"
    ),
    curve_fit_risk = list(
      actions = list(),
      validate = curve_fit_risk_validate,
      name = "Curve Fit Risk"
    ),
    goecode_tool = list(
      actions = list(
        inputnorm_gather = goecode_tool_inputnorm_gather,
        inputnorm_update = goecode_tool_inputnorm_update,
        inputnorm_build = goecode_tool_inputnorm_build,
        histcross_gather = goecode_tool_histcross_gather,
        histcross_update = goecode_tool_histcross_update,
        histcross_build = goecode_tool_histcross_build
      ),
      run = goecode_tool_inputnorm_update,
      validate = goecode_tool_validate,
      name = "Geocode Tool"
    )
  )
}

toolkit_action_names <- function(tool_id) {
  registry <- toolkit_registry()
  if (!tool_id %in% names(registry)) {
    stop(sprintf("Unknown tool_id: %s", tool_id), call. = FALSE)
  }

  paste(names(registry[[tool_id]]$actions), collapse = ", ")
}

toolkit_validate <- function(tool_id, workbook_path, context = list()) {
  registry <- toolkit_registry()
  if (!tool_id %in% names(registry)) {
    stop(sprintf("Unknown tool_id: %s", tool_id), call. = FALSE)
  }

  if (length(context) == 0) {
    context <- btk_build_context(
      tool_id = tool_id,
      workbook_path = workbook_path,
      create_output_dir = FALSE
    )
  }

  registry[[tool_id]]$validate(workbook_path = workbook_path, context = context)
}

toolkit_run <- function(tool_id, workbook_path, output_dir = NULL, context = list()) {
  toolkit_dispatch(
    tool_id = tool_id,
    action = "update",
    workbook_path = workbook_path,
    output_dir = output_dir,
    context = context
  )
}

toolkit_dispatch <- function(tool_id, action, workbook_path, output_dir = NULL, context = list()) {
  registry <- toolkit_registry()
  if (!tool_id %in% names(registry)) {
    stop(sprintf("Unknown tool_id: %s", tool_id), call. = FALSE)
  }

  action <- tolower(trimws(action))
  actions <- registry[[tool_id]]$actions
  if (!action %in% names(actions)) {
    supported <- names(actions)
    supported_text <- if (length(supported) == 0) {
      "none yet"
    } else {
      paste(supported, collapse = ", ")
    }

    return(sprintf(
      "Unsupported action '%s' for %s. Supported actions: %s",
      action,
      registry[[tool_id]]$name,
      supported_text
    ))
  }

  if (length(context) == 0) {
    context <- btk_build_context(
      tool_id = tool_id,
      workbook_path = workbook_path,
      output_dir = output_dir
    )
  }

  if (identical(tool_id, "xl_pricing_tool")) {
    xl_pricing_tool_validate_action(
      workbook_path = context$workbook_path,
      action = action,
      context = context
    )
  }

  btk_log_message(
    sprintf("Running action '%s' for tool '%s'", action, tool_id),
    sprintf("Workbook: %s", context$workbook_path),
    sprintf("Output: %s", context$output_dir),
    log_file = context$log_file
  )

  actions[[action]](
    workbook_path = context$workbook_path,
    output_dir = context$output_dir,
    context = context
  )
}
