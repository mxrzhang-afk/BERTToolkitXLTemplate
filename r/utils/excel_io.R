btk_read_sheet <- function(workbook_path, sheet) {
  btk_require_packages(c("readxl"))
  readxl::read_excel(workbook_path, sheet = sheet)
}

btk_write_status <- function(workbook_path, sheet = "Log", rows = data.frame()) {
  btk_require_packages(c("openxlsx"))

  wb <- openxlsx::loadWorkbook(workbook_path)
  if (!sheet %in% names(wb)) {
    openxlsx::addWorksheet(wb, sheet)
  }
  openxlsx::writeData(wb, sheet = sheet, x = rows, startRow = 1, startCol = 1)
  openxlsx::saveWorkbook(wb, workbook_path, overwrite = TRUE)
  invisible(workbook_path)
}
