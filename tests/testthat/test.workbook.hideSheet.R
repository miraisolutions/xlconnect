test_that("Tests around hiding Excel worksheets", {
  wb.xls <- loadWorkbook("resources/testWorkbookHideSheet.xls", create = TRUE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookHideSheet.xlsx", create = TRUE)
  visibleSheet <- "VisibleSheet"
  hiddenSheet <- "HiddenSheet"
  veryHiddenSheet <- "VeryHiddenSheet"
  createSheet(wb.xls, visibleSheet)
  createSheet(wb.xlsx, visibleSheet)
  createSheet(wb.xls, hiddenSheet)
  createSheet(wb.xlsx, hiddenSheet)
  createSheet(wb.xls, veryHiddenSheet)
  createSheet(wb.xlsx, veryHiddenSheet)

  # Check if sheets are hidden correspondingly (*.xls)
  hideSheet(wb.xls, hiddenSheet)
  expect_true(isSheetHidden(wb.xls, hiddenSheet))
  hideSheet(wb.xls, veryHiddenSheet, veryHidden = TRUE)
  expect_true(isSheetVeryHidden(wb.xls, veryHiddenSheet))

  # Check if sheets are hidden correspondingly (*.xlsx)
  hideSheet(wb.xlsx, hiddenSheet)
  expect_true(isSheetHidden(wb.xlsx, hiddenSheet))
  hideSheet(wb.xlsx, veryHiddenSheet, veryHidden = TRUE)
  expect_true(isSheetVeryHidden(wb.xlsx, veryHiddenSheet))
})
