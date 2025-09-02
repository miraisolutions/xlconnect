test_that("Tests around querying the active sheet index of an Excel workbook", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookActiveSheetIndexAndName.xls", create = FALSE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookActiveSheetIndexAndName.xlsx", create = FALSE)

  # Check that the active sheet index is 5 (*.xls)
  expect_true(getActiveSheetIndex(wb.xls) == 5)

  # Check that the active sheet index is 5 (*.xlsx)
  expect_true(getActiveSheetIndex(wb.xlsx) == 5)
})
