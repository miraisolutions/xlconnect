test_that("test.workbook.getActiveSheetName", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookActiveSheetIndexAndName.xls", create = FALSE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookActiveSheetIndexAndName.xlsx", create = FALSE)

  # Check that the active sheet name is 'Fifth Sheet' (*.xls)
  expect_true(getActiveSheetName(wb.xls) == "Fifth Sheet")

  # Check that the active sheet name is 'Fifth Sheet' (*.xlsx)
  expect_true(getActiveSheetName(wb.xlsx) == "Fifth Sheet")
})
