test_that("test.workbook.setActiveSheet", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookSetActiveSheet.xls", create = FALSE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookSetActiveSheet.xlsx", create = FALSE)

  # Check that setting the active sheet works ok (*.xls)
  # (assumes that 'getActiveSheetIndex' works fine)
  setActiveSheet(wb.xls, 1)
  expect_true(getActiveSheetIndex(wb.xls) == 1)
  setActiveSheet(wb.xls, 3)
  expect_true(getActiveSheetIndex(wb.xls) == 3)
  setActiveSheet(wb.xls, "Sheet2")
  expect_true(getActiveSheetIndex(wb.xls) == 2)

  # Check that setting the active sheet works ok (*.xlsx)
  # (assumes that 'getActiveSheetIndex' works fine)
  setActiveSheet(wb.xlsx, 1)
  expect_true(getActiveSheetIndex(wb.xlsx) == 1)
  setActiveSheet(wb.xlsx, 3)
  expect_true(getActiveSheetIndex(wb.xlsx) == 3)
  setActiveSheet(wb.xlsx, "Sheet2")
  expect_true(getActiveSheetIndex(wb.xlsx) == 2)

  # Check that setting an illegal active sheet throws an exception (*.xls)
  expect_error(setActiveSheet(wb.xls, 19), "IllegalArgumentException")
  expect_error(setActiveSheet(wb.xls, "SheetWhichDoesNotExist"), "IllegalArgumentException")

  # Check that setting an illegal active sheet throws an exception (*.xlsx)
  expect_error(setActiveSheet(wb.xlsx, 19), "IllegalArgumentException")
  expect_error(setActiveSheet(wb.xlsx, "SheetWhichDoesNotExist"), "IllegalArgumentException")
})
