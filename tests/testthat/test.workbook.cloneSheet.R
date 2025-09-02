test_that("cloning Excel worksheets", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookCloneSheet.xls", create = FALSE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookCloneSheet.xlsx", create = FALSE)

  # Check cloning of worksheet (*.xls)
  cloneSheet(wb.xls, sheet = "Test1", name = "Clone1")
  readWorksheet(wb.xls, sheet = "Clone1")

  # Check cloning of worksheet (*.xlsx)
  cloneSheet(wb.xlsx, sheet = "Test1", name = "Clone1")
  readWorksheet(wb.xlsx, sheet = "Clone1")

  # Check that attempting to clone a non-existing worksheet throws an exception (*.xls)
  expect_error(cloneSheet(wb.xls, sheet = "NotThere", name = "MyClone"), "IllegalArgumentException")

  # Check that attempting to clone a non-existing worksheet throws an exception (*.xlsx)
  expect_error(cloneSheet(wb.xlsx, sheet = "NotThere", name = "MyClone"), "IllegalArgumentException")

  # Check that attempting to assign an invalid name to a cloned sheet throws an exception (*.xls)
  expect_error(cloneSheet(wb.xls, sheet = "Test1", name = "'illegal"), "IllegalArgumentException")

  # Check that attempting to assign an invalid name to a cloned sheet throws an exception (*.xlsx)
  expect_error(cloneSheet(wb.xlsx, sheet = "Test1", name = "'illegal"), "IllegalArgumentException")
})
