test_that("test.workbook.getLastColumn", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookGetLastRow.xls", create = FALSE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookGetLastRow.xlsx", create = FALSE)

  # Check if last column is determined correctly (*.xls)
  expect_equal(as.vector(getLastColumn(wb.xls, "mtcars")), 11)
  expect_equal(as.vector(getLastColumn(wb.xls, "mtcars2")), 15)
  expect_equal(as.vector(getLastColumn(wb.xls, "mtcars3")), 19)

  # Check if last column is determined correctly (*.xlsx)
  expect_equal(as.vector(getLastColumn(wb.xlsx, "mtcars")), 11)
  expect_equal(as.vector(getLastColumn(wb.xlsx, "mtcars2")), 15)
  expect_equal(as.vector(getLastColumn(wb.xlsx, "mtcars3")), 19)

  # Check that querying the last column of a non-existing worksheet throws an exception (*.xls)
  expect_error(getLastColumn(wb.xls, "doesNotExist"))

  # Check that querying the last column of a non-existing worksheet throws an exception (*.xlsx)
  expect_error(getLastColumn(wb.xlsx, "doesNotExist"))

  # Last column of an empty worksheet is 1 (.xls)
  expect_equal(getLastColumn(wb.xls, "empty"), c(empty = 1))

  # Last column of an empty worksheet is 1 (.xlsx)
  expect_equal(getLastColumn(wb.xlsx, "empty"), c(empty = 1))
})
