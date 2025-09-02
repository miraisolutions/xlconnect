test_that("test.workbook.getLastColumn", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookGetLastRow.xls", create = FALSE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookGetLastRow.xlsx", create = FALSE)

  # Check if last column is determined correctly (*.xls)
  expect_equal(11, as.vector(getLastColumn(wb.xls, "mtcars")))
  expect_equal(15, as.vector(getLastColumn(wb.xls, "mtcars2")))
  expect_equal(19, as.vector(getLastColumn(wb.xls, "mtcars3")))

  # Check if last column is determined correctly (*.xlsx)
  expect_equal(11, as.vector(getLastColumn(wb.xlsx, "mtcars")))
  expect_equal(15, as.vector(getLastColumn(wb.xlsx, "mtcars2")))
  expect_equal(19, as.vector(getLastColumn(wb.xlsx, "mtcars3")))

  # Check that querying the last column of a non-existing worksheet throws an exception (*.xls)
  expect_error(getLastColumn(wb.xls, "doesNotExist"))

  # Check that querying the last column of a non-existing worksheet throws an exception (*.xlsx)
  expect_error(getLastColumn(wb.xlsx, "doesNotExist"))

  # Last column of an empty worksheet is 1 (.xls)
  expect_equal(c(empty = 1), getLastColumn(wb.xls, "empty"))

  # Last column of an empty worksheet is 1 (.xlsx)
  expect_equal(c(empty = 1), getLastColumn(wb.xlsx, "empty"))
})
