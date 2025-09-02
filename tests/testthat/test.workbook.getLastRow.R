test_that("querying the last (non-empty) row of a worksheet", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookGetLastRow.xls", create = FALSE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookGetLastRow.xlsx", create = FALSE)

  # Check if last row is determined correctly (*.xls)
  expect_equal(33, as.vector(getLastRow(wb.xls, "mtcars")))
  expect_equal(41, as.vector(getLastRow(wb.xls, "mtcars2")))
  expect_equal(49, as.vector(getLastRow(wb.xls, "mtcars3")))

  # Check if last row is determined correctly (*.xlsx)
  expect_equal(33, as.vector(getLastRow(wb.xlsx, "mtcars")))
  expect_equal(41, as.vector(getLastRow(wb.xlsx, "mtcars2")))
  expect_equal(49, as.vector(getLastRow(wb.xlsx, "mtcars3")))

  # Check that querying the last row of a non-existing worksheet throws an exception (*.xls)
  expect_error(getLastRow(wb.xls, "doesNotExist"))

  # Check that querying the last row of a non-existing worksheet throws an exception (*.xlsx)
  expect_error(getLastRow(wb.xlsx, "doesNotExist"))

  # Last row of an empty worksheet is 1 (.xls)
  expect_equal(c(empty = 1), getLastRow(wb.xls, "empty"))

  # Last row of an empty worksheet is 1 (.xlsx)
  expect_equal(c(empty = 1), getLastRow(wb.xlsx, "empty"))
})
