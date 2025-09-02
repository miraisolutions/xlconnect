test_that("querying the last (non-empty) row of a worksheet", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookGetLastRow.xls", create = FALSE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookGetLastRow.xlsx", create = FALSE)

  # Check if last row is determined correctly (*.xls)
  expect_equal(as.vector(getLastRow(wb.xls, "mtcars")), 33)
  expect_equal(as.vector(getLastRow(wb.xls, "mtcars2")), 41)
  expect_equal(as.vector(getLastRow(wb.xls, "mtcars3")), 49)

  # Check if last row is determined correctly (*.xlsx)
  expect_equal(as.vector(getLastRow(wb.xlsx, "mtcars")), 33)
  expect_equal(as.vector(getLastRow(wb.xlsx, "mtcars2")), 41)
  expect_equal(as.vector(getLastRow(wb.xlsx, "mtcars3")), 49)

  # Check that querying the last row of a non-existing worksheet throws an exception (*.xls)
  expect_error(getLastRow(wb.xls, "doesNotExist"))

  # Check that querying the last row of a non-existing worksheet throws an exception (*.xlsx)
  expect_error(getLastRow(wb.xlsx, "doesNotExist"))

  # Last row of an empty worksheet is 1 (.xls)
  expect_equal(getLastRow(wb.xls, "empty"), c(empty = 1))

  # Last row of an empty worksheet is 1 (.xlsx)
  expect_equal(getLastRow(wb.xlsx, "empty"), c(empty = 1))
})
