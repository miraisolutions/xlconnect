test_that("querying Excel table works as expected (*.xlsx) with simplify = TRUE", {
  # Create workbooks
  wb.xlsx <- loadWorkbook("resources/testWorkbookReadTable.xlsx", create = FALSE)

  # Check that querying Excel table works as expected (*.xlsx)
  res <- getTables(wb.xlsx, sheet = "Test", simplify = TRUE)
  expect_equal(res, "TestTable1")
})

test_that("querying Excel table works as expected (*.xlsx) with simplify = FALSE", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookReadTable.xlsx", create = FALSE)

  res <- getTables(wb.xlsx, sheet = "Test", simplify = FALSE)
  expect_equal(res, list(Test = "TestTable1"))
})

test_that("getTables returns an empty list for a sheet with no tables", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookReadTable.xlsx", create = FALSE)

  res <- getTables(wb.xlsx, sheet = "NoTableHere", simplify = TRUE)
  expect_equal(res, character(0))
})

test_that("trying to query tables from a non-existent sheet throws an exception (*.xlsx)", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookReadTable.xlsx", create = FALSE)

  # Check that trying to query tables from an non-existent sheet throws an exception (*.xlsx)
  expect_error(getTables(wb.xlsx, sheet = "DoesNotExist"))
})
