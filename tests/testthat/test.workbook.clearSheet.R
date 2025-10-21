test_that("clearing sheets returns empty sheets (*.xls)", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookClearCells.xls", create = FALSE)

  # Check that clearing sheets returns empty sheets (*.xls)
  clearSheet(wb.xls, c("clearSheet1", "clearSheet2"))
  res1 <- getLastRow(wb.xls, "clearSheet1")
  res2 <- getLastRow(wb.xls, "clearSheet2")
  expect_equal(c(res1, res2), c(clearSheet1 = 1, clearSheet2 = 1))
})

test_that("clearing sheets returns empty sheets (*.xlsx)", {
  # Create workbooks
  wb.xlsx <- loadWorkbook("resources/testWorkbookClearCells.xlsx", create = FALSE)

  # Check that clearing sheets returns empty sheets (*.xlsx)
  clearSheet(wb.xlsx, c("clearSheet1", "clearSheet2"))
  res1 <- getLastRow(wb.xlsx, "clearSheet1")
  res2 <- getLastRow(wb.xlsx, "clearSheet2")
  expect_equal(c(res1, res2), c(clearSheet1 = 1, clearSheet2 = 1))
})
