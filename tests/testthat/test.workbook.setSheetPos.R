test_that("setting Excel worksheet positions", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookSetSheetPos.xls", create = TRUE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookSetSheetPos.xlsx", create = TRUE)

  createSheet(wb.xls, c("A", "B", "C", "D"))
  createSheet(wb.xlsx, c("A", "B", "C", "D"))

  # Check positions (*.xls)
  setSheetPos(wb.xls, sheet = c("D", "B"), pos = c(2, 1))
  expect_equal(getSheets(wb.xls), c("B", "A", "D", "C"))
  expect_equal(getSheetPos(wb.xls, sheet = c("A", "B", "C", "D")), c(A = 2, B = 1, C = 4, D = 3))

  # Check positions (*.xlsx)
  setSheetPos(wb.xlsx, sheet = c("D", "B"), pos = c(2, 1))
  expect_equal(getSheets(wb.xlsx), c("B", "A", "D", "C"))
  expect_equal(getSheetPos(wb.xlsx, sheet = c("A", "B", "C", "D")), c(A = 2, B = 1, C = 4, D = 3))

  # Check that trying to set a non-existing index (out of bounds) results in an exception (*.xls)
  expect_error(setSheetPos(wb.xls, sheet = "A", pos = -1))
  expect_error(setSheetPos(wb.xls, sheet = "A", pos = 36298))

  # Check that trying to set a non-existing position (out of bounds) results in an exception (*.xlsx)
  expect_error(setSheetPos(wb.xlsx, sheet = "A", pos = -1))
  expect_error(setSheetPos(wb.xlsx, sheet = "A", pos = 36298))

  # Check that trying to set the position of a non-existing worksheet results in an exception (*.xls)
  expect_error(setSheetPos(wb.xls, sheet = "NotThere", pos = 2))

  # Check that trying to set the position of a non-existing worksheet results in an exception (*.xlsx)
  expect_error(setSheetPos(wb.xlsx, sheet = "NotThere", pos = 2))
})
