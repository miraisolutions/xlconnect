test_that("clearing ranges from references in an Excel Workbook works correctly for XLS", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookClearCells.xls", create = FALSE)

  checkDf <- data.frame(
    one = 1:5,
    two = c(NA, NA, 8, 9, 10),
    three = c(NA, NA, 13, 14, 15),
    four = 16:20,
    five = c(21, 22, NA, NA, 25),
    six = c(26, 27, NA, NA, 30),
    seven = 31:35,
    stringsAsFactors = F
  )

  # Check that clearing ranges from references returns the desired result (*.xls)
  clearRangeFromReference(wb.xls, c("clearRangeFromReference!D4:E5", "clearRangeFromReference!G6:H7"))
  res <- readWorksheet(wb.xls, "clearRangeFromReference", region = "C3:I8", header = TRUE)
  expect_equal(res, checkDf)
})

test_that("clearing ranges from references in an Excel Workbook works correctly for XLSX", {
  # Create workbooks
  wb.xlsx <- loadWorkbook("resources/testWorkbookClearCells.xlsx", create = FALSE)

  checkDf <- data.frame(
    one = 1:5,
    two = c(NA, NA, 8, 9, 10),
    three = c(NA, NA, 13, 14, 15),
    four = 16:20,
    five = c(21, 22, NA, NA, 25),
    six = c(26, 27, NA, NA, 30),
    seven = 31:35,
    stringsAsFactors = F
  )

  # Check that clearing ranges from references returns the desired result (*.xlsx)
  clearRangeFromReference(wb.xlsx, c("clearRangeFromReference!D4:E5", "clearRangeFromReference!G6:H7"))
  res <- readWorksheet(wb.xlsx, "clearRangeFromReference", region = "C3:I8", header = TRUE)
  expect_equal(res, checkDf)
})
