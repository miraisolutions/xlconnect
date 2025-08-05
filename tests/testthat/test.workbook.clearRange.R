test_that("clearRange works correctly for XLS", {
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
  clearRange(wb.xls, "clearRange", c(c(4, 4, 5, 5), c(6, 7, 7, 8)))
  res <- readWorksheet(wb.xls, "clearRange", region = "C3:I8", header = TRUE)
  expect_equal(checkDf, res)
})

test_that("clearRange works correctly for XLSX", {
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
  clearRange(wb.xlsx, "clearRange", c(c(4, 4, 5, 5), c(6, 7, 7, 8)))
  res <- readWorksheet(wb.xlsx, "clearRange", region = "C3:I8", header = TRUE)
  expect_equal(checkDf, res)
})
