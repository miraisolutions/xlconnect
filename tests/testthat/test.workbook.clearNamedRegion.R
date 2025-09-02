test_that("test.workbook.clearNamedRegion", {
  wb.xls <- loadWorkbook("resources/testWorkbookClearCells.xls", create = FALSE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookClearCells.xlsx", create = FALSE)
  checkDf <- data.frame(
    one = 1:5,
    two = 6:10,
    three = 11:15,
    four = 16:20,
    five = 21:25,
    six = 26:30,
    seven = 31:35,
    stringsAsFactors = F
  )

  # Check that clearing 2 of 3 named regions in a sheet returns only the third one (*.xls)
  clearNamedRegion(wb.xls, c("region1", "region2"))
  res <- readWorksheet(wb.xls, "clearNamedRegion", header = TRUE)
  expect_equal(checkDf, res)

  # Check that clearing 2 of 3 named regions in a sheet returns only the third one (*.xlsx)
  clearNamedRegion(wb.xlsx, c("region1", "region2"))
  res <- readWorksheet(wb.xlsx, "clearNamedRegion", header = TRUE)
  expect_equal(checkDf, res)

  # reset
  wb.xls <- loadWorkbook("resources/testWorkbookClearCells.xls", create = FALSE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookClearCells.xlsx", create = FALSE)

  # check that specifying the worksheet name doesn't find globally scoped names
  expect_error(clearNamedRegion(wb.xls, c("region1", "region2"), worksheetScope = "clearNamedRegion"))

  # Check that clearing 2 of 3 named regions in a sheet returns only the third one (*.xls)
  clearNamedRegion(wb.xls, c("region1", "region2"))
  res <- readWorksheet(wb.xls, "clearNamedRegion", header = TRUE)
  expect_equal(checkDf, res)

  # Check that clearing 2 of 3 named regions in a sheet returns only the third one (*.xlsx)
  clearNamedRegion(wb.xlsx, c("region1", "region2"))
  res <- readWorksheet(wb.xlsx, "clearNamedRegion", header = TRUE)
  expect_equal(checkDf, res)
})
