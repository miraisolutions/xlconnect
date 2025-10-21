test_that("Tests around querying Excel worksheets", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookSheets.xls", create = FALSE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookSheets.xlsx", create = FALSE)

  # Sheets defined in workbooks
  expectedSheets <- c("A1", "B 2", "$$", "=", "@}", "11. Oct.", "\"quote\"", "+0")

  # Check that all and only the expected sheets exist (*.xls)
  definedSheets <- getSheets(wb.xls)
  expect_true(
    length(setdiff(expectedSheets, definedSheets)) == 0 && length(setdiff(definedSheets, expectedSheets)) == 0
  )

  # Check that all and only the expected sheets exist (*.xlsx)
  definedSheets <- getSheets(wb.xlsx)
  expect_true(
    length(setdiff(expectedSheets, definedSheets)) == 0 && length(setdiff(definedSheets, expectedSheets)) == 0
  )
})
