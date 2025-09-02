test_that("createSheet throws an error for invalid sheet names in XLS", {
  wb.xls <- loadWorkbook("resources/createSheet.xls", create = TRUE)

  # Check that an exception is thrown when trying to create
  # a worksheet with a single quote at the beginning (*.xls)
  expect_error(createSheet(wb.xls, "'Invalid Sheet Name"), "IllegalArgumentException")

  # Check that an exception is thrown when trying to create
  # a worksheet with a single quote at the end (*.xls)
  expect_error(createSheet(wb.xls, "Invalid Sheet Name'"), "IllegalArgumentException")

  # Check that an exception is thrown when trying to create
  # a worksheet with a very long name (> 31 characters) (*.xls)
  expect_error(createSheet(wb.xls, "A very very very very very very very very long name"), "IllegalArgumentException")
})

test_that("createSheet throws an error for invalid sheet names in XLSX", {
  wb.xlsx <- loadWorkbook("resources/createSheet.xlsx", create = TRUE)

  # Check that an exception is thrown when trying to create
  # a worksheet with a single quote at the beginning (*.xlsx)
  expect_error(createSheet(wb.xlsx, "'Invalid Sheet Name"), "IllegalArgumentException")

  # Check that an exception is thrown when trying to create
  # a worksheet with a single quote at the end (*.xlsx)
  expect_error(createSheet(wb.xlsx, "Invalid Sheet Name'"), "IllegalArgumentException")

  # Check that an exception is thrown when trying to create
  # a worksheet with a very long name (> 31 characters) (*.xlsx)
  expect_error(createSheet(wb.xlsx, "A very very very very very very very very long name"), "IllegalArgumentException")
})

test_that("createSheet works for valid sheet names in XLS and XLSX", {
  wb.xls <- loadWorkbook("resources/createSheet.xls", create = TRUE)
  wb.xlsx <- loadWorkbook("resources/createSheet.xlsx", create = TRUE)
  sheetName <- "My Sheet"

  # Check if creating a legal worksheet is working properly (*.xls)
  # (assumes method existsSheet working properly)
  createSheet(wb.xls, sheetName)
  expect_true(existsSheet(wb.xls, sheetName))

  # Check if creating a legal worksheet is working properly (*.xlsx)
  # (assumes method existsSheet working properly)
  createSheet(wb.xlsx, sheetName)
  expect_true(existsSheet(wb.xlsx, sheetName))
})

test_that("creating an existing sheet is a NOP in XLS and XLSX", {
  wb.xls <- loadWorkbook("resources/createSheet.xls", create = TRUE)
  wb.xlsx <- loadWorkbook("resources/createSheet.xlsx", create = TRUE)
  sheetName <- "My Sheet"

  # Trying to create an already existing sheet should not cause
  # any issues (just skips) (*.xls)
  createSheet(wb.xls, sheetName) # First time
  expect_true(existsSheet(wb.xls, sheetName))
  createSheet(wb.xls, sheetName) # Second time
  expect_true(existsSheet(wb.xls, sheetName))

  # Trying to create an already existing sheet should not cause
  # any issues (just skips) (*.xlsx)
  createSheet(wb.xlsx, sheetName) # First time
  expect_true(existsSheet(wb.xlsx, sheetName))
  createSheet(wb.xlsx, sheetName) # Second time
  expect_true(existsSheet(wb.xlsx, sheetName))
})
