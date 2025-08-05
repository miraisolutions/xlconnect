test_that("createSheet throws an error for invalid sheet names in XLS", {
  wb.xls <- loadWorkbook("resources/createSheet.xls", create = TRUE)
  expect_error(createSheet(wb.xls, "'Invalid Sheet Name"), "IllegalArgumentException")
  expect_error(createSheet(wb.xls, "Invalid Sheet Name'"), "IllegalArgumentException")
  expect_error(createSheet(wb.xls, "A very very very very very very very very long name"), "IllegalArgumentException")
})

test_that("createSheet throws an error for invalid sheet names in XLSX", {
  wb.xlsx <- loadWorkbook("resources/createSheet.xlsx", create = TRUE)
  expect_error(createSheet(wb.xlsx, "'Invalid Sheet Name"), "IllegalArgumentException")
  expect_error(createSheet(wb.xlsx, "Invalid Sheet Name'"), "IllegalArgumentException")
  expect_error(createSheet(wb.xlsx, "A very very very very very very very very long name"), "IllegalArgumentException")
})

test_that("createSheet works for valid sheet names in XLS and XLSX", {
  wb.xls <- loadWorkbook("resources/createSheet.xls", create = TRUE)
  wb.xlsx <- loadWorkbook("resources/createSheet.xlsx", create = TRUE)
  sheetName <- "My Sheet"

  createSheet(wb.xls, sheetName)
  expect_true(existsSheet(wb.xls, sheetName))

  createSheet(wb.xlsx, sheetName)
  expect_true(existsSheet(wb.xlsx, sheetName))
})

test_that("creating an existing sheet is a NOP in XLS and XLSX", {
  wb.xls <- loadWorkbook("resources/createSheet.xls", create = TRUE)
  wb.xlsx <- loadWorkbook("resources/createSheet.xlsx", create = TRUE)
  sheetName <- "My Sheet"

  createSheet(wb.xls, sheetName) # First time
  expect_true(existsSheet(wb.xls, sheetName))
  createSheet(wb.xls, sheetName) # Second time
  expect_true(existsSheet(wb.xls, sheetName))

  createSheet(wb.xlsx, sheetName) # First time
  expect_true(existsSheet(wb.xlsx, sheetName))
  createSheet(wb.xlsx, sheetName) # Second time
  expect_true(existsSheet(wb.xlsx, sheetName))
})
