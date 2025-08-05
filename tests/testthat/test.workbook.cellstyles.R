context("Workbook Cell Style Functionality")

test_that("Basic cell style operations (create, exists, get) work as expected", {
  skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE for cellstyle tests")

  file.xls <- withr::local_tempfile(fileext = ".xls")
  file.xlsx <- withr::local_tempfile(fileext = ".xlsx")

  wb.xls <- loadWorkbook(file.xls, create = TRUE)
  wb.xlsx <- loadWorkbook(file.xlsx, create = TRUE)

  styleName <- "MyStyle"

  # Initial state: style should not exist
  expect_false(existsCellStyle(wb.xls, styleName))
  expect_false(existsCellStyle(wb.xlsx, styleName))
  expect_error(getCellStyle(wb.xls, styleName), "IllegalArgumentException", info = "XLS: Getting non-existent style")
  expect_error(getCellStyle(wb.xlsx, styleName), "IllegalArgumentException", info = "XLSX: Getting non-existent style")

  # Create style
  createCellStyle(wb.xls, styleName)
  createCellStyle(wb.xlsx, styleName)
  expect_true(existsCellStyle(wb.xls, styleName), info = "XLS: Style should exist after creation")
  expect_true(existsCellStyle(wb.xlsx, styleName), info = "XLSX: Style should exist after creation")

  # Attempting to create an existing style should error
  expect_error(createCellStyle(wb.xls, styleName), "IllegalArgumentException", info = "XLS: Creating existing style")
  expect_error(createCellStyle(wb.xlsx, styleName), "IllegalArgumentException", info = "XLSX: Creating existing style")

  # Getting an existing style should work
  cs_xls <- getCellStyle(wb.xls, styleName)
  expect_true(is(cs_xls, "cellstyle"), info = "XLS: getCellStyle should return a cellstyle object")
  cs_xlsx <- getCellStyle(wb.xlsx, styleName)
  expect_true(is(cs_xlsx, "cellstyle"), info = "XLSX: getCellStyle should return a cellstyle object")
})

test_that("getOrCreateCellStyle works as expected", {
  skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE for cellstyle tests")

  file.xls <- withr::local_tempfile(fileext = ".xls")
  file.xlsx <- withr::local_tempfile(fileext = ".xlsx")

  wb.xls <- loadWorkbook(file.xls, create = TRUE)
  wb.xlsx <- loadWorkbook(file.xlsx, create = TRUE)

  anotherStyleName <- "MyOtherStyle"

  # Style does not exist initially
  expect_false(existsCellStyle(wb.xls, anotherStyleName))
  expect_false(existsCellStyle(wb.xlsx, anotherStyleName))

  # getOrCreateCellStyle when style does not exist
  cs_xls_new <- getOrCreateCellStyle(wb.xls, anotherStyleName)
  expect_true(is(cs_xls_new, "cellstyle"), info = "XLS: getOrCreate (new) should return cellstyle")
  expect_true(existsCellStyle(wb.xls, anotherStyleName), info = "XLS: Style should exist after getOrCreate (new)")

  cs_xlsx_new <- getOrCreateCellStyle(wb.xlsx, anotherStyleName)
  expect_true(is(cs_xlsx_new, "cellstyle"), info = "XLSX: getOrCreate (new) should return cellstyle")
  expect_true(existsCellStyle(wb.xlsx, anotherStyleName), info = "XLSX: Style should exist after getOrCreate (new)")

  # getOrCreateCellStyle when style already exists
  cs_xls_existing <- getOrCreateCellStyle(wb.xls, anotherStyleName)
  expect_true(is(cs_xls_existing, "cellstyle"), info = "XLS: getOrCreate (existing) should return cellstyle")

  cs_xlsx_existing <- getOrCreateCellStyle(wb.xlsx, anotherStyleName)
  expect_true(is(cs_xlsx_existing, "cellstyle"), info = "XLSX: getOrCreate (existing) should return cellstyle")
})
