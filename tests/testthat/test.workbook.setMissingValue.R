test_that("writing and reading data with default missing value behaviour returns original data (XLS)", {
  wb.xls <- loadWorkbook("resources/testWorkbookSetMissingValue.xls", create = TRUE)
  data <- data.frame(A = c(4.2, -3.2, NA, 1.34), B = c("A", NA, "C", "D"), stringsAsFactors = FALSE)
  name <- "missing"
  createSheet(wb.xls, name = name)
  createName(wb.xls, name = name, formula = paste(name, "$A$1", sep = "!"))

  # Check that writing and reading the data with the default missing value behaviour returns
  # the original data (*.xls)
  writeNamedRegion(wb.xls, data, name = name)
  res <- readNamedRegion(wb.xls, name = name)
  attr(data, "worksheetScope") <- ""
  expect_equal(res, data)

  expect <- data.frame(
    A = c("4.2", "-3.2", "missing", "1.34"),
    B = c("A", "missing", "C", "D"),
    stringsAsFactors = FALSE
  )
  attr(expect, "worksheetScope") <- ""
  # Check that writing and reading the data with a specific missing value string
  # returns the original data but with the numeric column as a character and corresonding
  # missing value (*.xls)
  setMissingValue(wb.xls, value = "missing")
  writeNamedRegion(wb.xls, data, name = name)
  # Reset missing value string such that missing value string is read as string
  setMissingValue(wb.xls, value = NULL)
  res <- readNamedRegion(wb.xls, name = name)
  expect_equal(res, expect)
})

test_that("writing and reading data with default missing value behaviour returns original data (XLSX)", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookSetMissingValue.xlsx", create = TRUE)
  data <- data.frame(A = c(4.2, -3.2, NA, 1.34), B = c("A", NA, "C", "D"), stringsAsFactors = FALSE)
  name <- "missing"
  createSheet(wb.xlsx, name = name)
  createName(wb.xlsx, name = name, formula = paste(name, "$A$1", sep = "!"))

  # Check that writing and reading the data with the default missing value behaviour returns
  # the original data (*.xlsx)
  writeNamedRegion(wb.xlsx, data, name = name)
  res <- readNamedRegion(wb.xlsx, name = name)
  attr(data, "worksheetScope") <- ""
  expect_equal(res, data)

  expect <- data.frame(
    A = c("4.2", "-3.2", "missing", "1.34"),
    B = c("A", "missing", "C", "D"),
    stringsAsFactors = FALSE
  )
  attr(expect, "worksheetScope") <- ""
  # Check that writing and reading the data with a specific missing value string
  # returns the original data but with the numeric column as a character and corresonding
  # missing value (*.xlsx)
  setMissingValue(wb.xlsx, value = "missing")
  writeNamedRegion(wb.xlsx, data, name = name)
  # Reset missing value string such that missing value string is read as string
  setMissingValue(wb.xlsx, value = NULL)
  res <- readNamedRegion(wb.xlsx, name = name)
  expect_equal(res, expect)
})

test_that("reading data with multiple missing value strings works (XLS)", {
  wb.xls <- loadWorkbook("resources/testWorkbookMissingValue.xls", create = FALSE)
  expect <- data.frame(
    A = c(NA, -3.2, 3.4, NA, 8, NA),
    B = c("a", NA, "c", "x", "a", "o"),
    C = c(TRUE, TRUE, FALSE, NA, FALSE, NA),
    D = as.POSIXct(c("1981-12-01 00:00:00", "1981-12-02 00:00:00", NA, NA, NA, "1981-12-06 00:00:00")),
    stringsAsFactors = FALSE
  )
  # Check that reading data with multiple missing value strings works (*.xls)
  setMissingValue(wb.xls, value = c("NA", "missing", "empty"))
  res <- readNamedRegion(wb.xls, name = "Missing1")
  attr(expect, "worksheetScope") <- ""
  expect_equal(res, expect)
})

test_that("writing and reading data with same specific missing value string returns original data (XLS)", {
  wb.xls <- loadWorkbook("resources/testWorkbookSetMissingValue.xls", create = TRUE)
  data <- data.frame(A = c(4.2, -3.2, NA, 1.34), B = c("A", NA, "C", "D"), stringsAsFactors = FALSE)
  name <- "missing"
  createSheet(wb.xls, name = name)
  createName(wb.xls, name = name, formula = paste(name, "$A$1", sep = "!"))

  # Check that writing and reading the data with a specific missing value string
  # returns the original data (*.xls)
  setMissingValue(wb.xls, value = "missing")
  writeNamedRegion(wb.xls, data, name = name)
  res <- readNamedRegion(wb.xls, name = name)
  attr(data, "worksheetScope") <- ""
  expect_equal(res, data)
})

test_that("writing and reading data with same specific missing value string returns original data (XLSX)", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookSetMissingValue.xlsx", create = TRUE)
  data <- data.frame(A = c(4.2, -3.2, NA, 1.34), B = c("A", NA, "C", "D"), stringsAsFactors = FALSE)
  name <- "missing"
  createSheet(wb.xlsx, name = name)
  createName(wb.xlsx, name = name, formula = paste(name, "$A$1", sep = "!"))

  # Check that writing and reading the data with a specific missing value string
  # returns the original data (*.xlsx)
  setMissingValue(wb.xlsx, value = "missing")
  writeNamedRegion(wb.xlsx, data, name = name)
  res <- readNamedRegion(wb.xlsx, name = name)
  attr(data, "worksheetScope") <- ""
  expect_equal(res, data)
})

test_that("reading data with multiple missing value strings works (XLSX)", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookMissingValue.xlsx", create = FALSE)
  expect <- data.frame(
    A = c(NA, -3.2, 3.4, NA, 8, NA),
    B = c("a", NA, "c", "x", "a", "o"),
    C = c(TRUE, TRUE, FALSE, NA, FALSE, NA),
    D = as.POSIXct(c("1981-12-01 00:00:00", "1981-12-02 00:00:00", NA, NA, NA, "1981-12-06 00:00:00")),
    stringsAsFactors = FALSE
  )
  # Check that reading data with multiple missing value strings works (*.xlsx)
  setMissingValue(wb.xlsx, value = c("NA", "missing", "empty"))
  res <- readNamedRegion(wb.xlsx, name = "Missing1")
  attr(expect, "worksheetScope") <- ""
  expect_equal(res, expect)
})

test_that("reading data with multiple missing value strings works with list input (XLS)", {
  wb.xls <- loadWorkbook("resources/testWorkbookMissingValue.xls", create = FALSE)
  expect <- data.frame(
    A = c(NA, -3.2, NA, NA, 8, NA),
    B = c("a", NA, "c", "x", "a", "o"),
    C = c(TRUE, NA, FALSE, NA, FALSE, NA),
    D = as.POSIXct(c("1981-12-01 00:00:00", "1981-12-02 00:00:00", NA, NA, NA, "1981-12-06 00:00:00")),
    stringsAsFactors = FALSE
  )
  # Check that reading data with multiple missing value strings works (*.xls)
  setMissingValue(wb.xls, value = list("NA", "missing", "empty", -9999))
  res <- readNamedRegion(wb.xls, name = "Missing2")
  attr(expect, "worksheetScope") <- ""
  expect_equal(res, expect)
})

test_that("reading data with multiple missing value strings works with list input (XLSX)", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookMissingValue.xlsx", create = FALSE)
  expect <- data.frame(
    A = c(NA, -3.2, NA, NA, 8, NA),
    B = c("a", NA, "c", "x", "a", "o"),
    C = c(TRUE, NA, FALSE, NA, FALSE, NA),
    D = as.POSIXct(c("1981-12-01 00:00:00", "1981-12-02 00:00:00", NA, NA, NA, "1981-12-06 00:00:00")),
    stringsAsFactors = FALSE
  )
  # Check that reading data with multiple missing value strings works (*.xlsx)
  setMissingValue(wb.xlsx, value = list("NA", "missing", "empty", -9999))
  res <- readNamedRegion(wb.xlsx, name = "Missing2")
  attr(expect, "worksheetScope") <- ""
  expect_equal(res, expect)
})
