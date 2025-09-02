test_that("checking equality of data.frame's being written to and read from Excel worksheets - always run", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWriteAndReadWorksheet.xls", create = TRUE)
  wb.xlsx <- loadWorkbook("resources/testWriteAndReadWorksheet.xlsx", create = TRUE)
  testDataFrame <- function(wb, df, startRow, startCol) {
    worksheet <- deparse(substitute(df))
    createSheet(wb, worksheet)
    writeWorksheet(wb, df, worksheet, startRow = startRow, startCol = startCol)
    res <- readWorksheet(wb, worksheet, startRow = startRow, startCol = startCol, endRow = 0, endCol = 0)
    expect_equal(normalizeDataframe(df), res, ignore_attr = c("worksheetScope"))
  }

  # custom test dataset
  cdf <- data.frame(
    Column.A = c(1, 2, 3, NA, 5, 6, 7, 8, NA, 10),
    Column.B = c(-4, -3, NA, -1, 0, NA, NA, 3, 4, 5),
    Column.C = c("Anna", "???", NA, "", NA, "$!?&%", "(?2@?~?'^*#|)", "{}[]:,;-_<>", "\\sadf\n\nv", "a b c"),
    Column.D = c(pi, -pi, NA, sqrt(2), sqrt(0.3), -sqrt(pi), exp(1), log(2), sin(2), -tan(2)),
    Column.E = c(TRUE, TRUE, NA, NA, FALSE, FALSE, TRUE, NA, FALSE, TRUE),
    Column.F = c("High", "Medium", "Low", "Low", "Low", NA, NA, "Medium", "High", "High"),
    Column.G = c("High", "Medium", NA, "Low", "Low", "Medium", NA, "Medium", "High", "High"),
    Column.H = rep(c(as.Date("2021-10-30"), as.Date("2021-03-28"), NA), length = 10),
    # NOTE: Column.I is automatically converted to POSIXct!!!
    Column.I = rep(
      c(
        as.POSIXlt("2021-10-31 03:00:00"),
        as.POSIXlt(1582963631, origin = "1970-01-01"),
        NA,
        as.POSIXlt("2001-12-31 23:59:59")
      ),
      length = 10
    ),
    # NOTE: 1582963631 with origin="1970-01-01" corresponds to 2020 Feb 29
    Column.J = rep(
      c(
        as.POSIXct("2021-10-31 03:00:00"),
        as.POSIXct(1582963631, origin = "1970-01-01"),
        NA,
        as.POSIXct("2001-12-31 23:59:59")
      ),
      length = 10
    ),
    stringsAsFactors = F
  )
  cdf[["Column.F"]] <- factor(cdf[["Column.F"]])
  cdf[["Column.F"]] <- ordered(cdf[["Column.F"]], levels = c("Low", "Medium", "High"))

  # (*.xls)
  testDataFrame(wb.xls, cdf, 1, 1)
  # (*.xlsx)
  testDataFrame(wb.xlsx, cdf, 1, 1)

  # Check writing of data.frame with row names (*.xls)
  createSheet(wb.xls, "rownames")
  writeWorksheet(wb.xls, mtcars, "rownames", startRow = 9, startCol = 10, header = TRUE, rownames = "Car")
  res <- readWorksheet(wb.xls, "rownames", startRow = 9, startCol = 10)
  expect_equal(res, XLConnect:::includeRownames(mtcars, "Car"))

  # Check writing of data.frame with row names (*.xlsx)
  createSheet(wb.xlsx, "rownames")
  writeWorksheet(wb.xlsx, mtcars, "rownames", startRow = 9, startCol = 10, header = TRUE, rownames = "Car")
  res <- readWorksheet(wb.xlsx, "rownames", startRow = 9, startCol = 10)
  expect_equal(res, XLConnect:::includeRownames(mtcars, "Car"))

  # Check writing & reading of data.frame with row names (*.xls)
  createSheet(wb.xls, name = "rownames2")
  writeWorksheet(wb.xls, mtcars, "rownames2", startRow = 31, startCol = 8, header = TRUE, rownames = "Car")
  res <- readWorksheet(wb.xls, "rownames2", rownames = "Car")
  expect_equal(res, mtcars)
  expect_equal(attr(res, "row.names"), attr(mtcars, "row.names"))

  # Check writing & reading of data.frame with row names (*.xlsx)
  createSheet(wb.xlsx, name = "rownames2")
  writeWorksheet(wb.xlsx, mtcars, "rownames2", startRow = 31, startCol = 8, header = TRUE, rownames = "Car")
  res <- readWorksheet(wb.xlsx, "rownames2", rownames = "Car")
  expect_equal(res, mtcars)
  expect_equal(attr(res, "row.names"), attr(mtcars, "row.names"))
})

test_that("checking equality of data.frame's being written to and read from Excel worksheets - full test suite only", {
  skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")

  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWriteAndReadWorksheet.xls", create = TRUE)
  wb.xlsx <- loadWorkbook("resources/testWriteAndReadWorksheet.xlsx", create = TRUE)
  testDataFrame <- function(wb, df, startRow, startCol) {
    worksheet <- deparse(substitute(df))
    createSheet(wb, worksheet)
    writeWorksheet(wb, df, worksheet, startRow = startRow, startCol = startCol)
    res <- readWorksheet(wb, worksheet, startRow = startRow, startCol = startCol, endRow = 0, endCol = 0)
    expect_equal(normalizeDataframe(df), res, ignore_attr = c("worksheetScope", "row.names"))
  }

  # built-in dataset mtcars (*.xls)
  testDataFrame(wb.xls, mtcars, 8, 13)
  # built-in dataset mtcars (*.xlsx)
  testDataFrame(wb.xlsx, mtcars, 8, 13)

  # built-in dataset airquality (*.xls)
  testDataFrame(wb.xls, airquality, 2, 9)
  # built-in dataset airquality (*.xlsx)
  testDataFrame(wb.xlsx, airquality, 2, 9)

  # built-in dataset attenu (*.xls)
  testDataFrame(wb.xls, attenu, 7, 1)
  # built-in dataset attenu (*.xlsx)
  testDataFrame(wb.xlsx, attenu, 7, 1)

  # built-in dataset ChickWeight (*.xls)
  testDataFrame(wb.xls, ChickWeight, 8, 8)
  # built-in dataset ChickWeight (*.xlsx)
  testDataFrame(wb.xlsx, ChickWeight, 8, 8)

  # built-in dataset CO2 (*.xls)
  testDataFrame(wb.xls, CO2, 100, 12)
  # built-in dataset CO2 (*.xlsx)
  testDataFrame(wb.xlsx, CO2, 100, 12)

  # built-in dataset iris (*.xls)
  testDataFrame(wb.xls, iris, 1, 1)
  # built-in dataset iris (*.xlsx)
  testDataFrame(wb.xlsx, iris, 1, 1)

  # built-in dataset longley (*.xls)
  testDataFrame(wb.xls, longley, 5, 214)
  # built-in dataset longley (*.xlsx)
  testDataFrame(wb.xlsx, longley, 5, 214)

  # built-in dataset morley (*.xls)
  testDataFrame(wb.xls, morley, 1000, 6)
  # built-in dataset morley (*.xlsx)
  testDataFrame(wb.xlsx, morley, 1000, 6)

  # built-in dataset swiss (*.xls)
  testDataFrame(wb.xls, swiss, 4, 4)
  # built-in dataset swiss (*.xlsx)
  testDataFrame(wb.xlsx, swiss, 4, 4)
})
