test_that("checking equality of data.frame's being written to and read from Excel named regions - always run", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWriteAndReadNamedRegion.xls", create = TRUE)
  wb.xlsx <- loadWorkbook("resources/testWriteAndReadNamedRegion.xlsx", create = TRUE)
  testDataFrame <- function(wb, df, lref) {
    namedRegion <- deparse(substitute(df))
    createSheet(wb, name = namedRegion)
    createName(wb, name = namedRegion, formula = paste(namedRegion, lref, sep = "!"), worksheetScope = namedRegion)
    writeNamedRegion(wb, df, name = namedRegion, worksheetScope = namedRegion, header = TRUE)
    res <- readNamedRegion(wb, namedRegion, worksheetScope = namedRegion)
    expect_equal(res, normalizeDataframe(df, replaceInf = TRUE), ignore_attr = c("worksheetScope"))
  }
  # custom test dataset
  cdf <- data.frame(
    Column.A = c(1, 2, 3, NA, 5, Inf, 7, 8, NA, 10),
    Column.B = c(-4, -3, NA, -Inf, 0, NA, NA, 3, 4, 5),
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
  testDataFrame(wb.xls, cdf, "$X$100")
  # (*.xlsx)
  testDataFrame(wb.xlsx, cdf, "$X$100")

  # Check writing of data.frame with row names (*.xls)
  createSheet(wb.xls, name = "rownames")
  createName(wb.xls, name = "rownames", formula = "rownames!$F$16")
  writeNamedRegion(wb.xls, mtcars, name = "rownames", header = TRUE, rownames = "Car")
  res <- readNamedRegion(wb.xls, "rownames")
  expect_equal(res, XLConnect:::includeRownames(mtcars, "Car"), ignore_attr = c("worksheetScope"))

  # Check writing of data.frame with row names (*.xlsx)
  createSheet(wb.xlsx, name = "rownames")
  createName(wb.xlsx, name = "rownames", formula = "rownames!$F$16")
  writeNamedRegion(wb.xlsx, mtcars, name = "rownames", header = TRUE, rownames = "Car")
  res <- readNamedRegion(wb.xlsx, "rownames")
  expect_equal(res, XLConnect:::includeRownames(mtcars, "Car"), ignore_attr = c("worksheetScope"))

  # Check writing & reading of data.frame with row names (*.xls)
  createSheet(wb.xls, name = "rownames2")
  createName(wb.xls, name = "rownames2", formula = "rownames2!$K$5")
  writeNamedRegion(wb.xls, mtcars, name = "rownames2", header = TRUE, rownames = "Car")
  res <- readNamedRegion(wb.xls, "rownames2", rownames = "Car")
  expect_equal(res, mtcars)
  expect_equal(attr(res, "row.names"), attr(mtcars, "row.names"))

  # Check writing & reading of data.frame with row names (*.xlsx)
  createSheet(wb.xlsx, name = "rownames2")
  createName(wb.xlsx, name = "rownames2", formula = "rownames2!$K$5")
  writeNamedRegion(wb.xlsx, mtcars, name = "rownames2", header = TRUE, rownames = "Car")
  res <- readNamedRegion(wb.xlsx, "rownames2", rownames = "Car")
  expect_equal(res, mtcars)
  expect_equal(attr(res, "row.names"), attr(mtcars, "row.names"))
})

test_that("checking equality of data.frame's being written to and read from Excel named regions - full test suite only", {
  skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")

  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWriteAndReadNamedRegion.xls", create = TRUE)
  wb.xlsx <- loadWorkbook("resources/testWriteAndReadNamedRegion.xlsx", create = TRUE)
  testDataFrame <- function(wb, df, lref) {
    namedRegion <- deparse(substitute(df))
    createSheet(wb, name = namedRegion)
    createName(wb, name = namedRegion, formula = paste(namedRegion, lref, sep = "!"), worksheetScope = namedRegion)
    writeNamedRegion(wb, df, name = namedRegion, worksheetScope = namedRegion, header = TRUE)
    res <- readNamedRegion(wb, namedRegion, worksheetScope = namedRegion)
    expect_equal(res, normalizeDataframe(df, replaceInf = TRUE), ignore_attr = c("worksheetScope", "row.names"))
  }
  testDataFrameNameScope <- function(wb, df, lref) {
    namedRegion <- paste(deparse(substitute(df)), "1", sep = "")
    worksheetScopeName <- paste(namedRegion, "2", sep = "")
    createSheet(wb, name = namedRegion)
    createSheet(wb, name = worksheetScopeName)
    createName(
      wb,
      name = namedRegion,
      formula = paste(namedRegion, lref, sep = "!"),
      worksheetScope = worksheetScopeName
    )
    writeNamedRegion(wb, df, name = namedRegion, worksheetScope = worksheetScopeName, header = TRUE)
    res <- readNamedRegion(wb, namedRegion, worksheetScope = worksheetScopeName)
    expect_equal(res, normalizeDataframe(df, replaceInf = TRUE), ignore_attr = c("worksheetScope", "row.names"))
  }
  testDataFrameGlobalExplicit <- function(wb, df, lref) {
    namedRegion <- paste(deparse(substitute(df)), "forGlobal", sep = "")
    createSheet(wb, name = namedRegion)
    createName(wb, name = namedRegion, formula = paste(namedRegion, lref, sep = "!"), worksheetScope = "")
    writeNamedRegion(wb, df, name = namedRegion, worksheetScope = "", header = TRUE)
    res <- readNamedRegion(wb, namedRegion, worksheetScope = "")
    expect_equal(res, normalizeDataframe(df, replaceInf = TRUE), ignore_attr = c("worksheetScope"))
  }
  testDataFrameGlobalAndScoped <- function(wb, df_global, df_scoped, lref) {
    dfs <- list(df_global, df_scoped)
    namedRegion <- paste0(deparse(substitute(df_global)), "expect_global")
    sheet_2 <- paste0(namedRegion, "2")
    sheetNames <- c(namedRegion, sheet_2)
    scopeSheets <- c("", sheet_2)
    createSheet(wb, name = sheetNames)
    createName(wb, name = namedRegion, formula = paste(sheetNames, lref, sep = "!"), worksheetScope = scopeSheets)
    writeNamedRegion(wb, dfs, name = namedRegion, worksheetScope = scopeSheets, header = TRUE)
    res_full <- readNamedRegion(wb, namedRegion, worksheetScope = scopeSheets)
    dfs_norm <- list(normalizeDataframe(df_global, replaceInf = TRUE), normalizeDataframe(df_scoped, replaceInf = TRUE))
    expect_equal(res_full, dfs_norm, ignore_attr = c("worksheetScope", "row.names", "names"))
    res_global_prio <- readNamedRegion(wb, namedRegion)
    expect_equal(
      res_global_prio,
      normalizeDataframe(df_global, replaceInf = TRUE),
      ignore_attr = c("worksheetScope", "row.names")
    )
  }
  # built-in dataset mtcars (*.xls)
  testDataFrame(wb.xls, mtcars, "$C$8")
  # built-in dataset mtcars (*.xlsx)
  testDataFrame(wb.xlsx, mtcars, "$C$8")

  # built-in dataset airquality (*.xls)
  testDataFrame(wb.xls, airquality, "$F$13")
  # built-in dataset airquality (*.xlsx)
  testDataFrame(wb.xlsx, airquality, "$F$13")

  # built-in dataset attenu (*.xls)
  testDataFrame(wb.xls, attenu, "$A$8")
  # built-in dataset attenu (*.xlsx)
  testDataFrame(wb.xlsx, attenu, "$A$8")

  # built-in dataset attenu (*.xls) - explicit global scope
  testDataFrameGlobalExplicit(wb.xls, attenu, "$A$8")
  # built-in dataset attenu (*.xlsx) - explicit global scope
  testDataFrameGlobalExplicit(wb.xlsx, attenu, "$A$8")

  # built-in dataset ChickWeight (*.xls)
  testDataFrame(wb.xls, ChickWeight, "$BQ$7")
  # built-in dataset ChickWeight (*.xlsx)
  testDataFrame(wb.xlsx, ChickWeight, "$BQ$7")

  # built-in dataset ChickWeight (*.xls) - explicit global scope
  testDataFrameGlobalExplicit(wb.xls, ChickWeight, "$BQ$7")
  # built-in dataset ChickWeight (*.xlsx) - explicit global scope
  testDataFrameGlobalExplicit(wb.xlsx, ChickWeight, "$BQ$7")

  # built-in dataset CO2 (*.xls)
  CO <- CO2 # CO2 seems to be an illegal name
  testDataFrame(wb.xls, CO, "$L$1")
  # built-in dataset CO2 (*.xlsx)
  testDataFrame(wb.xlsx, CO, "$L$1")

  # built-in dataset iris (*.xls)
  testDataFrame(wb.xls, iris, "$BB$5")
  # built-in dataset iris (*.xlsx)
  testDataFrame(wb.xlsx, iris, "$BB$5")

  # built-in dataset longley (*.xls)
  testDataFrame(wb.xls, longley, "$AD$8")
  # built-in dataset longley (*.xlsx)
  testDataFrame(wb.xlsx, longley, "$AD$8")

  # built-in dataset morley (*.xls)
  testDataFrame(wb.xls, morley, "$K$4")
  # built-in dataset morley (*.xlsx)
  testDataFrame(wb.xlsx, morley, "$K$4")

  # built-in dataset morley (*.xls)
  testDataFrameNameScope(wb.xls, morley, "$K$4")
  # built-in dataset morley (*.xlsx)
  testDataFrameNameScope(wb.xlsx, morley, "$K$4")

  # built-in dataset swiss (*.xls)
  testDataFrame(wb.xls, swiss, "$M$2")
  # built-in dataset swiss (*.xlsx)
  testDataFrame(wb.xlsx, swiss, "$M$2")

  # built-in dataset swiss (*.xls)
  testDataFrameNameScope(wb.xls, swiss, "$M$2")
  # built-in dataset swiss (*.xlsx)
  testDataFrameNameScope(wb.xlsx, swiss, "$M$2")

  # built-in datasets swiss and morley (*.xls)
  testDataFrameGlobalAndScoped(wb.xls, swiss, morley, "$M$2")
  # built-in dataset swiss and morley (*.xlsx)
  testDataFrameGlobalAndScoped(wb.xlsx, swiss, morley, "$M$2")
})
