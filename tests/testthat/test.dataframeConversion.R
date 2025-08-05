test_that("test.dataframeConversion - always run", {
  testDataFrame <- function(df) {
    res <- XLConnect:::dataframeFromJava(XLConnect:::dataframeToJava(df), check.names = TRUE)
    expect_equal(normalizeDataframe(df), res, ignore_attr = c("worksheetScope"))
  }
  cdf <- data.frame(
    Column.A = c(1, 2, 3, NA, 5, Inf, 7, 8, NA, 10),
    Column.B = c(-4, -3, NA, -Inf, 0, NA, NA, 3, 4, 5),
    Column.C = c("Anna", "???", NA, "", NA, "$!?&%", "(?2@?~?'^*#|)", "{}[]:,;-_<>", "\\sadf\n\nv", "a b c"),
    Column.D = c(pi, -pi, NA, sqrt(2), sqrt(0.3), -sqrt(pi), exp(1), log(2), sin(2), -tan(2)),
    Column.E = c(TRUE, TRUE, NA, NA, FALSE, FALSE, TRUE, NA, FALSE, TRUE),
    Column.F = c("High", "Medium", "Low", "Low", "Low", NA, NA, "Medium", "High", "High"),
    Column.G = c("High", "Medium", NA, "Low", "Low", "Medium", NA, "Medium", "High", "High"),
    Column.H = rep(c(as.Date("2021-10-30"), as.Date("2021-03-28"), NA), length = 10),
    Column.I = rep(
      c(
        as.POSIXlt("2021-10-31 03:00:00"),
        as.POSIXlt(1582963631, origin = "1970-01-01"),
        NA,
        as.POSIXlt("2001-12-31 23:59:59")
      ),
      length = 10
    ),
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
  testDataFrame(cdf)
  expect_error(XLConnect:::dataframeToJava(search))
  expect_error(XLConnect:::dataframeFromJava(NULL))
  expect_error(XLConnect:::dataframeFromJava(NA))
  expect_error(XLConnect:::dataframeFromJava(9))
  expect_error(XLConnect:::dataframeFromJava(search))
})

test_that("test.dataframeConversion - full test suite only", {
  skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")

  testDataFrame <- function(df) {
    res <- XLConnect:::dataframeFromJava(XLConnect:::dataframeToJava(df), check.names = TRUE)
    expect_equal(normalizeDataframe(df), res, ignore_attr = c("worksheetScope", "row.names"))
  }
  testDataFrame(mtcars)
  testDataFrame(airquality)
  testDataFrame(attenu)
  testDataFrame(ChickWeight)
  testDataFrame(CO2)
  testDataFrame(iris)
  testDataFrame(longley)
  testDataFrame(morley)
  testDataFrame(swiss)
})
