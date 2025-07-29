# TODO use subtests / parameterized tests when mature - see https://github.com/r-lib/testthat/issues/2063 and https://github.com/google/patrick/issues/20

testDataFrame <- function(file, df) {
    worksheet <- deparse(substitute(df))
    print(paste("Writing dataset ", worksheet, "to file", file))
    name <- paste(worksheet, "Region", sep = "")
    writeWorksheetToFile(file, data = df, sheet = name)
    res <- readWorksheetFromFile(file, sheet = name)
    expect_equal(
        normalizeDataframe(df),
        res,
        ignore_attr = TRUE,
    )
}


test_that("test.writeWorksheetToFile - full test suite only", {
    skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
    file.xls <- withr::local_tempfile(fileext = ".xls")
    file.xlsx <- withr::local_tempfile(fileext = ".xlsx")

    testDataFrame(file.xls, mtcars)
    testDataFrame(file.xlsx, mtcars)
    testDataFrame(file.xls, airquality)
    testDataFrame(file.xlsx, airquality)
    testDataFrame(file.xls, attenu)
    testDataFrame(file.xlsx, attenu)
    testDataFrame(file.xls, ChickWeight)
    testDataFrame(file.xlsx, ChickWeight)
    CO = CO2
    testDataFrame(file.xls, CO2)
    testDataFrame(file.xlsx, CO2)
    testDataFrame(file.xls, iris)
    testDataFrame(file.xlsx, iris)
    testDataFrame(file.xls, longley)
    testDataFrame(file.xlsx, longley)
    testDataFrame(file.xls, morley)
    testDataFrame(file.xlsx, morley)
    testDataFrame(file.xls, swiss)
    testDataFrame(file.xlsx, swiss)
    cdf <- data.frame(
        Column.A = c(1, 2, 3, NA, 5, 6, 7, 8, NA, 10),
        Column.B = c(-4, -3, NA, -1, 0, NA, NA, 3, 4, 5),
        Column.C = c(
            "Anna",
            "???",
            NA,
            "",
            NA,
            "$!?&%",
            "(?2@?~?'^*#|)",
            "{}[]:,;-_<>",
            "\\sadf\n\nv",
            "a b c"
        ),
        Column.D = c(
            pi,
            -pi,
            NA,
            sqrt(2),
            sqrt(0.3),
            -sqrt(pi),
            exp(1),
            log(2),
            sin(2),
            -tan(2)
        ),
        Column.E = c(TRUE, TRUE, NA, NA, FALSE, FALSE, TRUE, NA, FALSE, TRUE),
        Column.F = c(
            "High",
            "Medium",
            "Low",
            "Low",
            "Low",
            NA,
            NA,
            "Medium",
            "High",
            "High"
        ),
        Column.G = c(
            "High",
            "Medium",
            NA,
            "Low",
            "Low",
            "Medium",
            NA,
            "Medium",
            "High",
            "High"
        ),
        Column.H = rep(
            c(as.Date("2021-10-30"), as.Date("2021-03-28"), NA),
            length = 10
        ),
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
        stringsAsFactors = FALSE
    )
    cdf[["Column.F"]] <- factor(cdf[["Column.F"]])
    cdf[["Column.F"]] <- ordered(
        cdf[["Column.F"]],
        levels = c("Low", "Medium", "High")
    )
    testDataFrame(file.xls, cdf)
    testDataFrame(file.xlsx, cdf)
})


test_that("clearSheets parameter works correctly for XLS", {
    skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
    file.xls <- withr::local_tempfile(fileext = ".xls")

    cdf <- data.frame(a = 1:10)
    df.short <- cdf[1, , drop = FALSE]

    writeWorksheetToFile(file.xls, data = df.short, sheet = "cdfRegion")

    writeWorksheetToFile(file.xls, data = cdf, sheet = "cdfRegion")
    expect_equal(
        nrow(readWorksheetFromFile(file.xls, sheet = "cdfRegion")),
        nrow(cdf)
    )

    writeWorksheetToFile(
        file.xls,
        data = df.short,
        sheet = "cdfRegion",
        clearSheets = TRUE
    )
    expect_equal(
        nrow(readWorksheetFromFile(file.xls, sheet = "cdfRegion")),
        nrow(df.short)
    )
})


test_that("clearSheets parameter works correctly for XLSX", {
    skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
    file.xlsx <- withr::local_tempfile(fileext = ".xlsx")

    cdf <- data.frame(a = 1:10)
    df.short <- cdf[1, , drop = FALSE]

    writeWorksheetToFile(file.xlsx, data = df.short, sheet = "cdfRegion")

    writeWorksheetToFile(file.xlsx, data = cdf, sheet = "cdfRegion")
    expect_equal(
        nrow(readWorksheetFromFile(file.xlsx, sheet = "cdfRegion")),
        nrow(cdf)
    )

    writeWorksheetToFile(
        file.xlsx,
        data = df.short,
        sheet = "cdfRegion",
        clearSheets = TRUE
    )
    expect_equal(
        nrow(readWorksheetFromFile(file.xlsx, sheet = "cdfRegion")),
        nrow(df.short)
    )
})
