test_that("test.writeWorksheetToFile - always run", {
    # Add tests here that should always run, if any.
    # For now, this block will be empty as all existing tests are conditional.
    expect_true(TRUE) # Placeholder to ensure the test block is not empty
})

test_that("test.writeWorksheetToFile - full test suite only", {
    skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")

    file.xls <- "testWriteWorksheetToFileWorkbook.xls"
    file.xlsx <- "testWriteWorksheetToFileWorkbook.xlsx"
    if (file.exists(file.xls)) {
        file.remove(file.xls)
    }
    if (file.exists(file.xlsx)) {
        file.remove(file.xlsx)
    }
    testDataFrame <- function(file, df) {
        worksheet <- deparse(substitute(df))
        print(paste("Writing dataset ", worksheet, "to file",
            file))
        name <- paste(worksheet, "Region", sep = "")
        writeWorksheetToFile(file, data = df, sheet = name)
        res <- readWorksheetFromFile(file, sheet = name)
        expect_equal(normalizeDataframe(df), res, check.attributes = FALSE,
            check.names = TRUE)
    }
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
    cdf <- data.frame(Column.A = c(1, 2, 3, NA, 5, 6, 7,
        8, NA, 10), Column.B = c(-4, -3, NA, -1, 0, NA, NA,
        3, 4, 5), Column.C = c("Anna", "???", NA, "", NA,
        "$!?&%", "(?2@?~?'^*#|)", "{}[]:,;-_<>", "\\sadf\n\nv",
        "a b c"), Column.D = c(pi, -pi, NA, sqrt(2), sqrt(0.3),
        -sqrt(pi), exp(1), log(2), sin(2), -tan(2)), Column.E = c(TRUE,
        TRUE, NA, NA, FALSE, FALSE, TRUE, NA, FALSE, TRUE),
        Column.F = c("High", "Medium", "Low", "Low", "Low",
            NA, NA, "Medium", "High", "High"), Column.G = c("High",
            "Medium", NA, "Low", "Low", "Medium", NA, "Medium",
            "High", "High"), Column.H = rep(c(as.Date("2021-10-30"),
            as.Date("2021-03-28"), NA), length = 10), Column.I = rep(c(as.POSIXlt("2021-10-31 03:00:00"),
            as.POSIXlt(1582963631, origin = "1970-01-01"),
            NA, as.POSIXlt("2001-12-31 23:59:59")), length = 10),
        Column.J = rep(c(as.POSIXct("2021-10-31 03:00:00"),
            as.POSIXct(1582963631, origin = "1970-01-01"),
            NA, as.POSIXct("2001-12-31 23:59:59")), length = 10),
        stringsAsFactors = FALSE)
    cdf[["Column.F"]] <- factor(cdf[["Column.F"]])
    cdf[["Column.F"]] <- ordered(cdf[["Column.F"]], levels = c("Low",
        "Medium", "High"))
    testDataFrame(file.xls, cdf)
    testDataFrame(file.xlsx, cdf)
    testClearSheets <- function(file, df) {
        df.short <- df[1, ]
        writeWorksheetToFile(file, data = c(df.short), sheet = "cdfRegion")
        expect_equal(nrow(readWorksheetFromFile(file, sheet = "cdfRegion")),
            nrow(df))
        writeWorksheetToFile(file, data = c(df.short), sheet = "cdfRegion",
            clearSheets = TRUE)
        expect_equal(nrow(readWorksheetFromFile(file, sheet = "cdfRegion")),
            1)
    }
    testClearSheets(file.xls, cdf)
    testClearSheets(file.xlsx, cdf)
})
