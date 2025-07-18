test_that("writeWorksheetToFile writes and reads a data frame to/from an XLS file", {
    skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
    file.xls <- "testWriteWorksheetToFileWorkbook.xls"
    on.exit(if (file.exists(file.xls)) file.remove(file.xls))

    testDataFrame <- function(df) {
        worksheet <- deparse(substitute(df))
        name <- paste(worksheet, "Region", sep = "")
        writeWorksheetToFile(file.xls, data = df, sheet = name)
        res <- readWorksheetFromFile(file.xls, sheet = name)
        expect_equal(normalizeDataframe(df), res, ignore_attr = TRUE)
    }

    testDataFrame(mtcars)
    testDataFrame(airquality)
})

test_that("writeWorksheetToFile writes and reads a data frame to/from an XLSX file", {
    skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
    file.xlsx <- "testWriteWorksheetToFileWorkbook.xlsx"
    on.exit(if (file.exists(file.xlsx)) file.remove(file.xlsx))

    testDataFrame <- function(df) {
        worksheet <- deparse(substitute(df))
        name <- paste(worksheet, "Region", sep = "")
        writeWorksheetToFile(file.xlsx, data = df, sheet = name)
        res <- readWorksheetFromFile(file.xlsx, sheet = name)
        expect_equal(normalizeDataframe(df), res, ignore_attr = TRUE)
    }

    testDataFrame(attenu)
    testDataFrame(ChickWeight)
})

test_that("clearSheets parameter works correctly for XLS", {
    skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
    file.xls <- "testWriteWorksheetToFileWorkbook.xls"
    on.exit(if (file.exists(file.xls)) file.remove(file.xls))

    cdf <- data.frame(a = 1:10)
    df.short <- cdf[1, , drop = FALSE]

    writeWorksheetToFile(file.xls, data = df.short, sheet = "cdfRegion")
    # expect_equal(nrow(readWorksheetFromFile(file.xls, sheet = "cdfRegion")), nrow(df.short))

    writeWorksheetToFile(file.xls, data = cdf, sheet = "cdfRegion")
    expect_equal(nrow(readWorksheetFromFile(file.xls, sheet = "cdfRegion")), nrow(cdf))

    writeWorksheetToFile(file.xls, data = df.short, sheet = "cdfRegion", clearSheets = TRUE)
    expect_equal(nrow(readWorksheetFromFile(file.xls, sheet = "cdfRegion")), nrow(df.short))
})

test_that("clearSheets parameter works correctly for XLSX", {
    skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
    file.xlsx <- "testWriteWorksheetToFileWorkbook.xlsx"
    on.exit(if (file.exists(file.xlsx)) file.remove(file.xlsx))

    cdf <- data.frame(a = 1:10)
    df.short <- cdf[1, , drop = FALSE]

    writeWorksheetToFile(file.xlsx, data = df.short, sheet = "cdfRegion")
    # expect_equal(nrow(readWorksheetFromFile(file.xlsx, sheet = "cdfRegion")), nrow(df.short))

    writeWorksheetToFile(file.xlsx, data = cdf, sheet = "cdfRegion")
    expect_equal(nrow(readWorksheetFromFile(file.xlsx, sheet = "cdfRegion")), nrow(cdf))

    writeWorksheetToFile(file.xlsx, data = df.short, sheet = "cdfRegion", clearSheets = TRUE)
    expect_equal(nrow(readWorksheetFromFile(file.xlsx, sheet = "cdfRegion")), nrow(df.short))
})
