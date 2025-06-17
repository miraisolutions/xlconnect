test_that("test.writeNamedRegionToFile", {
    if (getOption("FULL.TEST.SUITE")) {
        file.xls <- "testWriteNamedRegionToFileWorkbook.xls"
        file.xlsx <- "testWriteNamedRegionToFileWorkbook.xlsx"
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
            writeNamedRegionToFile(file, df, name, formula = paste(worksheet, 
                "A1", sep = "!"))
            res <- readNamedRegionFromFile(file, name)
            checkEquals(normalizeDataframe(df), res, check.attributes = FALSE, 
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
        testDataFrame(file.xls, CO)
        testDataFrame(file.xlsx, CO)
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
            stringsAsFactors = F)
        cdf[["Column.F"]] <- factor(cdf[["Column.F"]])
        cdf[["Column.F"]] <- ordered(cdf[["Column.F"]], levels = c("Low", 
            "Medium", "High"))
        testDataFrame(file.xls, cdf)
        testDataFrame(file.xlsx, cdf)
        file2.xls <- "wnrtf1.xls"
        file2.xlsx <- "wnrtf1.xlsx"
        if (file.exists(file2.xls)) 
            file.remove(file2.xls)
        if (file.exists(file2.xlsx)) 
            file.remove(file2.xlsx)
        checkNoException(writeNamedRegionToFile(file2.xls, data = mtcars, 
            name = "mtcars", formula = "'My Cars'!$A$1", header = TRUE))
        expect_true(file.exists(file2.xls))
        checkNoException(writeNamedRegionToFile(file2.xlsx, data = mtcars, 
            name = "mtcars", formula = "'My Cars'!$A$1", header = TRUE))
        expect_true(file.exists(file2.xlsx))
        testClearNamedRegions <- function(file, df) {
            df.short <- df[1, ]
            writeNamedRegionToFile(file, data = df.short, name = "cdfRegion")
            checkEquals(nrow(readNamedRegionFromFile(file, name = "cdfRegion")), 
                1)
            checkEquals(nrow(readWorksheetFromFile(file, sheet = "cdf")), 
                nrow(df))
            writeNamedRegionToFile(file, data = df, name = "cdfRegion")
            writeNamedRegionToFile(file, data = df.short, name = "cdfRegion", 
                clearNamedRegions = TRUE)
            checkEquals(nrow(readNamedRegionFromFile(file, name = "cdfRegion")), 
                1)
            checkEquals(nrow(readWorksheetFromFile(file, sheet = "cdf")), 
                1)
        }
        testClearNamedRegions(file.xls, cdf)
        testClearNamedRegions(file.xlsx, cdf)
        testClearNamedRegionsScoped <- function(file, df) {
            scope <- c("scope1", "scope2")
            clearParam <- c(TRUE, FALSE)
            df.short <- df[1, ]
            wb <- loadWorkbook(file, create = TRUE)
            createSheet(wb, scope)
            saveWorkbook(wb, file)
            writeNamedRegionToFile(file, data = df, name = "cdfRegionScoped", 
                formula = paste(scope, "A1", sep = "!"), worksheetScope = scope)
            writeNamedRegionToFile(file, data = df.short, name = "cdfRegionScoped", 
                worksheetScope = scope)
            checkEquals(nrow(readNamedRegionFromFile(file, name = "cdfRegionScoped", 
                worksheetScope = scope)[[1]]), 1)
            checkEquals(nrow(readNamedRegionFromFile(file, name = "cdfRegionScoped", 
                worksheetScope = scope)[[2]]), 1)
            checkEquals(nrow(readWorksheetFromFile(file, sheet = scope)[[1]]), 
                nrow(df))
            checkEquals(nrow(readWorksheetFromFile(file, sheet = scope)[[2]]), 
                nrow(df))
            writeNamedRegionToFile(file, data = df, name = "cdfRegionScoped", 
                worksheetScope = scope)
            writeNamedRegionToFile(file, data = df.short, name = "cdfRegionScoped", 
                clearNamedRegions = clearParam, worksheetScope = scope)
            checkEquals(nrow(readNamedRegionFromFile(file, name = "cdfRegionScoped", 
                worksheetScope = scope)[[1]]), 1)
            checkEquals(nrow(readNamedRegionFromFile(file, name = "cdfRegionScoped", 
                worksheetScope = scope)[[2]]), 1)
            checkEquals(nrow(readWorksheetFromFile(file, sheet = scope)[[1]]), 
                1)
            checkEquals(nrow(readWorksheetFromFile(file, sheet = scope)[[2]]), 
                nrow(df))
        }
        scopedfile.xls <- "testWriteNamedRegionToFileWorkbookScoped.xls"
        scopedfile.xlsx <- "testWriteNamedRegionToFileWorkbookScoped.xlsx"
        testClearNamedRegionsScoped(scopedfile.xls, cdf)
        testClearNamedRegionsScoped(scopedfile.xlsx, cdf)
    }
})

