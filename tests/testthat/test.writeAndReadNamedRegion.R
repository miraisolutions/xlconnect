test_that("test.writeAndReadNamedRegion", {
    wb.xls <- loadWorkbook("resources/testWriteAndReadNamedRegion.xls",
        create = TRUE
    )
    wb.xlsx <- loadWorkbook("resources/testWriteAndReadNamedRegion.xlsx",
        create = TRUE
    )
    testDataFrame <- function(wb, df, lref) {
        namedRegion <- deparse(substitute(df))
        createSheet(wb, name = namedRegion)
        createName(wb, name = namedRegion, formula = paste(namedRegion,
            lref,
            sep = "!"
        ), worksheetScope = namedRegion)
        writeNamedRegion(wb, df,
            name = namedRegion, worksheetScope = namedRegion,
            header = TRUE
        )
        res <- readNamedRegion(wb, namedRegion, worksheetScope = namedRegion)
        expect_equal(normalizeDataframe(df, replaceInf = TRUE),
            res,
            check.attributes = FALSE, check.names = TRUE
        )
    }
    testDataFrameNameScope <- function(wb, df, lref) {
        namedRegion <- paste(deparse(substitute(df)), "1", sep = "")
        worksheetScopeName <- paste(namedRegion, "2", sep = "")
        createSheet(wb, name = namedRegion)
        createSheet(wb, name = worksheetScopeName)
        createName(wb, name = namedRegion, formula = paste(namedRegion,
            lref,
            sep = "!"
        ), worksheetScope = worksheetScopeName)
        writeNamedRegion(wb, df,
            name = namedRegion, worksheetScope = worksheetScopeName,
            header = TRUE
        )
        res <- readNamedRegion(wb, namedRegion, worksheetScope = worksheetScopeName)
        expect_equal(normalizeDataframe(df, replaceInf = TRUE),
            res,
            check.attributes = FALSE, check.names = TRUE
        )
    }
    testDataFrameGlobalExplicit <- function(wb, df, lref) {
        namedRegion <- paste(deparse(substitute(df)), "forGlobal",
            sep = ""
        )
        createSheet(wb, name = namedRegion)
        createName(wb, name = namedRegion, formula = paste(namedRegion,
            lref,
            sep = "!"
        ), worksheetScope = "")
        writeNamedRegion(wb, df,
            name = namedRegion, worksheetScope = "",
            header = TRUE
        )
        res <- readNamedRegion(wb, namedRegion, worksheetScope = "")
        expect_equal(normalizeDataframe(df, replaceInf = TRUE),
            res,
            check.attributes = FALSE, check.names = TRUE
        )
    }
    testDataFrameGlobalAndScoped <- function(wb, df_global, df_scoped,
                                             lref) {
        dfs <- list(df_global, df_scoped)
        namedRegion <- paste0(
            deparse(substitute(df_global)),
            "expect_global"
        )
        sheet_2 <- paste0(namedRegion, "2")
        sheetNames <- c(namedRegion, sheet_2)
        scopeSheets <- c("", sheet_2)
        createSheet(wb, name = sheetNames)
        createName(wb, name = namedRegion, formula = paste(sheetNames,
            lref,
            sep = "!"
        ), worksheetScope = scopeSheets)
        writeNamedRegion(wb, dfs,
            name = namedRegion, worksheetScope = scopeSheets,
            header = TRUE
        )
        res_full <- readNamedRegion(wb, namedRegion, worksheetScope = scopeSheets)
        dfs_norm <- list(
            normalizeDataframe(df_global, replaceInf = TRUE),
            normalizeDataframe(df_scoped, replaceInf = TRUE)
        )
        expect_equal(dfs_norm, res_full,
            check.attributes = FALSE,
            check.names = TRUE
        )
        res_global_prio <- readNamedRegion(wb, namedRegion)
        expect_equal(normalizeDataframe(df_global, replaceInf = TRUE),
            res_global_prio,
            check.attributes = FALSE, check.names = TRUE
        )
    }
    if (getOption("FULL.TEST.SUITE")) {
        testDataFrame(wb.xls, mtcars, "$C$8")
        testDataFrame(wb.xlsx, mtcars, "$C$8")
        testDataFrame(wb.xls, airquality, "$F$13")
        testDataFrame(wb.xlsx, airquality, "$F$13")
        testDataFrame(wb.xls, attenu, "$A$8")
        testDataFrame(wb.xlsx, attenu, "$A$8")
        testDataFrameGlobalExplicit(wb.xls, attenu, "$A$8")
        testDataFrameGlobalExplicit(wb.xlsx, attenu, "$A$8")
        testDataFrame(wb.xls, ChickWeight, "$BQ$7")
        testDataFrame(wb.xlsx, ChickWeight, "$BQ$7")
        testDataFrameGlobalExplicit(wb.xls, ChickWeight, "$BQ$7")
        testDataFrameGlobalExplicit(wb.xlsx, ChickWeight, "$BQ$7")
        CO <- CO2
        testDataFrame(wb.xls, CO, "$L$1")
        testDataFrame(wb.xlsx, CO, "$L$1")
        testDataFrame(wb.xls, iris, "$BB$5")
        testDataFrame(wb.xlsx, iris, "$BB$5")
        testDataFrame(wb.xls, longley, "$AD$8")
        testDataFrame(wb.xlsx, longley, "$AD$8")
        testDataFrame(wb.xls, morley, "$K$4")
        testDataFrame(wb.xlsx, morley, "$K$4")
        testDataFrameNameScope(wb.xls, morley, "$K$4")
        testDataFrameNameScope(wb.xlsx, morley, "$K$4")
        testDataFrame(wb.xls, swiss, "$M$2")
        testDataFrame(wb.xlsx, swiss, "$M$2")
        testDataFrameNameScope(wb.xls, swiss, "$M$2")
        testDataFrameNameScope(wb.xlsx, swiss, "$M$2")
        testDataFrameGlobalAndScoped(wb.xls, swiss, morley, "$M$2")
        testDataFrameGlobalAndScoped(
            wb.xlsx, swiss, morley,
            "$M$2"
        )
    }
    cdf <- data.frame(
        Column.A = c(
            1, 2, 3, NA, 5, Inf, 7, 8,
            NA, 10
        ), Column.B = c(
            -4, -3, NA, -Inf, 0, NA, NA, 3,
            4, 5
        ), Column.C = c(
            "Anna", "???", NA, "", NA, "$!?&%",
            "(?2@?~?'^*#|)", "{}[]:,;-_<>", "\\sadf\n\nv", "a b c"
        ),
        Column.D = c(
            pi, -pi, NA, sqrt(2), sqrt(0.3), -sqrt(pi),
            exp(1), log(2), sin(2), -tan(2)
        ), Column.E = c(
            TRUE,
            TRUE, NA, NA, FALSE, FALSE, TRUE, NA, FALSE, TRUE
        ),
        Column.F = c(
            "High", "Medium", "Low", "Low", "Low", NA,
            NA, "Medium", "High", "High"
        ), Column.G = c(
            "High",
            "Medium", NA, "Low", "Low", "Medium", NA, "Medium",
            "High", "High"
        ), Column.H = rep(c(
            as.Date("2021-10-30"),
            as.Date("2021-03-28"), NA
        ), length = 10), Column.I = rep(c(
            as.POSIXlt("2021-10-31 03:00:00"),
            as.POSIXlt(1582963631, origin = "1970-01-01"), NA,
            as.POSIXlt("2001-12-31 23:59:59")
        ), length = 10),
        Column.J = rep(
            c(as.POSIXct("2021-10-31 03:00:00"), as.POSIXct(1582963631,
                origin = "1970-01-01"
            ), NA, as.POSIXct("2001-12-31 23:59:59")),
            length = 10
        ), stringsAsFactors = F
    )
    cdf[["Column.F"]] <- factor(cdf[["Column.F"]])
    cdf[["Column.F"]] <- ordered(cdf[["Column.F"]], levels = c(
        "Low",
        "Medium", "High"
    ))
    testDataFrame(wb.xls, cdf, "$X$100")
    testDataFrame(wb.xlsx, cdf, "$X$100")
    createSheet(wb.xls, name = "rownames")
    createName(wb.xls, name = "rownames", formula = "rownames!$F$16")
    writeNamedRegion(wb.xls, mtcars,
        name = "rownames", header = TRUE,
        rownames = "Car"
    )
    res <- readNamedRegion(wb.xls, "rownames")
    expect_equal(res, XLConnect:::includeRownames(mtcars, "Car"), check.attributes = FALSE)
    createSheet(wb.xlsx, name = "rownames")
    createName(wb.xlsx, name = "rownames", formula = "rownames!$F$16")
    writeNamedRegion(wb.xlsx, mtcars,
        name = "rownames", header = TRUE,
        rownames = "Car"
    )
    res <- readNamedRegion(wb.xlsx, "rownames")
    expect_equal(res, XLConnect:::includeRownames(mtcars, "Car"), check.attributes = FALSE)
    createSheet(wb.xls, name = "rownames2")
    createName(wb.xls, name = "rownames2", formula = "rownames2!$K$5")
    writeNamedRegion(wb.xls, mtcars,
        name = "rownames2", header = TRUE,
        rownames = "Car"
    )
    res <- readNamedRegion(wb.xls, "rownames2", rownames = "Car")
    expect_equal(res, mtcars)
    expect_equal(attr(res, "row.names"), attr(mtcars, "row.names"))
    createSheet(wb.xlsx, name = "rownames2")
    createName(wb.xlsx, name = "rownames2", formula = "rownames2!$K$5")
    writeNamedRegion(wb.xlsx, mtcars,
        name = "rownames2", header = TRUE,
        rownames = "Car"
    )
    res <- readNamedRegion(wb.xlsx, "rownames2", rownames = "Car")
    expect_equal(res, mtcars)
    expect_equal(attr(res, "row.names"), attr(mtcars, "row.names"))
})
