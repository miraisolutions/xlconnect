test_that("setMissingValue works correctly for a single value in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookSetMissingValue.xls", create = TRUE)
    data <- data.frame(A = c(4.2, -3.2, NA, 1.34), B = c("A", NA, "C", "D"), stringsAsFactors = FALSE)
    name <- "missing"
    createSheet(wb.xls, name = name)
    createName(wb.xls, name = name, formula = paste(name, "$A$1", sep = "!"))

    writeNamedRegion(wb.xls, data, name = name)
    res <- readNamedRegion(wb.xls, name = name)
    attr(data, "worksheetScope") <- ""
    expect_equal(res, data)

    expect <- data.frame(A = c("4.2", "-3.2", "missing", "1.34"), B = c("A", "missing", "C", "D"), stringsAsFactors = FALSE)
    attr(expect, "worksheetScope") <- ""
    setMissingValue(wb.xls, value = "missing")
    writeNamedRegion(wb.xls, data, name = name)
    setMissingValue(wb.xls, value = NULL)
    res <- readNamedRegion(wb.xls, name = name)
    expect_equal(res, expect)
})

test_that("setMissingValue works correctly for a single value in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookSetMissingValue.xlsx", create = TRUE)
    data <- data.frame(A = c(4.2, -3.2, NA, 1.34), B = c("A", NA, "C", "D"), stringsAsFactors = FALSE)
    name <- "missing"
    createSheet(wb.xlsx, name = name)
    createName(wb.xlsx, name = name, formula = paste(name, "$A$1", sep = "!"))

    writeNamedRegion(wb.xlsx, data, name = name)
    res <- readNamedRegion(wb.xlsx, name = name)
    expect_equal(res, data)

    expect <- data.frame(A = c("4.2", "-3.2", "missing", "1.34"), B = c("A", "missing", "C", "D"), stringsAsFactors = FALSE)
    setMissingValue(wb.xlsx, value = "missing")
    writeNamedRegion(wb.xlsx, data, name = name)
    setMissingValue(wb.xlsx, value = NULL)
    res <- readNamedRegion(wb.xlsx, name = name)
    expect_equal(res, expect)
})

test_that("setMissingValue works with a vector of values in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookMissingValue.xls", create = FALSE)
    expect <- data.frame(A = c(NA, -3.2, 3.4, NA, 8, NA), B = c("a", NA, "c", "x", "a", "o"),
                         C = c(TRUE, TRUE, FALSE, NA, FALSE, NA),
                         D = as.POSIXct(c("1981-12-01 00:00:00", "1981-12-02 00:00:00", NA, NA, NA, "1981-12-06 00:00:00")),
                         stringsAsFactors = FALSE)
    setMissingValue(wb.xls, value = c("NA", "missing", "empty"))
    res <- readNamedRegion(wb.xls, name = "Missing1")
    expect_equal(expect, res)
})

test_that("setMissingValue works with a vector of values in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookMissingValue.xlsx", create = FALSE)
    expect <- data.frame(A = c(NA, -3.2, 3.4, NA, 8, NA), B = c("a", NA, "c", "x", "a", "o"),
                         C = c(TRUE, TRUE, FALSE, NA, FALSE, NA),
                         D = as.POSIXct(c("1981-12-01 00:00:00", "1981-12-02 00:00:00", NA, NA, NA, "1981-12-06 00:00:00")),
                         stringsAsFactors = FALSE)
    setMissingValue(wb.xlsx, value = c("NA", "missing", "empty"))
    res <- readNamedRegion(wb.xlsx, name = "Missing1")
    expect_equal(expect, res)
})

test_that("setMissingValue works with a list of values in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookMissingValue.xls", create = FALSE)
    expect <- data.frame(A = c(NA, -3.2, NA, NA, 8, NA), B = c("a", NA, "c", "x", "a", "o"),
                         C = c(TRUE, NA, FALSE, NA, FALSE, NA),
                         D = as.POSIXct(c("1981-12-01 00:00:00", "1981-12-02 00:00:00", NA, NA, NA, "1981-12-06 00:00:00")),
                         stringsAsFactors = FALSE)
    setMissingValue(wb.xls, value = list("NA", "missing", "empty", -9999))
    res <- readNamedRegion(wb.xls, name = "Missing2")
    expect_equal(expect, res)
})

test_that("setMissingValue works with a list of values in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookMissingValue.xlsx", create = FALSE)
    expect <- data.frame(A = c(NA, -3.2, NA, NA, 8, NA), B = c("a", NA, "c", "x", "a", "o"),
                         C = c(TRUE, NA, FALSE, NA, FALSE, NA),
                         D = as.POSIXct(c("1981-12-01 00:00:00", "1981-12-02 00:00:00", NA, NA, NA, "1981-12-06 00:00:00")),
                         stringsAsFactors = FALSE)
    setMissingValue(wb.xlsx, value = list("NA", "missing", "empty", -9999))
    res <- readNamedRegion(wb.xlsx, name = "Missing2")
    expect_equal(expect, res)
})
