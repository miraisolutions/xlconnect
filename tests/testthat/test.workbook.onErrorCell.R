test_that("test.workbook.onErrorCell", {
    wb.xls <- loadWorkbook("resources/testWorkbookErrorCell.xls",
        create = FALSE
    )
    wb.xlsx <- loadWorkbook("resources/testWorkbookErrorCell.xlsx",
        create = FALSE
    )
    onErrorCell(wb.xls, XLC$ERROR.WARN)
    expect_warning(res <- try(readNamedRegion(wb.xls, name = "AA")), "Error detected in cell")
    expect_false(is(res, "try-error"))
    target <- data.frame(
        A = c("aa", "bb", "cc", NA, "ee", "ff"),
        stringsAsFactors = FALSE
    )
    attr(target, "worksheetScope") <- ""
    expect_equal(target, res)
    expect_warning(res <- try(readNamedRegion(wb.xls, name = "BB")), "Error detected in cell")
    expect_false(is(res, "try-error"))
    target <- data.frame(
        B = c(4.3, NA, -2.5, 1.6, NA, 9.7),
        stringsAsFactors = FALSE
    )
    attr(target, "worksheetScope") <- ""
    expect_equal(target, res)
    options(XLConnect.setCustomAttributes = FALSE)
    expect_warning(res <- try(readNamedRegion(wb.xls, name = "CC")), "Error detected in cell")
    expect_false(is(res, "try-error"))
    target <- data.frame(
        C = c(-53.2, NA, 34.1, -37.89, 0, 1.6),
        stringsAsFactors = FALSE
    )
    expect_equal(target, res)
    expect_warning(res <- try(readNamedRegion(wb.xls, name = "DD")), "Error detected in cell")
    expect_false(is(res, "try-error"))
    target <- data.frame(D = c(8.2, 2, 1, -0.5, NA, 3.1), stringsAsFactors = FALSE)
    expect_equal(target, res)
    expect_warning(res <- try(readNamedRegion(wb.xls, name = "EE")), "Error when trying to evaluate cell")
    expect_false(is(res, "try-error"))
    target <- data.frame(
        E = c("zz", "yy", NA, "ww", "vv", "uu"),
        stringsAsFactors = FALSE
    )
    expect_equal(target, res)
    onErrorCell(wb.xlsx, XLC$ERROR.WARN)
    expect_warning(res <- try(readNamedRegion(wb.xlsx, name = "AA")), "Error when trying to evaluate cell")
    expect_false(is(res, "try-error"))
    target <- data.frame(
        A = c("aa", "bb", "cc", NA, "ee", "ff"),
        stringsAsFactors = FALSE
    )
    expect_equal(target, res)
    expect_warning(res <- try(readNamedRegion(wb.xlsx, name = "BB")), "Error detected in cell")
    expect_false(is(res, "try-error"))
    target <- data.frame(
        B = c(4.3, NA, -2.5, 1.6, NA, 9.7),
        stringsAsFactors = FALSE
    )
    expect_equal(target, res)
    expect_warning(res <- try(readNamedRegion(wb.xlsx, name = "CC")), "Error detected in cell")
    expect_false(is(res, "try-error"))
    target <- data.frame(
        C = c(-53.2, NA, 34.1, -37.89, 0, 1.6),
        stringsAsFactors = FALSE
    )
    expect_equal(target, res)
    expect_warning(res <- try(readNamedRegion(wb.xlsx, name = "DD")), "Error detected in cell")
    expect_false(is(res, "try-error"))
    target <- data.frame(D = c(8.2, 2, 1, -0.5, NA, 3.1), stringsAsFactors = FALSE)
    expect_equal(target, res)
    expect_warning(res <- try(readNamedRegion(wb.xlsx, name = "EE")), "Error when trying to evaluate cell")
    expect_false(is(res, "try-error"))
    target <- data.frame(
        E = c("zz", "yy", NA, "ww", "vv", "uu"),
        stringsAsFactors = FALSE
    )
    expect_equal(target, res)
    options(XLConnect.setCustomAttributes = TRUE)
    onErrorCell(wb.xls, XLC$ERROR.STOP)
    res <- try(readNamedRegion(wb.xls, name = "AA"))
    expect_true(is(res, "try-error"))
    res <- try(readNamedRegion(wb.xls, name = "BB"))
    expect_true(is(res, "try-error"))
    res <- try(readNamedRegion(wb.xls, name = "CC"))
    expect_true(is(res, "try-error"))
    res <- try(readNamedRegion(wb.xls, name = "DD"))
    expect_true(is(res, "try-error"))
    res <- try(readNamedRegion(wb.xls, name = "EE"))
    expect_true(is(res, "try-error"))
    onErrorCell(wb.xlsx, XLC$ERROR.STOP)
    res <- try(readNamedRegion(wb.xlsx, name = "AA"))
    expect_true(is(res, "try-error"))
    res <- try(readNamedRegion(wb.xlsx, name = "BB"))
    expect_true(is(res, "try-error"))
    res <- try(readNamedRegion(wb.xlsx, name = "CC"))
    expect_true(is(res, "try-error"))
    res <- try(readNamedRegion(wb.xlsx, name = "DD"))
    expect_true(is(res, "try-error"))
    res <- try(readNamedRegion(wb.xlsx, name = "EE"))
    expect_true(is(res, "try-error"))
})
