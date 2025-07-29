test_that("existsName identifies global names in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookExistsNameAndSheet.xls", create = FALSE)
    res_xls_AA <- existsName(wb.xls, "AA")
    expect_true(res_xls_AA)
    xls_AA_scope <- attr(res_xls_AA, "worksheetScope")
    expect_true(is.null(xls_AA_scope) || xls_AA_scope == "")
    res_xls_BB <- existsName(wb.xls, "BB")
    expect_true(res_xls_BB)
    xls_BB_scope <- attr(res_xls_BB, "worksheetScope")
    expect_true(is.null(xls_BB_scope) || xls_BB_scope == "")
    res_xls_CC <- existsName(wb.xls, "CC")
    expect_true(res_xls_CC)
    xls_CC_scope <- attr(res_xls_CC, "worksheetScope")
    expect_true(is.null(xls_CC_scope) || xls_CC_scope == "")
})

test_that("existsName identifies non-existent and illegal names in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookExistsNameAndSheet.xls", create = FALSE)
    expect_false(existsName(wb.xls, "DD"))
    expect_false(existsName(wb.xls, "'illegal name"))
    expect_false(existsName(wb.xls, "%&$$-^~@afk20 235-??a?"))
})

test_that("existsName identifies global names in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookExistsNameAndSheet.xlsx", create = FALSE)
    res_xlsx_AA <- existsName(wb.xlsx, "AA")
    expect_true(res_xlsx_AA)
    xlsx_AA_scope <- attr(res_xlsx_AA, "worksheetScope")
    expect_true(is.null(xlsx_AA_scope) || xlsx_AA_scope == "")
    res_xlsx_BB <- existsName(wb.xlsx, "BB")
    expect_true(res_xlsx_BB)
    xlsx_BB_scope <- attr(res_xlsx_BB, "worksheetScope")
    expect_true(is.null(xlsx_BB_scope) || xlsx_BB_scope == "")
    res_xlsx_CC <- existsName(wb.xlsx, "CC")
    expect_true(res_xlsx_CC)
    xlsx_CC_scope <- attr(res_xlsx_CC, "worksheetScope")
    expect_true(is.null(xlsx_CC_scope) || xlsx_CC_scope == "")
})

test_that("existsName identifies non-existent and illegal names in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookExistsNameAndSheet.xlsx", create = FALSE)
    expect_false(existsName(wb.xlsx, "DD"))
    expect_false(existsName(wb.xlsx, "'illegal name"))
    expect_false(existsName(wb.xlsx, "%&$$-^~@afk20 235-??a?"))
})

test_that("existsName identifies sheet-scoped names in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookExistsNameAndSheet.xls", create = FALSE)
    res_xls_AA_1 <- existsName(wb.xls, "AA_1")
    expect_true(res_xls_AA_1)
    if (isTRUE(getOption("XLConnect.setCustomAttributes"))) {
        expect_equal(attr(res_xls_AA_1, "worksheetScope"), "AAA")
    }
    res_xls_BB_1 <- existsName(wb.xls, "BB_1")
    expect_true(res_xls_BB_1)
    if (isTRUE(getOption("XLConnect.setCustomAttributes"))) {
        expect_equal(attr(res_xls_BB_1, "worksheetScope"), "BBB")
    }
})

test_that("existsName identifies sheet-scoped names in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookExistsNameAndSheet.xlsx", create = FALSE)
    res_xlsx_AA_1 <- existsName(wb.xlsx, "AA_1")
    expect_true(res_xlsx_AA_1)
    if (isTRUE(getOption("XLConnect.setCustomAttributes"))) {
        expect_equal(attr(res_xlsx_AA_1, "worksheetScope"), "AAA")
    }
    res_xlsx_BB_1 <- existsName(wb.xlsx, "BB_1")
    expect_true(res_xlsx_BB_1)
    if (isTRUE(getOption("XLConnect.setCustomAttributes"))) {
        expect_equal(attr(res_xlsx_BB_1, "worksheetScope"), "BBB")
    }
})

test_that("existsName works when XLConnect.setCustomAttributes is FALSE", {
    wb.xls <- loadWorkbook("resources/testWorkbookExistsNameAndSheet.xls", create = FALSE)
    wb.xlsx <- loadWorkbook("resources/testWorkbookExistsNameAndSheet.xlsx", create = FALSE)
    options(XLConnect.setCustomAttributes = FALSE)
    expect_true(existsName(wb.xls, "AA_1"))
    expect_true(existsName(wb.xls, "BB_1"))
    expect_true(existsName(wb.xlsx, "AA_1"))
    expect_true(existsName(wb.xlsx, "BB_1"))
    options(XLConnect.setCustomAttributes = TRUE)
})
