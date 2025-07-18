test_that("isSheetHidden identifies hidden sheets correctly in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookHiddenSheets.xls", create = FALSE)
    expect_true(isSheetHidden(wb.xls, 2))
    expect_true(isSheetHidden(wb.xls, "BBB"))
})

test_that("isSheetHidden identifies non-hidden sheets correctly in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookHiddenSheets.xls", create = FALSE)
    expect_false(isSheetHidden(wb.xls, 1))
    expect_false(isSheetHidden(wb.xls, "AAA"))
    expect_false(isSheetHidden(wb.xls, 3))
    expect_false(isSheetHidden(wb.xls, "CCC"))
    expect_false(isSheetHidden(wb.xls, 4))
    expect_false(isSheetHidden(wb.xls, "DDD"))
})

test_that("isSheetHidden throws errors for invalid sheets in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookHiddenSheets.xls", create = FALSE)
    expect_error(isSheetHidden(wb.xls, 200))
    expect_error(isSheetHidden(wb.xls, "Sheet does not exist"))
    expect_error(isSheetHidden(wb.xls, "'Illegal sheet name"))
})

test_that("isSheetHidden identifies hidden sheets correctly in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookHiddenSheets.xlsx", create = FALSE)
    expect_true(isSheetHidden(wb.xlsx, 2))
    expect_true(isSheetHidden(wb.xlsx, "BBB"))
})

test_that("isSheetHidden identifies non-hidden sheets correctly in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookHiddenSheets.xlsx", create = FALSE)
    expect_false(isSheetHidden(wb.xlsx, 1))
    expect_false(isSheetHidden(wb.xlsx, "AAA"))
    expect_false(isSheetHidden(wb.xlsx, 3))
    expect_false(isSheetHidden(wb.xlsx, "CCC"))
    expect_false(isSheetHidden(wb.xlsx, 4))
    expect_false(isSheetHidden(wb.xlsx, "DDD"))
})

test_that("isSheetHidden throws errors for invalid sheets in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookHiddenSheets.xlsx", create = FALSE)
    expect_error(isSheetHidden(wb.xlsx, 200))
    expect_error(isSheetHidden(wb.xlsx, "Sheet does not exist"))
    expect_error(isSheetHidden(wb.xlsx, "'Illegal sheet name"))
})
