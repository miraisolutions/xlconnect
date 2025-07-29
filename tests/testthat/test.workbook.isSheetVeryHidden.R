test_that("isSheetVeryHidden identifies very hidden sheets correctly in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookHiddenSheets.xls", create = FALSE)
    expect_true(isSheetVeryHidden(wb.xls, 4))
    expect_true(isSheetVeryHidden(wb.xls, "DDD"))
})

test_that("isSheetVeryHidden identifies non-very hidden sheets correctly in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookHiddenSheets.xls", create = FALSE)
    expect_false(isSheetVeryHidden(wb.xls, 1))
    expect_false(isSheetVeryHidden(wb.xls, "AAA"))
    expect_false(isSheetVeryHidden(wb.xls, 2))
    expect_false(isSheetVeryHidden(wb.xls, "BBB"))
    expect_false(isSheetVeryHidden(wb.xls, 3))
    expect_false(isSheetVeryHidden(wb.xls, "CCC"))
})

test_that("isSheetVeryHidden throws errors for invalid sheets in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookHiddenSheets.xls", create = FALSE)
    expect_error(isSheetVeryHidden(wb.xls, 200))
    expect_error(isSheetVeryHidden(wb.xls, "Sheet does not exist"))
    expect_error(isSheetVeryHidden(wb.xls, "'Illegal sheet name"))
})

test_that("isSheetVeryHidden identifies very hidden sheets correctly in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookHiddenSheets.xlsx", create = FALSE)
    expect_true(isSheetVeryHidden(wb.xlsx, 4))
    expect_true(isSheetVeryHidden(wb.xlsx, "DDD"))
})

test_that("isSheetVeryHidden identifies non-very hidden sheets correctly in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookHiddenSheets.xlsx", create = FALSE)
    expect_false(isSheetVeryHidden(wb.xlsx, 1))
    expect_false(isSheetVeryHidden(wb.xlsx, "AAA"))
    expect_false(isSheetVeryHidden(wb.xlsx, 2))
    expect_false(isSheetVeryHidden(wb.xlsx, "BBB"))
    expect_false(isSheetVeryHidden(wb.xlsx, 3))
    expect_false(isSheetVeryHidden(wb.xlsx, "CCC"))
})

test_that("isSheetVeryHidden throws errors for invalid sheets in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookHiddenSheets.xlsx", create = FALSE)
    expect_error(isSheetVeryHidden(wb.xlsx, 200))
    expect_error(isSheetVeryHidden(wb.xlsx, "Sheet does not exist"))
    expect_error(isSheetVeryHidden(wb.xlsx, "'Illegal sheet name"))
})
