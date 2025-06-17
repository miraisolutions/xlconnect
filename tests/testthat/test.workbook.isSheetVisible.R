test_that("test.workbook.isSheetVisible", {
    wb.xls <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xls"), 
        create = FALSE)
    wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xlsx"), 
        create = FALSE)
    expect_false(isSheetVisible(wb.xls, 2))
    expect_false(isSheetVisible(wb.xls, "BBB"))
    expect_true(isSheetVisible(wb.xls, 1))
    expect_true(isSheetVisible(wb.xls, "AAA"))
    expect_true(isSheetVisible(wb.xls, 3))
    expect_true(isSheetVisible(wb.xls, "CCC"))
    expect_false(isSheetVisible(wb.xls, 4))
    expect_false(isSheetVisible(wb.xls, "DDD"))
    expect_false(isSheetVisible(wb.xlsx, 2))
    expect_false(isSheetVisible(wb.xlsx, "BBB"))
    expect_true(isSheetVisible(wb.xlsx, 1))
    expect_true(isSheetVisible(wb.xlsx, "AAA"))
    expect_true(isSheetVisible(wb.xlsx, 3))
    expect_true(isSheetVisible(wb.xlsx, "CCC"))
    expect_false(isSheetVisible(wb.xlsx, 4))
    expect_false(isSheetVisible(wb.xlsx, "DDD"))
})

