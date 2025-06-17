test_that("test.workbook.unhideSheet", {
    wb.xls <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xls"), 
        create = FALSE)
    wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xlsx"), 
        create = FALSE)
    unhideSheet(wb.xls, 2)
    unhideSheet(wb.xls, "DDD")
    expect_false(isSheetHidden(wb.xls, 2))
    expect_false(isSheetVeryHidden(wb.xls, "DDD"))
    unhideSheet(wb.xlsx, 2)
    unhideSheet(wb.xlsx, "DDD")
    expect_false(isSheetHidden(wb.xlsx, 2))
    expect_false(isSheetVeryHidden(wb.xlsx, "DDD"))
    expect_error(unhideSheet(wb.xls, 58))
    expect_error(unhideSheet(wb.xls, "SheetWhichDoesNotExist"))
    expect_error(unhideSheet(wb.xlsx, 58))
    expect_error(unhideSheet(wb.xlsx, "SheetWhichDoesNotExist"))
})

