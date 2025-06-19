test_that("test.workbook.removeSheet", {
    wb.xls <- loadWorkbook("resources/testWorkbookRemoveSheet.xls"),
        create = FALSE)
    wb.xlsx <- loadWorkbook("resources/testWorkbookRemoveSheet.xlsx"),
        create = FALSE)
    removeSheet(wb.xls, "BBB")
    expect_false(existsSheet(wb.xls, "BBB"))
    removeSheet(wb.xlsx, "BBB")
    expect_false(existsSheet(wb.xlsx, "BBB"))
    checkNoException(removeSheet(wb.xls, 35))
    checkNoException(removeSheet(wb.xls, "SheetWhichDoesNotExist"))
    checkNoException(removeSheet(wb.xlsx, 35))
    checkNoException(removeSheet(wb.xlsx, "SheetWhichDoesNotExist"))
})

