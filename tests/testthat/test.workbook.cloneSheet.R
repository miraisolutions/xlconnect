test_that("test.workbook.cloneSheet", {
    wb.xls <- loadWorkbook("resources/testWorkbookCloneSheet.xls"),
        create = FALSE)
    wb.xlsx <- loadWorkbook("resources/testWorkbookCloneSheet.xlsx"),
        create = FALSE)
    checkNoException(cloneSheet(wb.xls, sheet = "Test1", name = "Clone1"))
    checkNoException(readWorksheet(wb.xls, sheet = "Clone1"))
    checkNoException(cloneSheet(wb.xlsx, sheet = "Test1", name = "Clone1"))
    checkNoException(readWorksheet(wb.xlsx, sheet = "Clone1"))
    expect_error(cloneSheet(wb.xls, sheet = "NotThere", name = "MyClone"))
    expect_error(cloneSheet(wb.xlsx, sheet = "NotThere", name = "MyClone"))
    expect_error(cloneSheet(wb.xls, sheet = "Test1", name = "'illegal"))
    expect_error(cloneSheet(wb.xlsx, sheet = "Test1", name = "'illegal"))
})

