test_that("test.workbook.cloneSheet", {
    wb.xls <- loadWorkbook("resources/testWorkbookCloneSheet.xls",
        create = FALSE
    )
    wb.xlsx <- loadWorkbook("resources/testWorkbookCloneSheet.xlsx",
        create = FALSE
    )
    # cloneSheet does not throw an error if successful
    cloneSheet(wb.xls, sheet = "Test1", name = "Clone1")
    # readWorksheet does not throw an error if successful
    readWorksheet(wb.xls, sheet = "Clone1")
    # cloneSheet does not throw an error if successful
    cloneSheet(wb.xlsx, sheet = "Test1", name = "Clone1")
    # readWorksheet does not throw an error if successful
    readWorksheet(wb.xlsx, sheet = "Clone1")
    expect_error(cloneSheet(wb.xls, sheet = "NotThere", name = "MyClone"), "IllegalArgumentException")
    expect_error(cloneSheet(wb.xlsx, sheet = "NotThere", name = "MyClone"), "IllegalArgumentException")
    expect_error(cloneSheet(wb.xls, sheet = "Test1", name = "'illegal"), "IllegalArgumentException")
    expect_error(cloneSheet(wb.xlsx, sheet = "Test1", name = "'illegal"), "IllegalArgumentException")
})
