test_that("test.workbook.cloneSheet", {
    wb.xls <- loadWorkbook("resources/testWorkbookCloneSheet.xls",
        create = FALSE
    )
    wb.xlsx <- loadWorkbook("resources/testWorkbookCloneSheet.xlsx",
        create = FALSE
    )
    expect_error(cloneSheet(wb.xls, sheet = "Test1", name = "Clone1"), NA)
    expect_error(readWorksheet(wb.xls, sheet = "Clone1"))
    expect_error(cloneSheet(wb.xlsx, sheet = "Test1", name = "Clone1"), NA)
    expect_error(readWorksheet(wb.xlsx, sheet = "Clone1"))
    expect_error(cloneSheet(wb.xls, sheet = "NotThere", name = "MyClone"), NA)
    expect_error(cloneSheet(wb.xlsx, sheet = "NotThere", name = "MyClone"), NA)
    expect_error(cloneSheet(wb.xls, sheet = "Test1", name = "'illegal"), NA)
    expect_error(cloneSheet(wb.xlsx, sheet = "Test1", name = "'illegal"), NA)
})
