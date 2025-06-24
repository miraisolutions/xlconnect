test_that("test.workbook.removeName", {
    wb.xls <- loadWorkbook("resources/testWorkbookRemoveName.xls",
        create = FALSE
    )
    wb.xlsx <- loadWorkbook("resources/testWorkbookRemoveName.xlsx",
        create = FALSE
    )
    removeName(wb.xls, "AA")
    expect_false(existsName(wb.xls, "AA"))
    removeName(wb.xlsx, "AA")
    expect_false(existsName(wb.xlsx, "AA"))
    expect_error(removeName(wb.xls, "NameWhichDoesNotExist"), NA)
    expect_error(removeName(wb.xlsx, "NameWhichDoesNotExist"), NA)
})
