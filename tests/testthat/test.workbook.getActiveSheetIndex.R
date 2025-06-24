test_that("test.workbook.getActiveSheetIndex", {
    wb.xls <- loadWorkbook("resources/testWorkbookActiveSheetIndexAndName.xls",
        create = FALSE
    )
    wb.xlsx <- loadWorkbook("resources/testWorkbookActiveSheetIndexAndName.xlsx",
        create = FALSE
    )
    expect_true(getActiveSheetIndex(wb.xls) == 5)
    expect_true(getActiveSheetIndex(wb.xlsx) == 5)
})
