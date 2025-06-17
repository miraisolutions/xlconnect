test_that("test.workbook.getActiveSheetIndex", {
    wb.xls <- loadWorkbook(rsrc("resources/testWorkbookActiveSheetIndexAndName.xls"), 
        create = FALSE)
    wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookActiveSheetIndexAndName.xlsx"), 
        create = FALSE)
    expect_true(getActiveSheetIndex(wb.xls) == 5)
    expect_true(getActiveSheetIndex(wb.xlsx) == 5)
})

