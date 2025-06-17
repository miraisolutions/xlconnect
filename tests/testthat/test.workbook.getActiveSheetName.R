test_that("test.workbook.getActiveSheetName", {
    wb.xls <- loadWorkbook(rsrc("resources/testWorkbookActiveSheetIndexAndName.xls"), 
        create = FALSE)
    wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookActiveSheetIndexAndName.xlsx"), 
        create = FALSE)
    expect_true(getActiveSheetName(wb.xls) == "Fifth Sheet")
    expect_true(getActiveSheetName(wb.xlsx) == "Fifth Sheet")
})

