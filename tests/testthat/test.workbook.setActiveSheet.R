test_that("test.workbook.setActiveSheet", {
    wb.xls <- loadWorkbook("resources/testWorkbookSetActiveSheet.xls",
        create = FALSE
    )
    wb.xlsx <- loadWorkbook("resources/testWorkbookSetActiveSheet.xlsx",
        create = FALSE
    )
    setActiveSheet(wb.xls, 1)
    expect_true(getActiveSheetIndex(wb.xls) == 1)
    setActiveSheet(wb.xls, 3)
    expect_true(getActiveSheetIndex(wb.xls) == 3)
    setActiveSheet(wb.xls, "Sheet2")
    expect_true(getActiveSheetIndex(wb.xls) == 2)
    setActiveSheet(wb.xlsx, 1)
    expect_true(getActiveSheetIndex(wb.xlsx) == 1)
    setActiveSheet(wb.xlsx, 3)
    expect_true(getActiveSheetIndex(wb.xlsx) == 3)
    setActiveSheet(wb.xlsx, "Sheet2")
    expect_true(getActiveSheetIndex(wb.xlsx) == 2)
    expect_error(setActiveSheet(wb.xls, 19), "IllegalArgumentException")
    expect_error(setActiveSheet(wb.xls, "SheetWhichDoesNotExist"), "IllegalArgumentException")
    expect_error(setActiveSheet(wb.xlsx, 19), "IllegalArgumentException")
    expect_error(setActiveSheet(wb.xlsx, "SheetWhichDoesNotExist"), "IllegalArgumentException")
})
