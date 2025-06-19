test_that("test.workbook.getSheetPos", {
    wb.xls <- loadWorkbook("resources/testWorkbookGetSheetPos.xls"),
        create = TRUE)
    wb.xlsx <- loadWorkbook("resources/testWorkbookGetSheetPos.xlsx"),
        create = TRUE)
    createSheet(wb.xls, c("Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4"))
    createSheet(wb.xlsx, c("Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4"))
    expect_equal(c(`Sheet 3` = 3, `Sheet 2` = 2, `Sheet 4` = 4, 
        `Sheet 1` = 1), getSheetPos(wb.xls, c("Sheet 3", "Sheet 2", 
        "Sheet 4", "Sheet 1")))
    expect_equal(c(`Sheet 3` = 3, `Sheet 2` = 2, `Sheet 4` = 4, 
        `Sheet 1` = 1), getSheetPos(wb.xlsx, c("Sheet 3", "Sheet 2", 
        "Sheet 4", "Sheet 1")))
    expect_equal(c(NotExisting = 0), getSheetPos(wb.xls, "NotExisting"))
    expect_equal(0, as.vector(getSheetPos(wb.xls, "%#?%+?[-")))
    expect_equal(c(NotExisting = 0), getSheetPos(wb.xlsx, "NotExisting"))
    expect_equal(0, as.vector(getSheetPos(wb.xlsx, "%#?%+?[-")))
})

