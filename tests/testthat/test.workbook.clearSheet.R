test_that("test.workbook.clearSheet", {
    wb.xls <- loadWorkbook("resources/testWorkbookClearCells.xls",
        create = FALSE
    )
    wb.xlsx <- loadWorkbook("resources/testWorkbookClearCells.xlsx",
        create = FALSE
    )
    clearSheet(wb.xls, c("clearSheet1", "clearSheet2"))
    res1 <- getLastRow(wb.xls, "clearSheet1")
    res2 <- getLastRow(wb.xls, "clearSheet2")
    expect_equal(c(clearSheet1 = 1, clearSheet2 = 1), c(
        res1,
        res2
    ))
    clearSheet(wb.xlsx, c("clearSheet1", "clearSheet2"))
    res1 <- getLastRow(wb.xlsx, "clearSheet1")
    res2 <- getLastRow(wb.xlsx, "clearSheet2")
    expect_equal(c(clearSheet1 = 1, clearSheet2 = 1), c(
        res1,
        res2
    ))
})
