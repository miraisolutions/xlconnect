test_that("test.workbook.readTable", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookReadTable.xlsx"),
        create = FALSE)
    checkDf <- data.frame(NumericColumn = c(-23.63, NA, NA, 5.8, 
        3), StringColumn = c("Hello", NA, NA, NA, "World"), BooleanColumn = c(TRUE, 
        FALSE, FALSE, NA, NA), DateTimeColumn = as.POSIXct(c(NA, 
        NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07")), 
        stringsAsFactors = F)
    res <- readTable(wb.xlsx, sheet = "Test", table = "TestTable1")
    expect_equal(checkDf, res)
})

