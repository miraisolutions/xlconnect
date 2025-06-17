test_that("test.workbook.getReferenceCoordinatesForTable", {
    wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookReadTable.xlsx"), 
        create = FALSE)
    res <- getReferenceCoordinatesForTable(wb.xlsx, sheet = "Test", 
        table = "TestTable1")
    expect_equal(matrix(c(5, 10, 4, 7), ncol = 2), res)
})

