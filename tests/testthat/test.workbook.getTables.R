test_that("test.workbook.getTables", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookReadTable.xlsx"),
        create = FALSE)
    res <- getTables(wb.xlsx, sheet = "Test", simplify = TRUE)
    expect_equal("TestTable1", res)
    res <- getTables(wb.xlsx, sheet = "Test", simplify = FALSE)
    expect_equal(list(Test = "TestTable1"), res)
    res <- getTables(wb.xlsx, sheet = "NoTableHere", simplify = TRUE)
    expect_equal(character(0), res)
    expect_error(getTables(wb.xlsx, sheet = "DoesNotExist"))
})

