test_that("getTables returns a simplified list of tables", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookReadTable.xlsx", create = FALSE)
    res <- getTables(wb.xlsx, sheet = "Test", simplify = TRUE)
    expect_equal("TestTable1", res)
})

test_that("getTables returns a nested list of tables", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookReadTable.xlsx", create = FALSE)
    res <- getTables(wb.xlsx, sheet = "Test", simplify = FALSE)
    expect_equal(list(Test = "TestTable1"), res)
})

test_that("getTables returns an empty list for a sheet with no tables", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookReadTable.xlsx", create = FALSE)
    res <- getTables(wb.xlsx, sheet = "NoTableHere", simplify = TRUE)
    expect_equal(character(0), res)
})

test_that("getTables throws an error for a non-existent sheet", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookReadTable.xlsx", create = FALSE)
    expect_error(getTables(wb.xlsx, sheet = "DoesNotExist"))
})
