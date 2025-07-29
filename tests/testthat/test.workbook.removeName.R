test_that("removeName removes an existing name in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookRemoveName.xls", create = FALSE)
    removeName(wb.xls, "AA")
    expect_false(existsName(wb.xls, "AA"))
})

test_that("removeName removes an existing name in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookRemoveName.xlsx", create = FALSE)
    removeName(wb.xlsx, "AA")
    expect_false(existsName(wb.xlsx, "AA"))
})

test_that("removeName does not throw an error for a non-existent name in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookRemoveName.xls", create = FALSE)
    expect_error(removeName(wb.xls, "NameWhichDoesNotExist"), NA)
})

test_that("removeName does not throw an error for a non-existent name in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookRemoveName.xlsx", create = FALSE)
    expect_error(removeName(wb.xlsx, "NameWhichDoesNotExist"), NA)
})
