test_that("test.workbook.getLastRow", {
    wb.xls <- loadWorkbook("resources/testWorkbookGetLastRow.xls"),
        create = FALSE)
    wb.xlsx <- loadWorkbook("resources/testWorkbookGetLastRow.xlsx"),
        create = FALSE)
    expect_equal(33, as.vector(getLastRow(wb.xls, "mtcars")))
    expect_equal(41, as.vector(getLastRow(wb.xls, "mtcars2")))
    expect_equal(49, as.vector(getLastRow(wb.xls, "mtcars3")))
    expect_equal(33, as.vector(getLastRow(wb.xlsx, "mtcars")))
    expect_equal(41, as.vector(getLastRow(wb.xlsx, "mtcars2")))
    expect_equal(49, as.vector(getLastRow(wb.xlsx, "mtcars3")))
    expect_error(getLastRow(wb.xls, "doesNotExist"))
    expect_error(getLastRow(wb.xlsx, "doesNotExist"))
    expect_equal(c(empty = 1), getLastRow(wb.xls, "empty"))
    expect_equal(c(empty = 1), getLastRow(wb.xlsx, "empty"))
})

