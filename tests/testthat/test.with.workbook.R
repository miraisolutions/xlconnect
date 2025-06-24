test_that("test.with.workbook", {
    wb.xls <- loadWorkbook("resources/testWithWorkbook.xls",
        create = FALSE
    )
    wb.xlsx <- loadWorkbook("resources/testWithWorkbook.xlsx",
        create = FALSE
    )
    with(wb.xls,
        {
            expect_true(all(dim(AA) == c(8, 3)))
            expect_true(all(dim(BB) == c(5, 5)))
        },
        header = FALSE
    )
    with(wb.xlsx,
        {
            expect_true(all(dim(AA) == c(8, 3)))
            expect_true(all(dim(BB) == c(5, 5)))
        },
        header = FALSE
    )
})
