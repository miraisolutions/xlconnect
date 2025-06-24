test_that("test.workbook.getBoundingBox", {
    wb.xls <- loadWorkbook("resources/testWorkbookReadWorksheet.xls",
        create = FALSE
    )
    wb.xlsx <- loadWorkbook("resources/testWorkbookReadWorksheet.xlsx",
        create = FALSE
    )
    dim1 <- matrix(c(17, 6, 25, 9), dimnames = list(c(), c("Test5")))
    dim2 <- matrix(c(17, 7, 24, 9), dimnames = list(c(), c("Test5")))
    dim3 <- matrix(c(17, 7, 25, 9), dimnames = list(c(), c("Test5")))
    dim4 <- matrix(c(17, 6, 24, 9), dimnames = list(c(), c("Test5")))
    dim5 <- matrix(c(11, 6, 16, 9, 8, 4, 16, 7, 17, 6, 25, 9),
        ncol = 3, dimnames = list(c(), c("Test1", "Test4", "Test5"))
    )
    res <- getBoundingBox(wb.xls, sheet = "Test5")
    expect_equal(dim1, res)
    res <- getBoundingBox(wb.xlsx, sheet = "Test5")
    expect_equal(dim1, res)
    res <- getBoundingBox(wb.xls,
        sheet = "Test5", startRow = 17,
        startCol = 7, endRow = 24, endCol = 9
    )
    expect_equal(dim2, res)
    res <- getBoundingBox(wb.xlsx,
        sheet = "Test5", startRow = 17,
        startCol = 7, endRow = 24, endCol = 9
    )
    expect_equal(dim2, res)
    res <- getBoundingBox(wb.xls,
        sheet = "Test5", startCol = 7,
        endCol = 9
    )
    expect_equal(dim3, res)
    res <- getBoundingBox(wb.xlsx,
        sheet = "Test5", startCol = 7,
        endCol = 9
    )
    expect_equal(dim3, res)
    res <- getBoundingBox(wb.xls, sheet = "Test5", endRow = 24)
    expect_equal(dim4, res)
    res <- getBoundingBox(wb.xlsx, sheet = "Test5", endRow = 24)
    expect_equal(dim4, res)
    res <- getBoundingBox(wb.xls, sheet = c(
        "Test1", "Test4",
        "Test5"
    ))
    expect_equal(dim5, res)
    res <- getBoundingBox(wb.xlsx, sheet = c(
        "Test1", "Test4",
        "Test5"
    ))
    expect_equal(dim5, res)
})
