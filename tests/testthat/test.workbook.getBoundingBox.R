test_that("getBoundingBox returns correct dimensions for a single sheet in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookReadWorksheet.xls", create = FALSE)
    dim1 <- matrix(c(17, 6, 25, 9), dimnames = list(c(), c("Test5")))
    res <- getBoundingBox(wb.xls, sheet = "Test5")
    expect_equal(dim1, res)
})

test_that("getBoundingBox returns correct dimensions for a single sheet in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookReadWorksheet.xlsx", create = FALSE)
    dim1 <- matrix(c(17, 6, 25, 9), dimnames = list(c(), c("Test5")))
    res <- getBoundingBox(wb.xlsx, sheet = "Test5")
    expect_equal(dim1, res)
})

test_that("getBoundingBox returns correct dimensions for a specified region in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookReadWorksheet.xls", create = FALSE)
    dim2 <- matrix(c(17, 7, 24, 9), dimnames = list(c(), c("Test5")))
    res <- getBoundingBox(wb.xls, sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol = 9)
    expect_equal(dim2, res)
})

test_that("getBoundingBox returns correct dimensions for a specified region in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookReadWorksheet.xlsx", create = FALSE)
    dim2 <- matrix(c(17, 7, 24, 9), dimnames = list(c(), c("Test5")))
    res <- getBoundingBox(wb.xlsx, sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol = 9)
    expect_equal(dim2, res)
})

test_that("getBoundingBox returns correct dimensions with start and end columns in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookReadWorksheet.xls", create = FALSE)
    dim3 <- matrix(c(17, 7, 25, 9), dimnames = list(c(), c("Test5")))
    res <- getBoundingBox(wb.xls, sheet = "Test5", startCol = 7, endCol = 9)
    expect_equal(dim3, res)
})

test_that("getBoundingBox returns correct dimensions with start and end columns in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookReadWorksheet.xlsx", create = FALSE)
    dim3 <- matrix(c(17, 7, 25, 9), dimnames = list(c(), c("Test5")))
    res <- getBoundingBox(wb.xlsx, sheet = "Test5", startCol = 7, endCol = 9)
    expect_equal(dim3, res)
})

test_that("getBoundingBox returns correct dimensions with end row in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookReadWorksheet.xls", create = FALSE)
    dim4 <- matrix(c(17, 6, 24, 9), dimnames = list(c(), c("Test5")))
    res <- getBoundingBox(wb.xls, sheet = "Test5", endRow = 24)
    expect_equal(dim4, res)
})

test_that("getBoundingBox returns correct dimensions with end row in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookReadWorksheet.xlsx", create = FALSE)
    dim4 <- matrix(c(17, 6, 24, 9), dimnames = list(c(), c("Test5")))
    res <- getBoundingBox(wb.xlsx, sheet = "Test5", endRow = 24)
    expect_equal(dim4, res)
})

test_that("getBoundingBox returns correct dimensions for multiple sheets in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookReadWorksheet.xls", create = FALSE)
    dim5 <- matrix(c(11, 6, 16, 9, 8, 4, 16, 7, 17, 6, 25, 9), ncol = 3, dimnames = list(c(), c("Test1", "Test4", "Test5")))
    res <- getBoundingBox(wb.xls, sheet = c("Test1", "Test4", "Test5"))
    expect_equal(dim5, res)
})

test_that("getBoundingBox returns correct dimensions for multiple sheets in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookReadWorksheet.xlsx", create = FALSE)
    dim5 <- matrix(c(11, 6, 16, 9, 8, 4, 16, 7, 17, 6, 25, 9), ncol = 3, dimnames = list(c(), c("Test1", "Test4", "Test5")))
    res <- getBoundingBox(wb.xlsx, sheet = c("Test1", "Test4", "Test5"))
    expect_equal(dim5, res)
})
