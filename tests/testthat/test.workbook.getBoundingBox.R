test_that("getBoundingBox returns correct dimensions when no row/column is specified (*.xls)", {
  wb.xls <- loadWorkbook("resources/testWorkbookReadWorksheet.xls", create = FALSE)
  dim1 <- matrix(c(17, 6, 25, 9), dimnames = list(c(), c("Test5")))
  res <- getBoundingBox(wb.xls, sheet = "Test5")
  expect_equal(dim1, res)
})

test_that("getBoundingBox returns correct dimensions when no row/column is specified (*.xlsx)", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookReadWorksheet.xlsx", create = FALSE)
  dim1 <- matrix(c(17, 6, 25, 9), dimnames = list(c(), c("Test5")))
  res <- getBoundingBox(wb.xlsx, sheet = "Test5")
  expect_equal(dim1, res)
})

test_that("getBoundingBox returns correct dimensions when start and end cells are specified (*.xls)", {
  wb.xls <- loadWorkbook("resources/testWorkbookReadWorksheet.xls", create = FALSE)
  dim2 <- matrix(c(17, 7, 24, 9), dimnames = list(c(), c("Test5")))
  res <- getBoundingBox(wb.xls, sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol = 9)
  expect_equal(dim2, res)
})

test_that("getBoundingBox returns correct dimensions when start and end cells are specified (*.xlsx)", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookReadWorksheet.xlsx", create = FALSE)
  dim2 <- matrix(c(17, 7, 24, 9), dimnames = list(c(), c("Test5")))
  res <- getBoundingBox(wb.xlsx, sheet = "Test5", startRow = 17, startCol = 7, endRow = 24, endCol = 9)
  expect_equal(dim2, res)
})

test_that("getBoundingBox returns correct dimensions when start and end columns are specified (*.xls)", {
  wb.xls <- loadWorkbook("resources/testWorkbookReadWorksheet.xls", create = FALSE)
  dim3 <- matrix(c(17, 7, 25, 9), dimnames = list(c(), c("Test5")))
  res <- getBoundingBox(wb.xls, sheet = "Test5", startCol = 7, endCol = 9)
  expect_equal(dim3, res)
})

test_that("getBoundingBox returns correct dimensions when start and end columns are specified (*.xlsx)", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookReadWorksheet.xlsx", create = FALSE)
  dim3 <- matrix(c(17, 7, 25, 9), dimnames = list(c(), c("Test5")))
  res <- getBoundingBox(wb.xlsx, sheet = "Test5", startCol = 7, endCol = 9)
  expect_equal(dim3, res)
})

test_that("getBoundingBox returns correct dimensions when only the end row is specified (*.xls)", {
  wb.xls <- loadWorkbook("resources/testWorkbookReadWorksheet.xls", create = FALSE)
  dim4 <- matrix(c(17, 6, 24, 9), dimnames = list(c(), c("Test5")))
  res <- getBoundingBox(wb.xls, sheet = "Test5", endRow = 24)
  expect_equal(dim4, res)
})

test_that("getBoundingBox returns correct dimensions when only the end row is specified (*.xlsx)", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookReadWorksheet.xlsx", create = FALSE)
  dim4 <- matrix(c(17, 6, 24, 9), dimnames = list(c(), c("Test5")))
  res <- getBoundingBox(wb.xlsx, sheet = "Test5", endRow = 24)
  expect_equal(dim4, res)
})

test_that("getBoundingBox returns correct dimensions when multiple sheets are specified (*.xls)", {
  wb.xls <- loadWorkbook("resources/testWorkbookReadWorksheet.xls", create = FALSE)
  dim5 <- matrix(
    c(11, 6, 16, 9, 8, 4, 16, 7, 17, 6, 25, 9),
    ncol = 3,
    dimnames = list(c(), c("Test1", "Test4", "Test5"))
  )
  res <- getBoundingBox(wb.xls, sheet = c("Test1", "Test4", "Test5"))
  expect_equal(dim5, res)
})

test_that("getBoundingBox returns correct dimensions when multiple sheets are specified (*.xlsx)", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookReadWorksheet.xlsx", create = FALSE)
  dim5 <- matrix(
    c(11, 6, 16, 9, 8, 4, 16, 7, 17, 6, 25, 9),
    ncol = 3,
    dimnames = list(c(), c("Test1", "Test4", "Test5"))
  )
  res <- getBoundingBox(wb.xlsx, sheet = c("Test1", "Test4", "Test5"))
  expect_equal(dim5, res)
})
