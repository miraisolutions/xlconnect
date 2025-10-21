test_that("col2idx converts column names to indices", {
  expect_equal(col2idx(c("A", "AA", "HZR")), c(1, 27, 6102))
})

test_that("idx2col converts indices to column names", {
  expect_equal(idx2col(c(1, 27, 6102)), c("A", "AA", "HZR"))
})

test_that("idx2col and col2idx are inverse functions", {
  expect_equal(idx2col(col2idx(c("AWT", "FRT"))), c("AWT", "FRT"))
  expect_equal(col2idx(idx2col(c(3628, 867))), c(3628, 867))
})

test_that("idx2col handles invalid indices by returning empty strings", {
  expect_equal(idx2col(c(0, -1, -2628)), rep("", 3))
})
