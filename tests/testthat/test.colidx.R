test_that("col2idx converts column names to indices", {
  expect_equal(c(1, 27, 6102), col2idx(c("A", "AA", "HZR")))
})

test_that("idx2col converts indices to column names", {
  expect_equal(c("A", "AA", "HZR"), idx2col(c(1, 27, 6102)))
})

test_that("idx2col and col2idx are inverse functions", {
  expect_equal(c("AWT", "FRT"), idx2col(col2idx(c("AWT", "FRT"))))
  expect_equal(c(3628, 867), col2idx(idx2col(c(3628, 867))))
})

test_that("idx2col handles invalid indices by returning empty strings", {
  expect_equal(rep("", 3), idx2col(c(0, -1, -2628)))
})
