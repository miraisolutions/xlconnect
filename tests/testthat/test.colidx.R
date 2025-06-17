test_that("test.colidx", {
    expect_equal(c(1, 27, 6102), col2idx(c("A", "AA", "HZR")))
    expect_equal(c("A", "AA", "HZR"), idx2col(c(1, 27, 6102)))
    expect_equal(c("AWT", "FRT"), idx2col(col2idx(c("AWT", "FRT"))))
    expect_equal(c(3628, 867), col2idx(idx2col(c(3628, 867))))
    expect_equal(rep("", 3), idx2col(c(0, -1, -2628)))
})

