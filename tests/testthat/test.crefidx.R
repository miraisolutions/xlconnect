test_that("test.crefidx", {
    target <- matrix(c(5, 8, 14, 38), ncol = 2, byrow = TRUE)
    expect_equal(target, cref2idx(c("$H$5", "$AL$14")))
    expect_equal(c("$H$5", "$AL$14"), idx2cref(c(5, 8, 14, 38)))
    expect_equal(c("KRE3799", "J26789", "DX357"), idx2cref(cref2idx(c("$KRE3799", 
        "J$26789", "$DX$357")), absRow = FALSE, absCol = FALSE))
    x <- c(36, 43, 25, 13, 356, 46)
    target <- matrix(x, ncol = 2, byrow = TRUE)
    expect_equal(target, cref2idx(idx2cref(x)))
})

