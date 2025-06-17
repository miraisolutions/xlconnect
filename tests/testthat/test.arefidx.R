test_that("test.arefidx", {
    target <- matrix(c(1, 1, 8, 2, 9, 4, 26, 11), ncol = 4, byrow = TRUE)
    expect_equal(target, aref2idx(c("A1:B8", "D9:K26")))
    expect_equal(c("B3:I7", "E8:Q48"), idx2aref(c(3, 2, 7, 9, 
        8, 5, 48, 17)))
    expect_equal(c("D27:J54", "AA23:CD129"), idx2aref(aref2idx(c("D27:J54", 
        "AA23:CD129"))))
    x <- c(31, 6, 56, 8, 129, 17, 488, 37)
    target <- matrix(x, ncol = 4, byrow = TRUE)
    expect_equal(target, aref2idx(idx2aref(x)))
    expect_equal("BB35:BZ712", aref("BB35", c(678, 25)))
    expect_equal("AT18:BK33", aref(c(18, 46), c(16, 18)))
})

