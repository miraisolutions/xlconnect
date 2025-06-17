test_that("test.extractSheetName", {
    expect_equal(c("MySheet", "My Sheet", "My!Sheet", ""), extractSheetName(c("MySheet!$A$1", 
        "'My Sheet'!$A$1", "'My!Sheet'!$A$1", "MySheet")))
})

