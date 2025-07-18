test_that("extractSheetName extracts sheet names from formulas", {
    formulas <- c("MySheet!$A$1", "'My Sheet'!$A$1", "'My!Sheet'!$A$1", "MySheet")
    expected <- c("MySheet", "My Sheet", "My!Sheet", "")
    expect_equal(extractSheetName(formulas), expected)
})

