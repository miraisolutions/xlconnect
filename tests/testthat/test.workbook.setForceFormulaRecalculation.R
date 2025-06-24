test_that("test.workbook.setForceFormulaRecalculation", {
    wb.xlsx <- loadWorkbook("resources/testBug170.xlsx",
        create = FALSE
    )
    setForceFormulaRecalculation(wb.xlsx, 1, TRUE)
    expect_true(getForceFormulaRecalculation(wb.xlsx, 1))
    setForceFormulaRecalculation(
        wb.xlsx, c("Sheet1", "Sheet2"),
        FALSE
    )
    expect_false(getForceFormulaRecalculation(wb.xlsx, "Sheet2"))
    setForceFormulaRecalculation(wb.xlsx, "*", TRUE)
    expect_true(all(getForceFormulaRecalculation(wb.xlsx, "*")))
    expect_error(setForceFormulaRecalculation(wb.xlsx, 12, TRUE), "IllegalArgumentException")
    expect_error(setForceFormulaRecalculation(
        wb.xlsx, "SheetWhichDoesNotExist",
        TRUE
    ), "IllegalArgumentException")
})
