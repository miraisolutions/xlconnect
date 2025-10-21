test_that("workbook.setForceFormulaRecalculation - setting and getting force formula recalculation flags", {
  # Create workbook
  wb.xlsx <- loadWorkbook("resources/testBug170.xlsx", create = FALSE)

  # Check that setting the force formula recalculation flag works fine (*.xlsx)
  # (assumes that 'getForceFormulaRecalculation' works fine)
  setForceFormulaRecalculation(wb.xlsx, 1, TRUE)
  expect_true(getForceFormulaRecalculation(wb.xlsx, 1))
  setForceFormulaRecalculation(wb.xlsx, c("Sheet1", "Sheet2"), FALSE)
  expect_false(getForceFormulaRecalculation(wb.xlsx, "Sheet2"))

  # Check that passing multiple sheets doesn't cause problems
  setForceFormulaRecalculation(wb.xlsx, "*", TRUE)
  expect_true(all(getForceFormulaRecalculation(wb.xlsx, "*")))

  # Check that setting the force formula recalculation flag an illegal active sheet throws an exception (*.xlsx)
  expect_error(setForceFormulaRecalculation(wb.xlsx, 12, TRUE), "IllegalArgumentException")
  expect_error(setForceFormulaRecalculation(wb.xlsx, "SheetWhichDoesNotExist", TRUE), "IllegalArgumentException")
})
