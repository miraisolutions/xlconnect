test_that("Tests around querying Excel reference formulas", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookReferenceFormula.xls", create = FALSE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookReferenceFormula.xlsx", create = FALSE)

  # Check if reference formulas match (*.xls)
  expect_true(getReferenceFormula(wb.xls, "FirstName") == "Tabelle1!$A$1")
  expect_true(substring(getReferenceFormula(wb.xls, "SecondName"), 1, 5) == "#REF!")

  # Check if reference formulas match (*.xlsx)
  expect_true(getReferenceFormula(wb.xlsx, "FirstName") == "Tabelle1!$A$1")
  expect_true(substring(getReferenceFormula(wb.xlsx, "SecondName"), 1, 5) == "#REF!")

  # Check if reference formulas match and are in the global scope (*.xls)
  expect_formula <- "Tabelle1!$A$1"
  attributes(expect_formula) <- list(worksheetScope = "")
  expect_equal(getReferenceFormula(wb.xls, "FirstName"), expect_formula)

  expect_formula <- "#REF!"
  attributes(expect_formula) <- list(worksheetScope = "")
  expect_equal(substring(getReferenceFormula(wb.xls, "SecondName"), 1, 5), expect_formula)

  # Check if reference formulas match and are in the global scope (*.xlsx)
  expect_formula <- "Tabelle1!$A$1"
  attributes(expect_formula) <- list(worksheetScope = "")
  expect_equal(getReferenceFormula(wb.xlsx, "FirstName"), expect_formula)

  expect_formula <- "#REF!"
  attributes(expect_formula) <- list(worksheetScope = "")
  expect_equal(substring(getReferenceFormula(wb.xlsx, "SecondName"), 1, 5), expect_formula)
})
