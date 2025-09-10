test_that("writeNamedRegion throws error for non-data.frame objects in XLS", {
  # Check that trying to write an object which cannot be converted to a data.frame
  # causes an exception (*.xls)
  wb.xls <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xls", create = TRUE)
  createName(wb.xls, "test1", "Test1!$C$8")
  expect_error(writeNamedRegion(wb.xls, search, "test1"), "cannot coerce class")
})

test_that("writeNamedRegion throws error for non-data.frame objects in XLSX", {
  # Check that trying to write an object which cannot be converted to a data.frame
  # causes an exception (*.xlsx)
  wb.xlsx <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xlsx", create = TRUE)
  createName(wb.xlsx, "test1", "Test1!$C$8")
  expect_error(writeNamedRegion(wb.xlsx, search, "test1"), "cannot coerce class")
})

test_that("writeNamedRegion throws error for non-existent names in XLS", {
  # Check that attempting to write to a non-existing name causes an exception (*.xls)
  wb.xls <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xls", create = TRUE)
  expect_error(writeNamedRegion(wb.xls, mtcars, "nameDoesNotExist"), "IllegalArgumentException")
})

test_that("writeNamedRegion throws error for non-existent names in XLSX", {
  # Check that attempting to write to a non-existing name causes an exception (*.xlsx)
  wb.xlsx <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xlsx", create = TRUE)
  expect_error(writeNamedRegion(wb.xlsx, mtcars, "nameDoesNotExist"), "IllegalArgumentException")
})

test_that("writeNamedRegion throws error for non-existent sheets in XLS", {
  # Check that attempting to write to a name which referes to a non-existing sheet
  # causes an exception (*.xls)
  wb.xls <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xls", create = TRUE)
  createName(wb.xls, "nope", "NonExistingSheet!A1")
  expect_error(writeNamedRegion(wb.xls, mtcars, "nope"), "IllegalArgumentException")
})

test_that("writeNamedRegion throws error for non-existent sheets in XLSX", {
  # Check that attempting to write to a name which referes to a non-existing sheet
  # causes an exception (*.xlsx)
  wb.xlsx <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xlsx", create = TRUE)
  createName(wb.xlsx, "nope", "NonExistingSheet!A1")
  expect_error(writeNamedRegion(wb.xlsx, mtcars, "nope"), "IllegalArgumentException")
})

test_that("writeNamedRegion handles empty data frames in XLS", {
  # Check that writing an empty data.frame does not cause an error (*.xls)
  wb.xls <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xls", create = TRUE)
  createSheet(wb.xls, "empty")
  createName(wb.xls, "empty1", "empty!A1")
  createName(wb.xls, "empty2", "empty!D10")
  writeNamedRegion(wb.xls, data.frame(), "empty1")
  writeNamedRegion(wb.xls, data.frame(a = character(0), b = numeric(0)), "empty2")
  expect_true(TRUE)
})

test_that("writeNamedRegion handles empty data frames in XLSX", {
  # Check that writing an empty data.frame does not cause an error (*.xlsx)
  wb.xlsx <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xlsx", create = TRUE)
  createSheet(wb.xlsx, "empty")
  createName(wb.xlsx, "empty1", "empty!A1")
  createName(wb.xlsx, "empty2", "empty!D10")
  writeNamedRegion(wb.xlsx, data.frame(), "empty1")
  writeNamedRegion(wb.xlsx, data.frame(a = character(0), b = numeric(0)), "empty2")
  expect_true(TRUE)
})

# Helper function for testing formula overwrite behavior
test_overwrite_formula <- function(wb, expected, overwrite = TRUE) {
  writeNamedRegion(wb, mtcars, "mtcars_formula", overwriteFormulaCells = overwrite)
  res <- readNamedRegion(wb, "mtcars_formula")
  expect_equal(res, normalizeDataframe(expected), ignore_attr = c("worksheetScope", "row.names"))
}

test_that("writeNamedRegion with overwriteFormulaCells = FALSE works in XLS", {
  # Load workbooks with formulas on them
  wb.xls <- loadWorkbook("resources/testWorkbookOverwriteFormulas.xls")
  # Initialize mtcars_mod as is defined with the formula in the carb column (set equal to the gear column)
  mtcars_mod <- mtcars
  mtcars_mod$carb <- mtcars_mod$gear
  # Check that formulas can be kept in existing named region (*.xls)
  test_overwrite_formula(wb.xls, mtcars_mod, overwrite = FALSE)
})

test_that("writeNamedRegion with overwriteFormulaCells = FALSE works in XLSX", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookOverwriteFormulas.xlsx")
  mtcars_mod <- mtcars
  mtcars_mod$carb <- mtcars_mod$gear
  # Check that formulas can be kept in existing named region (*.xlsx)
  test_overwrite_formula(wb.xlsx, mtcars_mod, overwrite = FALSE)
})

test_that("writeNamedRegion with overwriteFormulaCells = TRUE works in XLS", {
  wb.xls <- loadWorkbook("resources/testWorkbookOverwriteFormulas.xls")
  # Check that formulas are overwritten (*.xls)
  test_overwrite_formula(wb.xls, mtcars)
})

test_that("writeNamedRegion with overwriteFormulaCells = TRUE works in XLSX", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookOverwriteFormulas.xlsx")
  # Check that formulas are overwritten (*.xlsx)
  test_overwrite_formula(wb.xlsx, mtcars)
})
