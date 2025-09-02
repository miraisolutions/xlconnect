test_that("writeWorksheet throws error for non-data.frame objects", {
  # Create workbooks
  wb.xls <- loadWorkbook(withr::local_tempfile(fileext = ".xls"), create = TRUE)
  wb.xlsx <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)

  # Check that trying to write an object which cannot be converted to a data.frame
  # causes an exception (*.xls)
  createSheet(wb.xls, "test1")
  expect_error(writeWorksheet(wb.xls, search, "test1"))

  # Check that trying to write an object which cannot be converted to a data.frame
  # causes an exception (*.xlsx)
  createSheet(wb.xlsx, "test1")
  expect_error(writeWorksheet(wb.xlsx, search, "test1"))
})

test_that("writeWorksheet throws error for non-existent sheets", {
  # Create workbooks
  wb.xls <- loadWorkbook(withr::local_tempfile(fileext = ".xls"), create = TRUE)
  wb.xlsx <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)

  # Check that attempting to write to a non-existing sheet causes an exception (*.xls)
  expect_error(writeWorksheet(wb.xls, mtcars, "sheetDoesNotExist"))

  # Check that attempting to write to a non-existing sheet causes an exception (*.xlsx)
  expect_error(writeWorksheet(wb.xlsx, mtcars, "sheetDoesNotExist"))
})

test_that("writeWorksheet can preserve formulas", {
  # Load workbooks with formulas on them
  wb.xls <- loadWorkbook(rsrc("testWorkbookOverwriteFormulas.xls"))
  wb.xlsx <- loadWorkbook(rsrc("testWorkbookOverwriteFormulas.xlsx"))

  # Initialize mtcars_mod as is defined with the formula in the carb column (set equal to the gear column)
  mtcars_mod <- mtcars
  mtcars_mod$carb <- mtcars_mod$gear

  test_overwrite_formula <- function(wb, expected, overwrite = TRUE) {
    writeWorksheet(wb, mtcars, "mtcars_formula", startRow = 9, startCol = 5, overwriteFormulaCells = overwrite)
    res <- readWorksheet(wb, "mtcars_formula")
    expect_equal(res, normalizeDataframe(expected), ignore_attr = c("worksheetScope", "row.names"))
  }

  # Check that formulas can be kept in existing named region (*.xls)
  test_overwrite_formula(wb.xls, mtcars_mod, overwrite = FALSE)

  # Check that formulas can be kept in existing named region (*.xlsx)
  test_overwrite_formula(wb.xlsx, mtcars_mod, overwrite = FALSE)
})

test_that("writeWorksheet can overwrite formulas", {
  # Load workbooks with formulas on them
  wb.xls <- loadWorkbook(rsrc("testWorkbookOverwriteFormulas.xls"))
  wb.xlsx <- loadWorkbook(rsrc("testWorkbookOverwriteFormulas.xlsx"))

  test_overwrite_formula <- function(wb, expected, overwrite = TRUE) {
    writeWorksheet(wb, mtcars, "mtcars_formula", startRow = 9, startCol = 5, overwriteFormulaCells = overwrite)
    res <- readWorksheet(wb, "mtcars_formula")
    expect_equal(res, normalizeDataframe(expected), ignore_attr = c("worksheetScope", "row.names"))
  }

  # Check that formulas are overwritten (*.xls)
  test_overwrite_formula(wb.xls, mtcars)

  # Check that formulas are overwritten (*.xlsx)
  test_overwrite_formula(wb.xlsx, mtcars)
})
