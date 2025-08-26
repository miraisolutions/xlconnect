test_that("writeWorksheet throws error for non-data.frame objects", {
  wb <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  createSheet(wb, "test1")
  expect_error(writeWorksheet(wb, search, "test1"))
})

test_that("writeWorksheet throws error for non-existent sheets", {
  wb <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  expect_error(writeWorksheet(wb, mtcars, "sheetDoesNotExist"))
})

test_that("writeWorksheet can preserve formulas", {
  wb.xls <- loadWorkbook(rsrc("testWorkbookOverwriteFormulas.xls"))
  wb.xlsx <- loadWorkbook(rsrc("testWorkbookOverwriteFormulas.xlsx"))

  mtcars_mod <- mtcars
  mtcars_mod$carb <- mtcars_mod$gear

  # Test for .xls
  writeWorksheet(wb.xls, mtcars, "mtcars_formula", startRow = 9, startCol = 5, overwriteFormulaCells = FALSE)
  res.xls <- readWorksheet(wb.xls, "mtcars_formula")
  expect_equal(res.xls, normalizeDataframe(mtcars_mod), ignore_attr = c("worksheetScope", "row.names"))

  # Test for .xlsx
  writeWorksheet(wb.xlsx, mtcars, "mtcars_formula", startRow = 9, startCol = 5, overwriteFormulaCells = FALSE)
  res.xlsx <- readWorksheet(wb.xlsx, "mtcars_formula")
  expect_equal(res.xlsx, normalizeDataframe(mtcars_mod), ignore_attr = c("worksheetScope", "row.names"))
})

test_that("writeWorksheet can overwrite formulas", {
  wb.xls <- loadWorkbook(rsrc("testWorkbookOverwriteFormulas.xls"))
  wb.xlsx <- loadWorkbook(rsrc("testWorkbookOverwriteFormulas.xlsx"))

  # Test for .xls
  writeWorksheet(wb.xls, mtcars, "mtcars_formula", startRow = 9, startCol = 5, overwriteFormulaCells = TRUE)
  res.xls <- readWorksheet(wb.xls, "mtcars_formula")
  expect_equal(res.xls, normalizeDataframe(mtcars), ignore_attr = c("worksheetScope", "row.names"))

  # Test for .xlsx
  writeWorksheet(wb.xlsx, mtcars, "mtcars_formula", startRow = 9, startCol = 5, overwriteFormulaCells = TRUE)
  res.xlsx <- readWorksheet(wb.xlsx, "mtcars_formula")
  expect_equal(res.xlsx, normalizeDataframe(mtcars), ignore_attr = c("worksheetScope", "row.names"))
})
