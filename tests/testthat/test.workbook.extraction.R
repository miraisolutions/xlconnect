test_that("workbook extraction & replacement operators work for XLS worksheets", {
  wb.xls <- loadWorkbook("resources/testWorkbookExtractionOperators.xls", create = TRUE)

  # Check that writing a data set to a worksheet works (*.xls)
  wb.xls["mtcars1"] <- mtcars
  expect_true("mtcars1" %in% getSheets(wb.xls))
  expect_equal(as.vector(getLastRow(wb.xls, "mtcars1")), 33)
  wb.xls["mtcars2", startRow = 6, startCol = 11, header = FALSE] <- mtcars
  expect_true("mtcars2" %in% getSheets(wb.xls))
  expect_equal(as.vector(getLastRow(wb.xls, "mtcars2")), 37)
  # Check that reading data from a worksheet works (*.xls)
  expect_equal(dim(wb.xls["mtcars1"]), c(32, 11))
  expect_equal(dim(wb.xls["mtcars2"]), c(31, 11))
})

test_that("workbook extraction & replacement operators work for XLSX worksheets", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookExtractionOperators.xlsx", create = TRUE)

  # Check that writing a data set to a worksheet works (*.xlsx)
  wb.xlsx["mtcars1"] <- mtcars
  expect_true("mtcars1" %in% getSheets(wb.xlsx))
  expect_equal(as.vector(getLastRow(wb.xlsx, "mtcars1")), 33)
  wb.xlsx["mtcars2", startRow = 6, startCol = 11, header = FALSE] <- mtcars
  expect_true("mtcars2" %in% getSheets(wb.xlsx))
  expect_equal(as.vector(getLastRow(wb.xlsx, "mtcars2")), 37)
  # Check that reading data from a worksheet works (*.xlsx)
  expect_equal(dim(wb.xlsx["mtcars1"]), c(32, 11))
  expect_equal(dim(wb.xlsx["mtcars2"]), c(31, 11))
})

test_that("workbook extraction & replacement operators work for XLS named regions", {
  wb.xls <- loadWorkbook("resources/testWorkbookExtractionOperators.xls", create = TRUE)

  # Check that writing data to a named region works (*.xls)
  wb.xls[["mtcars3", "mtcars3!$B$7"]] <- mtcars
  expect_true("mtcars3" %in% getDefinedNames(wb.xls))
  expect_equal(as.vector(getLastRow(wb.xls, "mtcars3")), 39)
  wb.xls[["mtcars4", "mtcars4!$D$8", rownames = "Car"]] <- mtcars
  expect_true("mtcars4" %in% getDefinedNames(wb.xls))
  expect_equal(as.vector(getLastRow(wb.xls, "mtcars4")), 40)
  # Check that reading data from a named region works (*.xls)
  expect_equal(dim(wb.xls[["mtcars3"]]), c(32, 11))
  expect_equal(dim(wb.xls[["mtcars4"]]), c(32, 12))
})

test_that("workbook extraction & replacement operators work for XLSX named regions", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookExtractionOperators.xlsx", create = TRUE)

  # Check that writing data to a named region works (*.xlsx)
  wb.xlsx[["mtcars3", "mtcars3!$B$7"]] <- mtcars
  expect_true("mtcars3" %in% getDefinedNames(wb.xlsx))
  expect_equal(as.vector(getLastRow(wb.xlsx, "mtcars3")), 39)
  wb.xlsx[["mtcars4", "mtcars4!$D$8", rownames = "Car"]] <- mtcars
  expect_true("mtcars4" %in% getDefinedNames(wb.xlsx))
  expect_equal(as.vector(getLastRow(wb.xlsx, "mtcars4")), 40)
  # Check that reading data from a named region works (*.xlsx)
  expect_equal(dim(wb.xlsx[["mtcars3"]]), c(32, 11))
  expect_equal(dim(wb.xlsx[["mtcars4"]]), c(32, 12))
})
