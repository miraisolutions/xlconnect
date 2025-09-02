test_that("workbook extraction & replacement operators work for XLS worksheets", {
  wb.xls <- loadWorkbook("resources/testWorkbookExtractionOperators.xls", create = TRUE)

  # Check that writing a data set to a worksheet works (*.xls)
  wb.xls["mtcars1"] <- mtcars
  expect_true("mtcars1" %in% getSheets(wb.xls))
  expect_equal(33, as.vector(getLastRow(wb.xls, "mtcars1")))
  wb.xls["mtcars2", startRow = 6, startCol = 11, header = FALSE] <- mtcars
  expect_true("mtcars2" %in% getSheets(wb.xls))
  expect_equal(37, as.vector(getLastRow(wb.xls, "mtcars2")))
  # Check that reading data from a worksheet works (*.xls)
  expect_equal(c(32, 11), dim(wb.xls["mtcars1"]))
  expect_equal(c(31, 11), dim(wb.xls["mtcars2"]))
})

test_that("workbook extraction & replacement operators work for XLSX worksheets", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookExtractionOperators.xlsx", create = TRUE)

  # Check that writing a data set to a worksheet works (*.xlsx)
  wb.xlsx["mtcars1"] <- mtcars
  expect_true("mtcars1" %in% getSheets(wb.xlsx))
  expect_equal(33, as.vector(getLastRow(wb.xlsx, "mtcars1")))
  wb.xlsx["mtcars2", startRow = 6, startCol = 11, header = FALSE] <- mtcars
  expect_true("mtcars2" %in% getSheets(wb.xlsx))
  expect_equal(37, as.vector(getLastRow(wb.xlsx, "mtcars2")))
  # Check that reading data from a worksheet works (*.xlsx)
  expect_equal(c(32, 11), dim(wb.xlsx["mtcars1"]))
  expect_equal(c(31, 11), dim(wb.xlsx["mtcars2"]))
})

test_that("workbook extraction & replacement operators work for XLS named regions", {
  wb.xls <- loadWorkbook("resources/testWorkbookExtractionOperators.xls", create = TRUE)

  # Check that writing data to a named region works (*.xls)
  wb.xls[["mtcars3", "mtcars3!$B$7"]] <- mtcars
  expect_true("mtcars3" %in% getDefinedNames(wb.xls))
  expect_equal(39, as.vector(getLastRow(wb.xls, "mtcars3")))
  wb.xls[["mtcars4", "mtcars4!$D$8", rownames = "Car"]] <- mtcars
  expect_true("mtcars4" %in% getDefinedNames(wb.xls))
  expect_equal(40, as.vector(getLastRow(wb.xls, "mtcars4")))
  # Check that reading data from a named region works (*.xls)
  expect_equal(c(32, 11), dim(wb.xls[["mtcars3"]]))
  expect_equal(c(32, 12), dim(wb.xls[["mtcars4"]]))
})

test_that("workbook extraction & replacement operators work for XLSX named regions", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookExtractionOperators.xlsx", create = TRUE)

  # Check that writing data to a named region works (*.xlsx)
  wb.xlsx[["mtcars3", "mtcars3!$B$7"]] <- mtcars
  expect_true("mtcars3" %in% getDefinedNames(wb.xlsx))
  expect_equal(39, as.vector(getLastRow(wb.xlsx, "mtcars3")))
  wb.xlsx[["mtcars4", "mtcars4!$D$8", rownames = "Car"]] <- mtcars
  expect_true("mtcars4" %in% getDefinedNames(wb.xlsx))
  expect_equal(40, as.vector(getLastRow(wb.xlsx, "mtcars4")))
  # Check that reading data from a named region works (*.xlsx)
  expect_equal(c(32, 11), dim(wb.xlsx[["mtcars3"]]))
  expect_equal(c(32, 12), dim(wb.xlsx[["mtcars4"]]))
})
