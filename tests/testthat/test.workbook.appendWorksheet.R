test_that("test.workbook.appendWorksheet", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookAppend.xls")
  wb.xlsx <- loadWorkbook("resources/testWorkbookAppend.xlsx")

  # Check that appending data to a worksheet produces the expected result (*.xls)
  appendWorksheet(wb.xls, mtcars, sheet = "mtcars")
  res <- readWorksheet(wb.xls, sheet = "mtcars")
  rownames(res) <- as.character(seq_len(nrow(res)))
  expect_equal(getLastRow(wb.xls, "mtcars"), c(mtcars = 73))
  expected_data_xls <- rbind(mtcars, mtcars)
  rownames(expected_data_xls) <- as.character(seq_len(nrow(expected_data_xls)))
  expect_equal(normalizeDataframe(res), normalizeDataframe(expected_data_xls))

  # Check that appending data to a named region produces the expected result (*.xlsx)
  appendWorksheet(wb.xlsx, mtcars, sheet = "mtcars")
  res <- readWorksheet(wb.xlsx, sheet = "mtcars")
  rownames(res) <- as.character(seq_len(nrow(res)))
  expect_equal(getLastRow(wb.xlsx, "mtcars"), c(mtcars = 73))
  expected_data_xlsx <- rbind(mtcars, mtcars)
  rownames(expected_data_xlsx) <- as.character(seq_len(nrow(expected_data_xlsx)))
  expect_equal(normalizeDataframe(res), normalizeDataframe(expected_data_xlsx))

  # Check that trying to append to an non-existing worksheet throws an error (*.xls)
  expect_error(appendWorksheet(wb.xls, mtcars, sheet = "doesNotExist"), "IllegalArgumentException")

  # Check that trying to append to an non-existing worksheet throws an error (*.xlsx)
  expect_error(appendWorksheet(wb.xlsx, mtcars, sheet = "doesNotExist"), "IllegalArgumentException")
})
