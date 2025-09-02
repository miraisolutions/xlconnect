test_that("Test reproducing error log from issue 181 / PR 183", {
  # loading this workbook reproduces this error, but only the first time it is loaded during a session
  # it seems however that this does not prevent the workbook from being read and even when the "error"
  # is logged it isn't considered an exception and behaves more like a regular log message
  wb <- loadWorkbook("resources/testBug181.xlsx")
  obj <- readWorksheet(wb, "Sheet1")
  expect_true(is.data.frame(obj))
})
