test_that("test.workbook.testBug181", {
  wb <- loadWorkbook("resources/testBug181.xlsx")
  obj <- readWorksheet(wb, "Sheet1")
  expect_true(is.data.frame(obj))
})
