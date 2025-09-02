test_that("test.workbook.getReferenceCoordinatesForTable", {
  # Create workbooks
  wb.xlsx <- loadWorkbook("resources/testWorkbookReadTable.xlsx", create = FALSE)

  # Check that querying the reference coordinates of an Excel table works as expected (*.xlsx)
  res <- getReferenceCoordinatesForTable(wb.xlsx, sheet = "Test", table = "TestTable1")
  expect_equal(matrix(c(5, 10, 4, 7), ncol = 2), res)
})
