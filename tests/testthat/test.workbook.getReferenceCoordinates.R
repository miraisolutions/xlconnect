test_that("getReferenceCoordinatesForName returns correct coordinates for a valid name in XLS", {
  wb.xls <- loadWorkbook("resources/testWorkbookReferenceFormula.xls", create = FALSE)
  expect_true(all(getReferenceCoordinatesForName(wb.xls, "FirstName") == matrix(c(1, 1, 1, 1), nrow = 2, byrow = TRUE)))
})

test_that("getReferenceCoordinatesForName throws an error for a non-existent name in XLS", {
  wb.xls <- loadWorkbook("resources/testWorkbookReferenceFormula.xls", create = FALSE)
  expect_error(getReferenceCoordinatesForName(wb.xls, "NonExistentName"))
})

test_that("getReferenceCoordinatesForName returns correct coordinates for a valid name in XLSX", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookReferenceFormula.xlsx", create = FALSE)
  expect_true(all(
    getReferenceCoordinatesForName(wb.xlsx, "FirstName") == matrix(c(1, 1, 1, 1), nrow = 2, byrow = TRUE)
  ))
})

test_that("getReferenceCoordinatesForName throws an error for a non-existent name in XLSX", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookReferenceFormula.xlsx", create = FALSE)
  expect_error(getReferenceCoordinatesForName(wb.xlsx, "NonExistentName"))
})
