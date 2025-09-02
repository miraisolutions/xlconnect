test_that("when removing a name from a worksheet it does not exist anymore (*.xls) - assumes 'existsName' is working correctly", {
  wb.xls <- loadWorkbook("resources/testWorkbookRemoveName.xls", create = FALSE)
  removeName(wb.xls, "AA")
  expect_false(existsName(wb.xls, "AA"))
})

test_that("when removing a name from a worksheet it does not exist anymore (*.xlsx) - assumes 'existsName' is working correctly", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookRemoveName.xlsx", create = FALSE)
  removeName(wb.xlsx, "AA")
  expect_false(existsName(wb.xlsx, "AA"))
})

test_that("attempting to remove a non-existing name does not throw an exception (*.xls)", {
  wb.xls <- loadWorkbook("resources/testWorkbookRemoveName.xls", create = FALSE)
  expect_error(removeName(wb.xls, "NameWhichDoesNotExist"), NA)
})

test_that("attempting to remove a non-existing name does not throw an exception (*.xlsx)", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookRemoveName.xlsx", create = FALSE)
  expect_error(removeName(wb.xlsx, "NameWhichDoesNotExist"), NA)
})
