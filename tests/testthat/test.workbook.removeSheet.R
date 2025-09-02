test_that("removeSheet removes an existing sheet in XLS - assumes 'existsSheet' to be working properly", {
  wb.xls <- loadWorkbook("resources/testWorkbookRemoveSheet.xls", create = FALSE)
  # Check that removing a sheet works fine (*.xls)
  removeSheet(wb.xls, "BBB")
  expect_false(existsSheet(wb.xls, "BBB"))
})

test_that("removeSheet removes an existing sheet in XLSX - assumes 'existsSheet' to be working properly", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookRemoveSheet.xlsx", create = FALSE)
  # Check that removing a sheet works fine (*.xlsx)
  removeSheet(wb.xlsx, "BBB")
  expect_false(existsSheet(wb.xlsx, "BBB"))
})

test_that("removeSheet does not throw an exception when removing a non-existing sheet in XLS", {
  wb.xls <- loadWorkbook("resources/testWorkbookRemoveSheet.xls", create = FALSE)
  # Check that removing a non-existing sheet does not throw an exception (*.xls)
  expect_error(removeSheet(wb.xls, 35), NA)
  expect_error(removeSheet(wb.xls, "SheetWhichDoesNotExist"), NA)
})

test_that("removeSheet does not throw an exception when removing a non-existing sheet in XLSX", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookRemoveSheet.xlsx", create = FALSE)
  # Check that removing a non-existing sheet does not throw an exception (*.xlsx)
  expect_error(removeSheet(wb.xlsx, 35), NA)
  expect_error(removeSheet(wb.xlsx, "SheetWhichDoesNotExist"), NA)
})
