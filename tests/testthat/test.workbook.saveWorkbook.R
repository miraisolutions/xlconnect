test_that("saveWorkbook saves an XLS file", {
  skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
  # Create workbooks
  file.xls <- withr::local_tempfile(fileext = ".xls")

  wb.xls <- loadWorkbook(file.xls, create = TRUE)

  # Files don't exist yet
  expect_false(file.exists(file.xls))

  saveWorkbook(wb.xls)

  # Check that file exists after saving (*.xls)
  expect_true(file.exists(file.xls))
})

test_that("saveWorkbook saves an XLSX file", {
  skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
  # Create workbooks
  file.xlsx <- withr::local_tempfile(fileext = ".xlsx")

  wb.xlsx <- loadWorkbook(file.xlsx, create = TRUE)

  # Files don't exist yet
  expect_false(file.exists(file.xlsx))

  saveWorkbook(wb.xlsx)

  # Check that file exists after saving (*.xlsx)
  expect_true(file.exists(file.xlsx))
})

test_that("saveWorkbook saves an XLS file to a new location", {
  skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
  file.xls <- withr::local_tempfile(fileext = ".xls")
  newFile.xls <- withr::local_tempfile(fileext = ".xls")

  wb.xls <- loadWorkbook(file.xls, create = TRUE)

  # Check save as (*.xls)
  saveWorkbook(wb.xls, file = newFile.xls)
  expect_true(file.exists(newFile.xls))
})

test_that("saveWorkbook saves an XLSX file to a new location", {
  skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
  file.xlsx <- withr::local_tempfile(fileext = ".xlsx")
  newFile.xlsx <- withr::local_tempfile(fileext = ".xlsx")

  wb.xlsx <- loadWorkbook(file.xlsx, create = TRUE)

  # Check save as (*.xlsx)
  saveWorkbook(wb.xlsx, file = newFile.xlsx)
  expect_true(file.exists(newFile.xlsx))
})
