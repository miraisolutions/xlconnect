test_that("saveWorkbook saves an XLS file", {
  skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
  file.xls <- withr::local_tempfile(fileext = ".xls")

  wb.xls <- loadWorkbook(file.xls, create = TRUE)
  expect_false(file.exists(file.xls))
  saveWorkbook(wb.xls)
  expect_true(file.exists(file.xls))
})

test_that("saveWorkbook saves an XLSX file", {
  skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
  file.xlsx <- withr::local_tempfile(fileext = ".xlsx")

  wb.xlsx <- loadWorkbook(file.xlsx, create = TRUE)
  expect_false(file.exists(file.xlsx))
  saveWorkbook(wb.xlsx)
  expect_true(file.exists(file.xlsx))
})

test_that("saveWorkbook saves an XLS file to a new location", {
  skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
  file.xls <- withr::local_tempfile(fileext = ".xls")
  newFile.xls <- withr::local_tempfile(fileext = ".xls")

  wb.xls <- loadWorkbook(file.xls, create = TRUE)
  saveWorkbook(wb.xls, file = newFile.xls)
  expect_true(file.exists(newFile.xls))
})

test_that("saveWorkbook saves an XLSX file to a new location", {
  skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
  file.xlsx <- withr::local_tempfile(fileext = ".xlsx")
  newFile.xlsx <- withr::local_tempfile(fileext = ".xlsx")

  wb.xlsx <- loadWorkbook(file.xlsx, create = TRUE)
  saveWorkbook(wb.xlsx, file = newFile.xlsx)
  expect_true(file.exists(newFile.xlsx))
})
