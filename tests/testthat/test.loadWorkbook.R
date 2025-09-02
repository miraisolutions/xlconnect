test_that("loading non-existent files throws an error", {
  # Check that an exception is thrown when trying to open
  # a non-existent file (*.xls)
  expect_error(loadWorkbook(rsrc("fileWhichDoesNotExist.xls")))

  # Check that an exception is thrown when trying to open
  # a non-existent file (*.xlsx)
  expect_error(loadWorkbook(rsrc("fileWhichDoesNotExist.xlsx")))
})

test_that("loading existing valid XLS and XLSX files works", {
  # Check that an instance of the workbook class is returned
  # for an already existing file (*.xls)
  wb_xls <- loadWorkbook(rsrc("testLoadWorkbook.xls"))
  expect_true(is(wb_xls, "workbook"))

  # Check that an instance of the workbook class is returned
  # for an already existing file (*.xlsx)
  wb_xlsx <- loadWorkbook(rsrc("testLoadWorkbook.xlsx"))
  expect_true(is(wb_xlsx, "workbook"))
})

test_that("creating a new XLS file on the fly works", {
  # Check that an instance of the workbook class is returned
  # for a file created on-the-fly (*.xls)
  file_to_create_xls <- withr::local_tempfile(fileext = ".xls")
  wb_create_xls <- loadWorkbook(file_to_create_xls, create = TRUE)
  expect_true(is(wb_create_xls, "workbook"))
  saveWorkbook(wb_create_xls, file_to_create_xls)
  expect_true(file.exists(file_to_create_xls))
})

test_that("creating a new XLSX file on the fly works", {
  # Check that an instance of the workbook class is returned
  # for a file created on-the-fly (*.xlsx)
  file_to_create_xlsx <- withr::local_tempfile(fileext = ".xlsx")
  wb_create_xlsx <- loadWorkbook(file_to_create_xlsx, create = TRUE)
  expect_true(is(wb_create_xlsx, "workbook"))
  saveWorkbook(wb_create_xlsx, file_to_create_xlsx)
  expect_true(file.exists(file_to_create_xlsx))
})

test_that("loading a password-protected file (testBug61.xlsx) works", {
  pwdProtectedFile1 <- rsrc("testBug61.xlsx")

  # Check that opening a password protected file throws an error
  # if no password is specified
  expect_error(
    loadWorkbook(pwdProtectedFile1),
    "EncryptedDocumentException (Java): The supplied spreadsheet is protected, but no password was supplied",
    fixed = TRUE
  )

  # Check that opening a password protected file throws an error
  # if a wrong password is specified
  expect_error(
    loadWorkbook(pwdProtectedFile1, password = "wrong"),
    "EncryptedDocumentException (Java): Password incorrect",
    fixed = TRUE
  )

  # Check that a password protected file can be opened if the
  # correct password is specified
  wb_pwd1 <- loadWorkbook(pwdProtectedFile1, password = "mirai")
  expect_true(is(wb_pwd1, "workbook"))
})

test_that("loading a password-protected file (testBug106.xlsx) works", {
  pwdProtectedFile2 <- rsrc("testBug106.xlsx")

  # Check that opening a password protected file throws an error
  # if no password is specified
  # Excel2019
  expect_error(
    loadWorkbook(pwdProtectedFile2),
    "EncryptedDocumentException (Java): The supplied spreadsheet is protected, but no password was supplied",
    fixed = TRUE
  )

  # Check that opening a password protected file throws an error
  # if a wrong password is specified
  # Excel2019
  expect_error(
    loadWorkbook(pwdProtectedFile2, password = "wrong"),
    "EncryptedDocumentException (Java): Password incorrect",
    fixed = TRUE
  )

  # Check that a password protected file can be opened if the
  # correct password is specified
  # Excel2019
  wb_pwd2 <- loadWorkbook(pwdProtectedFile2, password = "mirai")
  expect_true(is(wb_pwd2, "workbook"))
})

test_that("loading a file relative to the working directory works (#212)", {
  # Check that loadWorkbook can handle relative paths (#212)
  setwd(rsrc(""))
  wb_direct_name <- loadWorkbook("testLoadWorkbook.xls") # avoid using rsrc here, as it creates an absolute path
  expect_true(is(wb_direct_name, "workbook"))
})
