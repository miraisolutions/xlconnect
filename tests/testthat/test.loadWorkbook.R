test_that("loading non-existent files throws an error", {
  expect_error(loadWorkbook(rsrc("fileWhichDoesNotExist.xls")))
  expect_error(loadWorkbook(rsrc("fileWhichDoesNotExist.xlsx")))
})

test_that("loading existing valid XLS and XLSX files works", {
  wb_xls <- loadWorkbook(rsrc("testLoadWorkbook.xls"))
  expect_true(is(wb_xls, "workbook"))

  wb_xlsx <- loadWorkbook(rsrc("testLoadWorkbook.xlsx"))
  expect_true(is(wb_xlsx, "workbook"))
})

test_that("creating a new XLS file on the fly works", {
  file_to_create_xls <- withr::local_tempfile(fileext = ".xls")
  wb_create_xls <- loadWorkbook(file_to_create_xls, create = TRUE)
  expect_true(is(wb_create_xls, "workbook"))
  saveWorkbook(wb_create_xls, file_to_create_xls)
  expect_true(file.exists(file_to_create_xls))
})

test_that("creating a new XLSX file on the fly works", {
  file_to_create_xlsx <- withr::local_tempfile(fileext = ".xlsx")
  wb_create_xlsx <- loadWorkbook(file_to_create_xlsx, create = TRUE)
  expect_true(is(wb_create_xlsx, "workbook"))
  saveWorkbook(wb_create_xlsx, file_to_create_xlsx)
  expect_true(file.exists(file_to_create_xlsx))
})

test_that("loading a password-protected file (testBug61.xlsx) works", {
  pwdProtectedFile1 <- rsrc("testBug61.xlsx")
  expect_error(
    loadWorkbook(pwdProtectedFile1),
    "EncryptedDocumentException (Java): The supplied spreadsheet is protected, but no password was supplied",
    fixed = TRUE
  )
  expect_error(
    loadWorkbook(pwdProtectedFile1, password = "wrong"),
    "EncryptedDocumentException (Java): Password incorrect",
    fixed = TRUE
  )
  wb_pwd1 <- loadWorkbook(pwdProtectedFile1, password = "mirai")
  expect_true(is(wb_pwd1, "workbook"))
})

test_that("loading a password-protected file (testBug106.xlsx) works", {
  # Excel2019
  pwdProtectedFile2 <- rsrc("testBug106.xlsx")
  expect_error(
    loadWorkbook(pwdProtectedFile2),
    "EncryptedDocumentException (Java): The supplied spreadsheet is protected, but no password was supplied",
    fixed = TRUE
  )
  expect_error(
    loadWorkbook(pwdProtectedFile2, password = "wrong"),
    "EncryptedDocumentException (Java): Password incorrect",
    fixed = TRUE
  )
  wb_pwd2 <- loadWorkbook(pwdProtectedFile2, password = "mirai")
  expect_true(is(wb_pwd2, "workbook"))
})

test_that("loading a file relative to the working directory works (#212)", {
  setwd(rsrc(""))
  wb_direct_name <- loadWorkbook("testLoadWorkbook.xls") # avoid using rsrc here, as it creates an absolute path
  expect_true(is(wb_direct_name, "workbook"))
})
