context("loadWorkbook Functionality")

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

test_that("creating new XLS and XLSX files on the fly works", {
    # Test creating an XLS file
    file_to_create_xls <- rsrc("fileCreatedOnTheFly.xls")
    on.exit(if (file.exists(file_to_create_xls)) file.remove(file_to_create_xls), add = TRUE) # Ensure cleanup

    wb_create_xls <- loadWorkbook(file_to_create_xls, create = TRUE)
    expect_true(is(wb_create_xls, "workbook"))
    expect_true(file.exists(file_to_create_xls))

    # Test creating an XLSX file
    file_to_create_xlsx <- rsrc("fileCreatedOnTheFly.xlsx")
    on.exit(if (file.exists(file_to_create_xlsx)) file.remove(file_to_create_xlsx), add = TRUE) # Ensure cleanup

    wb_create_xlsx <- loadWorkbook(file_to_create_xlsx, create = TRUE)
    expect_true(is(wb_create_xlsx, "workbook"))
    expect_true(file.exists(file_to_create_xlsx))
})

test_that("loading password-protected files works correctly", {
    # Test case 1: testBug61.xlsx
    pwdProtectedFile1 <- rsrc("testBug61.xlsx")
    expect_error(loadWorkbook(pwdProtectedFile1), "Workbook is password-protected")
    expect_error(loadWorkbook(pwdProtectedFile1, password = "wrong"), "Incorrect password")
    wb_pwd1 <- loadWorkbook(pwdProtectedFile1, password = "mirai")
    expect_true(is(wb_pwd1, "workbook"))

    # Test case 2: testBug106.xlsx
    pwdProtectedFile2 <- rsrc("testBug106.xlsx")
    expect_error(loadWorkbook(pwdProtectedFile2), "Workbook is password-protected")
    expect_error(loadWorkbook(pwdProtectedFile2, password = "wrong"), "Incorrect password")
    wb_pwd2 <- loadWorkbook(pwdProtectedFile2, password = "mirai")
    expect_true(is(wb_pwd2, "workbook"))
})

test_that("loading a file using rsrc (previously relied on setwd) works", {
    wb_direct_name <- loadWorkbook(rsrc("testLoadWorkbook.xls"))
    expect_true(is(wb_direct_name, "workbook"))
})
