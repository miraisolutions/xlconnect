test_that("createName can create a legal name", {
  # (this test assumes 'existsName' is working fine)
  wb <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  createName(wb, "Test", "Test!$C$5")
  res_xlsx_global <- existsName(wb, "Test")
  expect_true(res_xlsx_global)
  xlsx_global_scope <- attr(res_xlsx_global, "worksheetScope")
  expect_true(is.null(xlsx_global_scope) || xlsx_global_scope == "")
})

test_that("createName throws an error for illegal names", {
  wb <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  expect_error(createName(wb, "'Test", "Test!$C$10"), "IllegalArgumentException")
})

test_that("createName throws an error for illegal formulas", {
  wb <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  expect_error(createName(wb, "IllegalFormula", "??-%&"), "IllegalArgumentException")
})

test_that("createName throws an error for existing names without overwrite", {
  wb <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  createName(wb, "ImHere", "ImHere!$B$9")
  expect_error(createName(wb, "ImHere", "There!$A$2"), "IllegalArgumentException")
})

test_that("createName can overwrite an existing name", {
  wb <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  createName(wb, "CurrentlyHere", "CurrentlyHere!$D$8")
  createName(wb, "CurrentlyHere", "NowThere!$C$3", overwrite = TRUE)
  # TODO: Should actually rather check that new formula is correct
  expect_true(existsName(wb, "CurrentlyHere"))
})

test_that("createName handles formula parsing errors correctly", {
  wb.xls <- loadWorkbook(withr::local_tempfile(fileext = ".xls"), create = TRUE)
  # This call should not produce an error:
  expect_error(createName(wb.xls, "aName", "Test!A1A4"), "IllegalArgumentException")
  createName(wb.xls, "aName", "Test!A1")
  expect_true(existsName(wb.xls, "aName"))

  # NOTE: This seems to have changed with POI 3.11-beta1 - creating a
  # name with an invalid formula does not throw an exception anymore!
})

test_that("createName can create a worksheet-scoped name", {
  # we first need to create a worksheet:
  # global names can be created without an existing sheet, but not scoped ones.
  wb <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  sheetName <- "Test_Scoped_Sheet"
  createSheet(wb, name = sheetName)
  expect_true(existsSheet(wb, sheetName))
  createName(wb, "Test_1", paste0(sheetName, "!$C$5"), worksheetScope = sheetName)
  res_xlsx <- existsName(wb, "Test_1")
  expect_true(res_xlsx)

  expect_equal(attr(res_xlsx, "worksheetScope"), sheetName)
})
