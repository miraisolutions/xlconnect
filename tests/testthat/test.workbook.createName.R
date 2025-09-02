test_that("createName can create a legal name (*.xls)", {
  # Check that creating a legal name works ok (*.xls)
  # (this test assumes 'existsName' is working fine)
  wb.xls <- loadWorkbook(withr::local_tempfile(fileext = ".xls"), create = TRUE)
  createName(wb.xls, "Test", "Test!$C$5")
  expect_true(existsName(wb.xls, "Test"))
})

test_that("createName can create a legal name (*.xlsx)", {
  # Check that creating a legal name works ok (*.xlsx)
  # (this test assumes 'existsName' is working fine)
  wb.xlsx <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  createName(wb.xlsx, "Test", "Test!$C$5")
  expect_true(existsName(wb.xlsx, "Test"))
})

test_that("createName throws an error for illegal names (*.xls)", {
  # Check that trying to create an illegal name throws
  # an exception (*.xls)
  wb.xls <- loadWorkbook(withr::local_tempfile(fileext = ".xls"), create = TRUE)
  expect_error(createName(wb.xls, "'Test", "Test!$C$10"))
})

test_that("createName throws an error for illegal names (*.xlsx)", {
  # Check that trying to create an illegal name throws
  # an exception (*.xlsx)
  wb.xlsx <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  expect_error(createName(wb.xlsx, "'Test", "Test!$C$10"))
})

test_that("createName throws an error for illegal formulas (*.xls)", {
  # Check that trying to create a name with an illegal formula
  # throws an exception (*.xls)
  wb.xls <- loadWorkbook(withr::local_tempfile(fileext = ".xls"), create = TRUE)
  expect_error(createName(wb.xls, "IllegalFormula", "??-%&"))
})

test_that("createName throws an error for illegal formulas (*.xlsx)", {
  # Check that trying to create a name with an illegal formula
  # throws an exception (*.xlsx)
  wb.xlsx <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  expect_error(createName(wb.xlsx, "IllegalFormula", "??-%&"))
})

test_that("createName throws an error for existing names without overwrite (*.xls)", {
  # Check that trying to create an already existing name without
  # specifying 'overwrite = TRUE' throws an exception (*.xls)
  wb.xls <- loadWorkbook(withr::local_tempfile(fileext = ".xls"), create = TRUE)
  createName(wb.xls, "ImHere", "ImHere!$B$9")
  expect_error(createName(wb.xls, "ImHere", "There!$A$2"))
})

test_that("createName throws an error for existing names without overwrite (*.xlsx)", {
  # Check that trying to create an already existing name without
  # specifying 'overwrite = TRUE' throws an exception (*.xlsx)
  wb.xlsx <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  createName(wb.xlsx, "ImHere", "ImHere!$B$9")
  expect_error(createName(wb.xlsx, "ImHere", "There!$A$2"))
})

test_that("createName can overwrite an existing name (*.xls)", {
  # Check that overwriting an existing name works ok (*.xls)
  wb.xls <- loadWorkbook(withr::local_tempfile(fileext = ".xls"), create = TRUE)
  createName(wb.xls, "CurrentlyHere", "CurrentlyHere!$D$8")
  createName(wb.xls, "CurrentlyHere", "NowThere!$C$3", overwrite = TRUE)
  # TODO: Should actually rather check that new formula is correct
  expect_true(existsName(wb.xls, "CurrentlyHere"))
})

test_that("createName can overwrite an existing name (*.xlsx)", {
  # Check that overwriting an existing name works ok (*.xlsx)
  wb.xlsx <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  createName(wb.xlsx, "CurrentlyHere", "CurrentlyHere!$D$8")
  createName(wb.xlsx, "CurrentlyHere", "NowThere!$C$3", overwrite = TRUE)
  # TODO: Should actually rather check that new formula is correct
  expect_true(existsName(wb.xlsx, "CurrentlyHere"))
})

test_that("createName handles formula parsing errors correctly (*.xls)", {
  # Check that after trying to write a name with an illegal formula
  # (which throws an exception), the name remains available (*.xls)
  wb.xls <- loadWorkbook(withr::local_tempfile(fileext = ".xls"), create = TRUE)
  expect_error(createName(wb.xls, "aName", "Test!A1A4"))
  createName(wb.xls, "aName", "Test!A1")
  expect_true(existsName(wb.xls, "aName"))
})

test_that("createName handles formula parsing errors correctly (*.xlsx)", {
  # Check that after trying to write a name with an illegal formula
  # (which throws an exception), the name remains available (*.xlsx)
  #
  # NOTE: This seems to have changed with POI 3.11-beta1 - creating a
  # name with an invalid formula does not throw an exception anymore!
  wb.xlsx <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  # checkException(createName(wb.xlsx, "aName", "Test!A1A4"))
  # checkNoException(createName(wb.xlsx, "aName", "Test!A1"))
  # checkEquals(existsName(wb.xlsx, "aName"), TRUE, check.attributes = FALSE)
})

test_that("createName can create a worksheet-scoped name (*.xls)", {
  # Check that creating a worksheet-scoped name works (*.xls)
  # we first need to create a worksheet:
  # global names can be created without an existing sheet, but not scoped ones.
  wb.xls <- loadWorkbook(withr::local_tempfile(fileext = ".xls"), create = TRUE)
  sheetName <- "Test_Scoped_Sheet"

  createSheet(wb.xls, name = sheetName)
  expect_true(existsSheet(wb.xls, sheetName))

  expect_found = TRUE
  attr(expect_found, "worksheetScope") <- sheetName
  createName(wb.xls, "Test_1", "Test!$C$5", worksheetScope = sheetName)
  expect_equal(existsName(wb.xls, "Test_1"), expect_found)
})

test_that("createName can create a worksheet-scoped name (*.xlsx)", {
  # Check that creating a worksheet-scoped name works (*.xlsx)
  # (this test assumes 'existsName' is working fine)
  wb.xlsx <- loadWorkbook(withr::local_tempfile(fileext = ".xlsx"), create = TRUE)
  sheetName <- "Test_Scoped_Sheet"

  createSheet(wb.xlsx, name = sheetName)
  expect_true(existsSheet(wb.xlsx, sheetName))

  expect_found = TRUE
  attr(expect_found, "worksheetScope") <- sheetName
  createName(wb.xlsx, "Test_1", "Test!$C$5", worksheetScope = sheetName)
  expect_equal(existsName(wb.xlsx, "Test_1"), expect_found)
})
