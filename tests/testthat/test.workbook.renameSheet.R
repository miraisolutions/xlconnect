test_that("test.workbook.renameSheet", {
  # Create workbooks
  wb.xls <- loadWorkbook("resources/testWorkbookRenameSheet.xls", create = TRUE)
  wb.xlsx <- loadWorkbook("resources/testWorkbookRenameSheet.xlsx", create = TRUE)

  # Check that renaming a sheet (using its name) works fine (*.xls)
  # (assumes 'createSheet' and 'existsSheet' to be working properly)
  createSheet(wb.xls, name = "OldName1")
  renameSheet(wb.xls, sheet = "OldName1", newName = "NewName1")
  expect_true(existsSheet(wb.xls, "NewName1"))
  expect_false(existsSheet(wb.xls, "OldName1"))

  # Check that renaming a sheet (using its name) works fine (*.xlsx)
  # (assumes 'createSheet' and 'existsSheet' to be working properly)
  createSheet(wb.xlsx, name = "OldName1")
  renameSheet(wb.xlsx, sheet = "OldName1", newName = "NewName1")
  expect_true(existsSheet(wb.xlsx, "NewName1"))
  expect_false(existsSheet(wb.xlsx, "OldName1"))

  # Check that renaming a sheet (using its index) works fine (*.xls)
  # (assumes 'createSheet' and 'existsSheet' to be working properly)
  createSheet(wb.xls, name = "OldName2")
  renameSheet(wb.xls, sheet = 2, newName = "NewName2")
  expect_true(existsSheet(wb.xls, "NewName2"))
  expect_false(existsSheet(wb.xls, "OldName2"))

  # Check that renaming a sheet (using its index) works fine (*.xlsx)
  # (assumes 'createSheet' and 'existsSheet' to be working properly)
  createSheet(wb.xlsx, name = "OldName2")
  renameSheet(wb.xlsx, sheet = 2, newName = "NewName2")
  expect_true(existsSheet(wb.xlsx, "NewName2"))
  expect_false(existsSheet(wb.xlsx, "OldName2"))

  # Check that renaming a non-existing sheet throws an exception (*.xls)
  expect_error(renameSheet(wb.xls, sheet = "NonExisting", newName = "ShouldStillNotExist"), "IllegalArgumentException")

  # Check that renaming a non-existing sheet throws an exception (*.xlsx)
  expect_error(renameSheet(wb.xlsx, sheet = "NonExisting", newName = "ShouldStillNotExist"), "IllegalArgumentException")

  # Check that renaming a sheet with an invalid new name throws an exception (*.xls)
  createSheet(wb.xls, name = "SomeName")
  expect_error(renameSheet(wb.xls, sheet = "SomeName", newName = "'invalid"), "IllegalArgumentException")

  # Check that renaming a sheet with an invalid new name throws an exception (*.xlsx)
  createSheet(wb.xlsx, name = "SomeName")
  expect_error(renameSheet(wb.xlsx, sheet = "SomeName", newName = "'invalid"), "IllegalArgumentException")
})
