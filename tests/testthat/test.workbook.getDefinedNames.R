test_that("getDefinedNames returns valid names only - check that all and only the expected names exist (*.xls)", {
  wb.xls <- loadWorkbook("resources/testWorkbookDefinedNames.xls", create = FALSE)

  # Names defined in workbooks
  expectedNamesValidOnly <- c("FirstName", "SecondName", "FourthName", "FifthName")
  definedNames <- getDefinedNames(wb.xls, validOnly = TRUE)
  expect_true(setequal(expectedNamesValidOnly, definedNames))
})

test_that("getDefinedNames returns all names - check that all and only the expected names exist (*.xls)", {
  wb.xls <- loadWorkbook("resources/testWorkbookDefinedNames.xls", create = FALSE)

  # Names defined in workbooks
  expectedNamesAll <- c("FirstName", "SecondName", "ThirdName", "FourthName", "FifthName")
  definedNames <- getDefinedNames(wb.xls, validOnly = FALSE)
  expect_true(setequal(expectedNamesAll, definedNames))
})

test_that("getDefinedNames returns valid names only - check that all and only the expected names exist (*.xlsx)", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookDefinedNames.xlsx", create = FALSE)

  # Names defined in workbooks
  expectedNamesValidOnly <- c("FirstName", "SecondName", "FourthName", "FifthName")
  definedNames <- getDefinedNames(wb.xlsx, validOnly = TRUE)
  expect_true(setequal(expectedNamesValidOnly, definedNames))
})

test_that("getDefinedNames returns all names - check that all and only the expected names exist (*.xlsx)", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookDefinedNames.xlsx", create = FALSE)

  # Names defined in workbooks
  expectedNamesAll <- c("FirstName", "SecondName", "ThirdName", "FourthName", "FifthName")
  definedNames <- getDefinedNames(wb.xlsx, validOnly = FALSE)
  expect_true(setequal(expectedNamesAll, definedNames))
})

test_that("getDefinedNames returns scoped names correctly in XLS", {
  # scoped names
  wb.xls <- loadWorkbook("resources/testWorkbookGetDefinedNamesScoped.xls", create = FALSE)

  expectedNames <- c("ScopedName1", "ScopedName2")
  expectedScopes <- c("scoped_1", "scoped_2")

  res_xls <- getDefinedNames(wb.xls, validOnly = TRUE)
  expect_true(setequal(res_xls, expectedNames))
  expect_true(setequal(attributes(res_xls)$worksheetScope, expectedScopes))
})

test_that("getDefinedNames returns scoped names correctly in XLSX", {
  # scoped names
  wb.xlsx <- loadWorkbook("resources/testWorkbookGetDefinedNamesScoped.xlsx", create = FALSE)

  expectedNames <- c("ScopedName1", "ScopedName2")
  expectedScopes <- c("scoped_1", "scoped_2")

  res_xlsx <- getDefinedNames(wb.xlsx, validOnly = TRUE)
  expect_true(setequal(res_xlsx, expectedNames))
  expect_true(setequal(attributes(res_xlsx)$worksheetScope, expectedScopes))
})
