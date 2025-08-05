test_that("getDefinedNames returns valid names only in XLS", {
  wb.xls <- loadWorkbook("resources/testWorkbookDefinedNames.xls", create = FALSE)
  expectedNamesValidOnly <- c("FirstName", "SecondName", "FourthName", "FifthName")
  definedNames <- getDefinedNames(wb.xls, validOnly = TRUE)
  expect_true(setequal(expectedNamesValidOnly, definedNames))
})

test_that("getDefinedNames returns all names in XLS", {
  wb.xls <- loadWorkbook("resources/testWorkbookDefinedNames.xls", create = FALSE)
  expectedNamesAll <- c("FirstName", "SecondName", "ThirdName", "FourthName", "FifthName")
  definedNames <- getDefinedNames(wb.xls, validOnly = FALSE)
  expect_true(setequal(expectedNamesAll, definedNames))
})

test_that("getDefinedNames returns valid names only in XLSX", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookDefinedNames.xlsx", create = FALSE)
  expectedNamesValidOnly <- c("FirstName", "SecondName", "FourthName", "FifthName")
  definedNames <- getDefinedNames(wb.xlsx, validOnly = TRUE)
  expect_true(setequal(expectedNamesValidOnly, definedNames))
})

test_that("getDefinedNames returns all names in XLSX", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookDefinedNames.xlsx", create = FALSE)
  expectedNamesAll <- c("FirstName", "SecondName", "ThirdName", "FourthName", "FifthName")
  definedNames <- getDefinedNames(wb.xlsx, validOnly = FALSE)
  expect_true(setequal(expectedNamesAll, definedNames))
})

test_that("getDefinedNames returns scoped names correctly in XLS", {
  wb.xls <- loadWorkbook("resources/testWorkbookGetDefinedNamesScoped.xls", create = FALSE)
  expectedNames <- c("ScopedName1", "ScopedName2")
  expectedScopes <- c("scoped_1", "scoped_2")
  res_xls <- getDefinedNames(wb.xls, validOnly = TRUE)
  expect_true(setequal(res_xls, expectedNames))
})

test_that("getDefinedNames returns scoped names correctly in XLSX", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookGetDefinedNamesScoped.xlsx", create = FALSE)
  expectedNames <- c("ScopedName1", "ScopedName2")
  expectedScopes <- c("scoped_1", "scoped_2")
  res_xlsx <- getDefinedNames(wb.xlsx, validOnly = TRUE)
  expect_true(setequal(res_xlsx, expectedNames))
})
