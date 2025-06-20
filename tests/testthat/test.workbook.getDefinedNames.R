test_that("test.workbook.getDefinedNames", {
    wb.xls <- loadWorkbook("resources/testWorkbookDefinedNames.xls"",
        create = FALSE)
    wb.xlsx <- loadWorkbook("resources/testWorkbookDefinedNames.xlsx"",
        create = FALSE)
    expectedNamesValidOnly <- c("FirstName", "SecondName", "FourthName", 
        "FifthName")
    expectedNamesAll <- c("FirstName", "SecondName", "ThirdName", 
        "FourthName", "FifthName")
    definedNames <- getDefinedNames(wb.xls, validOnly = TRUE)
    expect_true(length(setdiff(expectedNamesValidOnly, definedNames)) == 
        0 && length(setdiff(definedNames, expectedNamesValidOnly)) == 
        0)
    definedNames <- getDefinedNames(wb.xls, validOnly = FALSE)
    expect_true(length(setdiff(expectedNamesAll, definedNames)) == 
        0 && length(setdiff(definedNames, expectedNamesAll)) == 
        0)
    definedNames <- getDefinedNames(wb.xlsx, validOnly = TRUE)
    expect_true(length(setdiff(expectedNamesValidOnly, definedNames)) == 
        0 && length(setdiff(definedNames, expectedNamesValidOnly)) == 
        0)
    definedNames <- getDefinedNames(wb.xlsx, validOnly = FALSE)
    expect_true(length(setdiff(expectedNamesAll, definedNames)) == 
        0 && length(setdiff(definedNames, expectedNamesAll)) == 
        0)
    wb.xls <- loadWorkbook("resources/testWorkbookGetDefinedNamesScoped.xls"",
        create = FALSE)
    wb.xlsx <- loadWorkbook("resources/testWorkbookGetDefinedNamesScoped.xlsx"",
        create = FALSE)
    expectedNames <- c("ScopedName1", "ScopedName2")
    expectedScopes <- c("scoped_1", "scoped_2")
    res_xls <- getDefinedNames(wb.xls, validOnly = TRUE)
    expect_true(setequal(res_xls, expectedNames))
    expect_true(setequal(attributes(res_xls)$worksheetScope, 
        expectedScopes))
    res_xlsx <- getDefinedNames(wb.xlsx, validOnly = TRUE)
    expect_true(setequal(res_xlsx, expectedNames))
    expect_true(setequal(attributes(res_xlsx)$worksheetScope, 
        expectedScopes))
})

