test_that("test.workbook.getSheets", {
    wb.xls <- loadWorkbook("resources/testWorkbookSheets.xls",
        create = FALSE
    )
    wb.xlsx <- loadWorkbook("resources/testWorkbookSheets.xlsx",
        create = FALSE
    )
    expectedSheets <- c(
        "A1", "B 2", "$$", "=", "@}", "11. Oct.",
        "\"quote\"", "+0"
    )
    definedSheets <- getSheets(wb.xls)
    expect_true(length(setdiff(expectedSheets, definedSheets)) ==
        0 && length(setdiff(definedSheets, expectedSheets)) ==
        0)
    definedSheets <- getSheets(wb.xlsx)
    expect_true(length(setdiff(expectedSheets, definedSheets)) ==
        0 && length(setdiff(definedSheets, expectedSheets)) ==
        0)
})
