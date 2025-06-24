test_that("test.workbook.createSheet", {
    wb.xls <- loadWorkbook("resources/createSheet.xls",
        create = TRUE
    )
    wb.xlsx <- loadWorkbook("resources/createSheet.xlsx",
        create = TRUE
    )
    expect_error(createSheet(wb.xls, "'Invalid Sheet Name"), "IllegalArgumentException")
    expect_error(createSheet(wb.xlsx, "'Invalid Sheet Name"), "IllegalArgumentException")
    expect_error(createSheet(wb.xls, "Invalid Sheet Name'"), "IllegalArgumentException")
    expect_error(createSheet(wb.xlsx, "Invalid Sheet Name'"), "IllegalArgumentException")
    expect_error(createSheet(wb.xls, "A very very very very very very very very long name"), "IllegalArgumentException")
    expect_error(createSheet(wb.xlsx, "A very very very very very very very very long name"), "IllegalArgumentException")
    sheetName <- "My Sheet"
    # createSheet should not throw an error for valid sheet names
    createSheet(wb.xls, sheetName)
    expect_true(existsSheet(wb.xls, sheetName))
    createSheet(wb.xlsx, sheetName)
    expect_true(existsSheet(wb.xlsx, sheetName))
    # Creating an existing sheet should also not throw an error (it's a NOP)
    createSheet(wb.xls, sheetName)
    expect_true(existsSheet(wb.xls, sheetName))
    createSheet(wb.xlsx, sheetName)
    expect_true(existsSheet(wb.xlsx, sheetName))
})
