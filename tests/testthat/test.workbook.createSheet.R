test_that("test.workbook.createSheet", {
    wb.xls <- loadWorkbook("resources/createSheet.xls",
        create = TRUE
    )
    wb.xlsx <- loadWorkbook("resources/createSheet.xlsx",
        create = TRUE
    )
    expect_error(createSheet(wb.xls, "'Invalid Sheet Name"), NA)
    expect_error(createSheet(wb.xlsx, "'Invalid Sheet Name"), NA)
    expect_error(createSheet(wb.xls, "Invalid Sheet Name'"), NA)
    expect_error(createSheet(wb.xlsx, "Invalid Sheet Name'"), NA)
    expect_error(createSheet(wb.xls, "A very very very very very very very very long name"), NA)
    expect_error(createSheet(wb.xlsx, "A very very very very very very very very long name"), NA)
    sheetName <- "My Sheet"
    try(createSheet(wb.xls, sheetName))
    expect_true(existsSheet(wb.xls, sheetName))
    try(createSheet(wb.xlsx, sheetName))
    expect_true(existsSheet(wb.xlsx, sheetName))
    try(createSheet(wb.xls, sheetName))
    expect_true(existsSheet(wb.xls, sheetName))
    try(createSheet(wb.xlsx, sheetName))
    expect_true(existsSheet(wb.xlsx, sheetName))
})
