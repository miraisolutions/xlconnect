test_that("removeSheet removes an existing sheet in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookRemoveSheet.xls", create = FALSE)
    removeSheet(wb.xls, "BBB")
    expect_false(existsSheet(wb.xls, "BBB"))
})

test_that("removeSheet removes an existing sheet in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookRemoveSheet.xlsx", create = FALSE)
    removeSheet(wb.xlsx, "BBB")
    expect_false(existsSheet(wb.xlsx, "BBB"))
})

test_that("removeSheet does not throw an error for a non-existent sheet in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookRemoveSheet.xls", create = FALSE)
    expect_error(removeSheet(wb.xls, 35), NA)
    expect_error(removeSheet(wb.xls, "SheetWhichDoesNotExist"), NA)
})

test_that("removeSheet does not throw an error for a non-existent sheet in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookRemoveSheet.xlsx", create = FALSE)
    expect_error(removeSheet(wb.xlsx, 35), NA)
    expect_error(removeSheet(wb.xlsx, "SheetWhichDoesNotExist"), NA)
})
