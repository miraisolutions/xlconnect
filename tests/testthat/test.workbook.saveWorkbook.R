test_that("test.workbook.saveWorkbook - always run", {
    # Add tests here that should always run, if any.
    # For now, this block will be empty as all existing tests are conditional.
    expect_true(TRUE) # Placeholder to ensure the test block is not empty
})

test_that("test.workbook.saveWorkbook - full test suite only", {
    skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")

    file.xls <- "testWorkbookSaveWorkbook.xls"
    file.xlsx <- "testWorkbookSaveWorkbook.xlsx"
    file.remove(file.xls)
    file.remove(file.xlsx)
    wb.xls <- loadWorkbook(file.xls, create = TRUE)
    wb.xlsx <- loadWorkbook(file.xlsx, create = TRUE)
    expect_false(file.exists(file.xls))
    expect_false(file.exists(file.xlsx))
    saveWorkbook(wb.xls)
    saveWorkbook(wb.xlsx)
    expect_true(file.exists(file.xls))
    expect_true(file.exists(file.xlsx))
    newFile.xls <- "saveAsWorkbook.xls"
    if (file.exists(newFile.xls)) {
        file.remove(newFile.xls)
    }
    saveWorkbook(wb.xls, file = newFile.xls)
    expect_true(file.exists(newFile.xls))
    newFile.xlsx <- "saveAsWorkbook.xlsx"
    if (file.exists(newFile.xlsx)) {
        file.remove(newFile.xlsx)
    }
    saveWorkbook(wb.xlsx, file = newFile.xlsx)
    expect_true(file.exists(newFile.xlsx))
})
