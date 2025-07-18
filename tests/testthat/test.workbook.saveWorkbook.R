test_that("saveWorkbook saves an XLS file", {
    skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
    file.xls <- "testWorkbookSaveWorkbook.xls"
    on.exit(if (file.exists(file.xls)) file.remove(file.xls))

    wb.xls <- loadWorkbook(file.xls, create = TRUE)
    expect_false(file.exists(file.xls))
    saveWorkbook(wb.xls)
    expect_true(file.exists(file.xls))
})

test_that("saveWorkbook saves an XLSX file", {
    skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
    file.xlsx <- "testWorkbookSaveWorkbook.xlsx"
    on.exit(if (file.exists(file.xlsx)) file.remove(file.xlsx))

    wb.xlsx <- loadWorkbook(file.xlsx, create = TRUE)
    expect_false(file.exists(file.xlsx))
    saveWorkbook(wb.xlsx)
    expect_true(file.exists(file.xlsx))
})

test_that("saveWorkbook saves an XLS file to a new location", {
    skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
    file.xls <- "testWorkbookSaveWorkbook.xls"
    newFile.xls <- "saveAsWorkbook.xls"
    on.exit({
        if (file.exists(file.xls)) file.remove(file.xls)
        if (file.exists(newFile.xls)) file.remove(newFile.xls)
    })

    wb.xls <- loadWorkbook(file.xls, create = TRUE)
    saveWorkbook(wb.xls, file = newFile.xls)
    expect_true(file.exists(newFile.xls))
})

test_that("saveWorkbook saves an XLSX file to a new location", {
    skip_if_not(getOption("FULL.TEST.SUITE"), "FULL.TEST.SUITE is not TRUE")
    file.xlsx <- "testWorkbookSaveWorkbook.xlsx"
    newFile.xlsx <- "saveAsWorkbook.xlsx"
    on.exit({
        if (file.exists(file.xlsx)) file.remove(file.xlsx)
        if (file.exists(newFile.xlsx)) file.remove(newFile.xlsx)
    })

    wb.xlsx <- loadWorkbook(file.xlsx, create = TRUE)
    saveWorkbook(wb.xlsx, file = newFile.xlsx)
    expect_true(file.exists(newFile.xlsx))
})
