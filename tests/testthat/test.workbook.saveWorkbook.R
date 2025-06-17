test_that("test.workbook.saveWorkbook", {
    if (getOption("FULL.TEST.SUITE")) {
        file.xls <- rsrc("resources/testWorkbookSaveWorkbook.xls")
        file.xlsx <- rsrc("resources/testWorkbookSaveWorkbook.xlsx")
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
        if (file.exists(newFile.xls)) 
            file.remove(newFile.xls)
        saveWorkbook(wb.xls, file = newFile.xls)
        expect_true(file.exists(newFile.xls))
        newFile.xlsx <- "saveAsWorkbook.xlsx"
        if (file.exists(newFile.xlsx)) 
            file.remove(newFile.xlsx)
        saveWorkbook(wb.xlsx, file = newFile.xlsx)
        expect_true(file.exists(newFile.xlsx))
    }
})

