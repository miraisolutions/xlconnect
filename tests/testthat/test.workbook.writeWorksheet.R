test_that("test.workbook.writeWorksheet", {
    wb.xls <- loadWorkbook("resources/testWorkbookWriteWorksheet.xls"),
        create = TRUE)
    wb.xlsx <- loadWorkbook("resources/testWorkbookWriteWorksheet.xlsx"),
        create = TRUE)
    createSheet(wb.xls, "test1")
    expect_error(writeWorksheet(wb.xls, search, "test1"))
    createSheet(wb.xlsx, "test1")
    expect_error(writeWorksheet(wb.xlsx, search, "test1"))
    expect_error(writeWorksheet(wb.xls, mtcars, "sheetDoesNotExist"))
    expect_error(writeWorksheet(wb.xlsx, mtcars, "sheetDoesNotExist"))
    wb.xls <- loadWorkbook("resources/testWorkbookOverwriteFormulas.xls"))
    wb.xlsx <- loadWorkbook("resources/testWorkbookOverwriteFormulas.xlsx"))
    mtcars_mod = mtcars
    mtcars_mod$carb = mtcars_mod$gear
    test_overwrite_formula <- function(wb, expected, overwrite = TRUE) {
        writeWorksheet(wb, mtcars, "mtcars_formula", startRow = 9, 
            startCol = 5, overwriteFormulaCells = overwrite)
        res = readWorksheet(wb, "mtcars_formula")
        checkEquals(res, normalizeDataframe(expected), check.attributes = FALSE, 
            check.names = TRUE)
    }
    test_overwrite_formula(wb.xls, mtcars_mod, overwrite = FALSE)
    test_overwrite_formula(wb.xlsx, mtcars_mod, overwrite = FALSE)
    test_overwrite_formula(wb.xls, mtcars)
    test_overwrite_formula(wb.xlsx, mtcars)
})

