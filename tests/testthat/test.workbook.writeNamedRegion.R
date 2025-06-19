test_that("test.workbook.writeNamedRegion", {
    wb.xls <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xls"),
        create = TRUE)
    wb.xlsx <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xlsx"),
        create = TRUE)
    createName(wb.xls, "test1", "Test1!$C$8")
    expect_error(writeNamedRegion(wb.xls, search, "test1"))
    createName(wb.xlsx, "test1", "Test1!$C$8")
    expect_error(writeNamedRegion(wb.xlsx, search, "test1"))
    expect_error(writeNamedRegion(wb.xls, mtcars, "nameDoesNotExist"))
    expect_error(writeNamedRegion(wb.xlsx, mtcars, "nameDoesNotExist"))
    createName(wb.xls, "nope", "NonExistingSheet!A1")
    expect_error(writeNamedRegion(wb.xls, mtcars, "nope"))
    createName(wb.xlsx, "nope", "NonExistingSheet!A1")
    expect_error(writeNamedRegion(wb.xlsx, mtcars, "nope"))
    createSheet(wb.xls, "empty")
    createName(wb.xls, "empty1", "empty!A1")
    createName(wb.xls, "empty2", "empty!D10")
    checkNoException(writeNamedRegion(wb.xls, data.frame(), "empty1"))
    checkNoException(writeNamedRegion(wb.xls, data.frame(a = character(0), 
        b = numeric(0)), "empty2"))
    createSheet(wb.xlsx, "empty")
    createName(wb.xlsx, "empty1", "empty!A1")
    createName(wb.xlsx, "empty2", "empty!D10")
    checkNoException(writeNamedRegion(wb.xlsx, data.frame(), 
        "empty1"))
    checkNoException(writeNamedRegion(wb.xlsx, data.frame(a = character(0), 
        b = numeric(0)), "empty2"))
    wb.xls <- loadWorkbook("resources/testWorkbookOverwriteFormulas.xls"))
    wb.xlsx <- loadWorkbook("resources/testWorkbookOverwriteFormulas.xlsx"))
    mtcars_mod = mtcars
    mtcars_mod$carb = mtcars_mod$gear
    test_overwrite_formula <- function(wb, expected, overwrite = TRUE) {
        writeNamedRegion(wb, mtcars, "mtcars_formula", overwriteFormulaCells = overwrite)
        res = readNamedRegion(wb, "mtcars_formula")
        checkEquals(res, normalizeDataframe(expected), check.attributes = FALSE, 
            check.names = TRUE)
    }
    test_overwrite_formula(wb.xls, mtcars_mod, overwrite = FALSE)
    test_overwrite_formula(wb.xlsx, mtcars_mod, overwrite = FALSE)
    test_overwrite_formula(wb.xls, mtcars)
    test_overwrite_formula(wb.xlsx, mtcars)
})

