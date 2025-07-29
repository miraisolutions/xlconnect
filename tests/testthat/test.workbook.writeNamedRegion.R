test_that("writeNamedRegion throws error for non-data.frame objects in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xls", create = TRUE)
    createName(wb.xls, "test1", "Test1!$C$8")
    expect_error(writeNamedRegion(wb.xls, search, "test1"), "cannot coerce class")
})

test_that("writeNamedRegion throws error for non-data.frame objects in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xlsx", create = TRUE)
    createName(wb.xlsx, "test1", "Test1!$C$8")
    expect_error(writeNamedRegion(wb.xlsx, search, "test1"), "cannot coerce class")
})

test_that("writeNamedRegion throws error for non-existent names in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xls", create = TRUE)
    expect_error(writeNamedRegion(wb.xls, mtcars, "nameDoesNotExist"), "IllegalArgumentException")
})

test_that("writeNamedRegion throws error for non-existent names in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xlsx", create = TRUE)
    expect_error(writeNamedRegion(wb.xlsx, mtcars, "nameDoesNotExist"), "IllegalArgumentException")
})

test_that("writeNamedRegion throws error for non-existent sheets in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xls", create = TRUE)
    createName(wb.xls, "nope", "NonExistingSheet!A1")
    expect_error(writeNamedRegion(wb.xls, mtcars, "nope"), "IllegalArgumentException")
})

test_that("writeNamedRegion throws error for non-existent sheets in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xlsx", create = TRUE)
    createName(wb.xlsx, "nope", "NonExistingSheet!A1")
    expect_error(writeNamedRegion(wb.xlsx, mtcars, "nope"), "IllegalArgumentException")
})

test_that("writeNamedRegion handles empty data frames in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xls", create = TRUE)
    createSheet(wb.xls, "empty")
    createName(wb.xls, "empty1", "empty!A1")
    createName(wb.xls, "empty2", "empty!D10")
    writeNamedRegion(wb.xls, data.frame(), "empty1")
    writeNamedRegion(wb.xls, data.frame(a = character(0), b = numeric(0)), "empty2")
    expect_true(TRUE)
})

test_that("writeNamedRegion handles empty data frames in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookWriteNamedRegion.xlsx", create = TRUE)
    createSheet(wb.xlsx, "empty")
    createName(wb.xlsx, "empty1", "empty!A1")
    createName(wb.xlsx, "empty2", "empty!D10")
    writeNamedRegion(wb.xlsx, data.frame(), "empty1")
    writeNamedRegion(wb.xlsx, data.frame(a = character(0), b = numeric(0)), "empty2")
    expect_true(TRUE)
})

test_that("writeNamedRegion with overwriteFormulaCells = FALSE works in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookOverwriteFormulas.xls")
    mtcars_mod <- mtcars
    mtcars_mod$carb <- mtcars_mod$gear
    writeNamedRegion(wb.xls, mtcars, "mtcars_formula", overwriteFormulaCells = FALSE)
    res <- readNamedRegion(wb.xls, "mtcars_formula")
    expect_equal(res, normalizeDataframe(mtcars_mod), ignore_attr = TRUE)
})

test_that("writeNamedRegion with overwriteFormulaCells = FALSE works in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookOverwriteFormulas.xlsx")
    mtcars_mod <- mtcars
    mtcars_mod$carb <- mtcars_mod$gear
    writeNamedRegion(wb.xlsx, mtcars, "mtcars_formula", overwriteFormulaCells = FALSE)
    res <- readNamedRegion(wb.xlsx, "mtcars_formula")
    expect_equal(res, normalizeDataframe(mtcars_mod), ignore_attr = TRUE)
})

test_that("writeNamedRegion with overwriteFormulaCells = TRUE works in XLS", {
    wb.xls <- loadWorkbook("resources/testWorkbookOverwriteFormulas.xls")
    writeNamedRegion(wb.xls, mtcars, "mtcars_formula", overwriteFormulaCells = TRUE)
    res <- readNamedRegion(wb.xls, "mtcars_formula")
    expect_equal(res, normalizeDataframe(mtcars), ignore_attr = TRUE)
})

test_that("writeNamedRegion with overwriteFormulaCells = TRUE works in XLSX", {
    wb.xlsx <- loadWorkbook("resources/testWorkbookOverwriteFormulas.xlsx")
    writeNamedRegion(wb.xlsx, mtcars, "mtcars_formula", overwriteFormulaCells = TRUE)
    res <- readNamedRegion(wb.xlsx, "mtcars_formula")
    expect_equal(res, normalizeDataframe(mtcars), ignore_attr = TRUE)
})
