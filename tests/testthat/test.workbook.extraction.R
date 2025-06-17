test_that("test.workbook.extraction", {
    wb.xls <- loadWorkbook(rsrc("resources/testWorkbookExtractionOperators.xls"), 
        create = TRUE)
    wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookExtractionOperators.xlsx"), 
        create = TRUE)
    wb.xls["mtcars1"] = mtcars
    expect_true("mtcars1" %in% getSheets(wb.xls))
    expect_equal(33, as.vector(getLastRow(wb.xls, "mtcars1")))
    wb.xls["mtcars2", startRow = 6, startCol = 11, header = FALSE] = mtcars
    expect_true("mtcars2" %in% getSheets(wb.xls))
    expect_equal(37, as.vector(getLastRow(wb.xls, "mtcars2")))
    wb.xlsx["mtcars1"] = mtcars
    expect_true("mtcars1" %in% getSheets(wb.xlsx))
    expect_equal(33, as.vector(getLastRow(wb.xlsx, "mtcars1")))
    wb.xlsx["mtcars2", startRow = 6, startCol = 11, header = FALSE] = mtcars
    expect_true("mtcars2" %in% getSheets(wb.xlsx))
    expect_equal(37, as.vector(getLastRow(wb.xlsx, "mtcars2")))
    expect_equal(c(32, 11), dim(wb.xls["mtcars1"]))
    expect_equal(c(31, 11), dim(wb.xls["mtcars2"]))
    expect_equal(c(32, 11), dim(wb.xlsx["mtcars1"]))
    expect_equal(c(31, 11), dim(wb.xlsx["mtcars2"]))
    wb.xls[["mtcars3", "mtcars3!$B$7"]] = mtcars
    expect_true("mtcars3" %in% getDefinedNames(wb.xls))
    expect_equal(39, as.vector(getLastRow(wb.xls, "mtcars3")))
    wb.xls[["mtcars4", "mtcars4!$D$8", rownames = "Car"]] = mtcars
    expect_true("mtcars4" %in% getDefinedNames(wb.xls))
    expect_equal(40, as.vector(getLastRow(wb.xls, "mtcars4")))
    wb.xlsx[["mtcars3", "mtcars3!$B$7"]] = mtcars
    expect_true("mtcars3" %in% getDefinedNames(wb.xlsx))
    expect_equal(39, as.vector(getLastRow(wb.xlsx, "mtcars3")))
    wb.xlsx[["mtcars4", "mtcars4!$D$8", rownames = "Car"]] = mtcars
    expect_true("mtcars4" %in% getDefinedNames(wb.xlsx))
    expect_equal(40, as.vector(getLastRow(wb.xlsx, "mtcars4")))
    expect_equal(c(32, 11), dim(wb.xls[["mtcars3"]]))
    expect_equal(c(32, 12), dim(wb.xls[["mtcars4"]]))
    expect_equal(c(32, 11), dim(wb.xlsx[["mtcars3"]]))
    expect_equal(c(32, 12), dim(wb.xlsx[["mtcars4"]]))
})

