test_that("test.workbook.getReferenceCoordinates", {
    wb.xls <- loadWorkbook("resources/testWorkbookReferenceFormula.xls"),
        create = FALSE)
    wb.xlsx <- loadWorkbook("resources/testWorkbookReferenceFormula.xlsx"),
        create = FALSE)
    expect_true(all(getReferenceCoordinatesForName(wb.xls, "FirstName") == 
        matrix(c(1, 1, 1, 1), nrow = 2, byrow = TRUE)))
    expect_error(getReferenceCoordinatesForName(wb.xls, "NonExistentName"))
    expect_true(all(getReferenceCoordinatesForName(wb.xlsx, "FirstName") == 
        matrix(c(1, 1, 1, 1), nrow = 2, byrow = TRUE)))
    expect_error(getReferenceCoordinatesForName(wb.xlsx, "NonExistentName"))
})

