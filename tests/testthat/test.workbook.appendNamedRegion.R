test_that("test.workbook.appendNamedRegion", {
    test_overwrite_formula <- function(wb, expected, coords,
                                       overwrite = TRUE) {
        appendNamedRegion(wb, mtcars,
            name = "mtcars_formula",
            overwriteFormulaCells = overwrite
        )
        res <- readNamedRegion(wb, name = "mtcars_formula")
        expect_equal(res, normalizeDataframe(expected),
            check.attributes = FALSE,
            check.names = TRUE
        )
        expect_equal(
            getReferenceCoordinatesForName(wb, "mtcars_formula"),
            coords
        )
    }
    wb.xls <- loadWorkbook("resources/testWorkbookAppend.xls")
    wb.xlsx <- loadWorkbook("resources/testWorkbookAppend.xlsx")
    refCoord <- matrix(c(9, 5, 73, 15), ncol = 2, byrow = TRUE)
    appendNamedRegion(wb.xls, mtcars, name = "mtcars")
    res <- readNamedRegion(wb.xls, name = "mtcars")
    rownames(res) <- as.character(seq_len(nrow(res)))
    expected_data_xls <- rbind(mtcars, mtcars)
    rownames(expected_data_xls) <- as.character(seq_len(nrow(expected_data_xls)))
    expect_equal(normalizeDataframe(expected_data_xls), normalizeDataframe(res))
    expect_equal(refCoord, getReferenceCoordinatesForName(
        wb.xls,
        "mtcars"
    ))
    appendNamedRegion(wb.xlsx, mtcars, name = "mtcars")
    res <- readNamedRegion(wb.xlsx, name = "mtcars")
    rownames(res) <- as.character(seq_len(nrow(res)))
    expected_data_xlsx <- rbind(mtcars, mtcars)
    rownames(expected_data_xlsx) <- as.character(seq_len(nrow(expected_data_xlsx)))
    expect_equal(normalizeDataframe(expected_data_xlsx), normalizeDataframe(res))
    expect_equal(refCoord, getReferenceCoordinatesForName(
        wb.xlsx,
        "mtcars"
    ))
    expect_error(appendNamedRegion(wb.xls, mtcars, name = "doesNotExist"))
    expect_error(appendNamedRegion(wb.xlsx, mtcars, name = "doesNotExist"))
    mtcars_mod <- mtcars
    mtcars_mod$carb <- mtcars_mod$gear
    test_overwrite_formula(wb.xls, rbind(mtcars_mod, mtcars_mod),
        refCoord,
        overwrite = FALSE
    )
    test_overwrite_formula(wb.xlsx, rbind(mtcars_mod, mtcars_mod),
        refCoord,
        overwrite = FALSE
    )
    wb.xls <- loadWorkbook("resources/testWorkbookAppend.xls")
    wb.xlsx <- loadWorkbook("resources/testWorkbookAppend.xlsx")
    refCoord <- matrix(c(9, 5, 194, 15), ncol = 2, byrow = TRUE)
    refCoordFormula <- matrix(c(9, 5, 73, 15), ncol = 2, byrow = TRUE)
    appendNamedRegion(wb.xls, airquality, name = "mtcars")
    expect_equal(refCoord, getReferenceCoordinatesForName(
        wb.xls,
        "mtcars"
    ))
    appendNamedRegion(wb.xlsx, airquality, name = "mtcars")
    expect_equal(refCoord, getReferenceCoordinatesForName(
        wb.xlsx,
        "mtcars"
    ))
    test_overwrite_formula(
        wb.xls, rbind(mtcars_mod, mtcars),
        refCoordFormula
    )
    test_overwrite_formula(
        wb.xlsx, rbind(mtcars_mod, mtcars),
        refCoordFormula
    )
})
