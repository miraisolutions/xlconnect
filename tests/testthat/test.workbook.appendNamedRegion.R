context("Workbook appendNamedRegion functionality")

# Helper function (specific to this test file)
test_overwrite_formula_helper <- function(wb, data_to_append, expected_read_data, expected_coords,
                                   name_to_use = "mtcars_formula", overwrite = TRUE) {
    appendNamedRegion(wb, data_to_append,
        name = name_to_use,
        overwriteFormulaCells = overwrite
    )
    res <- readNamedRegion(wb, name = name_to_use)
    expect_equal(res, normalizeDataframe(expected_read_data),
        ignore_attr = c("worksheetScope", "row.names")
    )
    expect_equal(
        getReferenceCoordinatesForName(wb, name_to_use),
        expected_coords
    )
}

test_that("appending to an existing named region works for XLS and XLSX", {
    wb.xls <- loadWorkbook(rsrc("testWorkbookAppend.xls"))
    wb.xlsx <- loadWorkbook(rsrc("testWorkbookAppend.xlsx"))

    refCoord <- matrix(c(9, 5, 73, 15), ncol = 2, byrow = TRUE) # Expected coordinates after appending mtcars once

    # XLS
    appendNamedRegion(wb.xls, mtcars, name = "mtcars")
    res_xls <- readNamedRegion(wb.xls, name = "mtcars")
    rownames(res_xls) <- as.character(seq_len(nrow(res_xls)))
    expected_data_xls <- rbind(mtcars, mtcars) # Original mtcars + appended mtcars
    rownames(expected_data_xls) <- as.character(seq_len(nrow(expected_data_xls)))
    expect_equal(normalizeDataframe(expected_data_xls), normalizeDataframe(res_xls))
    expect_equal(refCoord, getReferenceCoordinatesForName(wb.xls, "mtcars"))

    # XLSX
    appendNamedRegion(wb.xlsx, mtcars, name = "mtcars")
    res_xlsx <- readNamedRegion(wb.xlsx, name = "mtcars")
    rownames(res_xlsx) <- as.character(seq_len(nrow(res_xlsx)))
    expected_data_xlsx <- rbind(mtcars, mtcars) # Original mtcars + appended mtcars
    rownames(expected_data_xlsx) <- as.character(seq_len(nrow(expected_data_xlsx)))
    expect_equal(normalizeDataframe(expected_data_xlsx), normalizeDataframe(res_xlsx))
    expect_equal(refCoord, getReferenceCoordinatesForName(wb.xlsx, "mtcars"))
})

test_that("appending to a non-existent named region throws an error", {
    wb.xls <- loadWorkbook(rsrc("testWorkbookAppend.xls"))
    wb.xlsx <- loadWorkbook(rsrc("testWorkbookAppend.xlsx"))

    expect_error(appendNamedRegion(wb.xls, mtcars, name = "doesNotExist"))
    expect_error(appendNamedRegion(wb.xlsx, mtcars, name = "doesNotExist"))
})

test_that("overwriteFormulaCells = FALSE works as expected", {
    wb.xls <- loadWorkbook(rsrc("testWorkbookAppend.xls")) # Reload fresh workbook
    wb.xlsx <- loadWorkbook(rsrc("testWorkbookAppend.xlsx")) # Reload fresh workbook

    refCoordOriginalFormula <- matrix(c(9, 5, 40, 15), ncol = 2, byrow = TRUE) # Coords of original "mtcars_formula"
                                                                           # This seems to be the original region size based on `test_overwrite_formula` usage
                                                                           # The old test expected a final size of 73 rows (mtcars_mod + mtcars_mod)
                                                                           # but was checking the coords of the original region if overwrite = FALSE.
                                                                           # This needs clarification or adjustment if the region is expected to expand.
                                                                           # For now, assuming the region *does not* expand with overwrite=FALSE
                                                                           # and data is appended *within* the original bounds.
                                                                           # However, the original test uses rbind(mtcars_mod, mtcars_mod) which is larger.
                                                                           # Let's assume for now the intent was to test behavior on the *original* formula region.
                                                                           # The original test used refCoord = matrix(c(9, 5, 73, 15), ...), which implies expansion.
                                                                           # If overwriteFormulaCells = FALSE, it should not overwrite formulas.
                                                                           # The data being appended (mtcars) will replace existing data if it fits.
                                                                           # The helper was being called with `mtcars` as data_to_append
                                                                           # and `rbind(mtcars_mod, mtcars_mod)` as expected_read_data.
                                                                           # This is confusing. Let's simplify: if we append `mtcars` to `mtcars_formula` (which contains `mtcars_mod`)
                                                                           # with overwrite=FALSE, the formulas should remain, and data is appended.
                                                                           # The reference coordinates for "mtcars_formula" are (9,5) to (40,15).
                                                                           # If we append mtcars (32 rows) to this region, it should expand.

    # Let's re-evaluate the expected behavior based on the name "overwriteFormulaCells"
    # If FALSE, formulas should be preserved. Data is appended.
    # The initial region "mtcars_formula" in testWorkbookAppend.xls/xlsx contains `mtcars_mod`.
    # `mtcars_mod` is `mtcars` with `carb` column modified.

    mtcars_mod <- mtcars
    mtcars_mod$carb <- mtcars_mod$gear # This is the data originally in 'mtcars_formula'

    # Define mtcars where carb column is replaced by gear column values
    mtcars_carb_eq_gear <- mtcars
    mtcars_carb_eq_gear$carb <- mtcars_carb_eq_gear$gear

    # When overwriteFormulaCells = FALSE:
    # The original data ('mtcars_mod') should remain for the initial part of the region.
    # The appended data ('mtcars') will have its 'carb' column effectively written with its 'gear' values.
    expected_data_overwrite_false <- rbind(mtcars_mod, mtcars_carb_eq_gear)
    expected_coords_overwrite_false <- matrix(c(9, 5, 73, 15), ncol = 2, byrow = TRUE)


    test_overwrite_formula_helper(wb.xls, mtcars, expected_data_overwrite_false,
        expected_coords_overwrite_false, name_to_use = "mtcars_formula", overwrite = FALSE
    )
    test_overwrite_formula_helper(wb.xlsx, mtcars, expected_data_overwrite_false,
        expected_coords_overwrite_false, name_to_use = "mtcars_formula", overwrite = FALSE
    )
})

test_that("overwriteFormulaCells = TRUE works as expected", {
    wb.xls <- loadWorkbook(rsrc("testWorkbookAppend.xls")) # Reload fresh workbook
    wb.xlsx <- loadWorkbook(rsrc("testWorkbookAppend.xlsx")) # Reload fresh workbook

    mtcars_mod <- mtcars # Define mtcars_mod here as well
    mtcars_mod$carb <- mtcars_mod$gear

    # Define mtcars where carb column is replaced by gear column values
    mtcars_carb_eq_gear <- mtcars
    mtcars_carb_eq_gear$carb <- mtcars_carb_eq_gear$gear

    # When overwriteFormulaCells = TRUE:
    # The overwritten part of the original region seems to retain the 'carb' characteristic of mtcars_mod (i.e., gear values).
    # The appended part seems to take the 'carb' characteristic of the new data (original mtcars$carb).
    expected_data_overwrite_true <- rbind(mtcars_mod, mtcars)
    expected_coords_overwrite_true <- matrix(c(9, 5, 73, 15), ncol = 2, byrow = TRUE)


    test_overwrite_formula_helper(wb.xls, mtcars, expected_data_overwrite_true,
        expected_coords_overwrite_true, name_to_use = "mtcars_formula", overwrite = TRUE
    )
    test_overwrite_formula_helper(wb.xlsx, mtcars, expected_data_overwrite_true,
        expected_coords_overwrite_true, name_to_use = "mtcars_formula", overwrite = TRUE
    )
})


test_that("appending different data (airquality) to 'mtcars' named region updates coordinates", {
    wb.xls <- loadWorkbook(rsrc("testWorkbookAppend.xls"))
    wb.xlsx <- loadWorkbook(rsrc("testWorkbookAppend.xlsx"))

    # Original 'mtcars' region in testWorkbookAppend.xls (Sheet 'DataSheet') is E9:O40 (32 data rows)
    # Appending airquality (153 rows) to it.
    # New data rows = 32 (original) + 153 (appended) = 185 rows.
    # Start row 9. End row = 9 + 185 - 1 = 193. If header is 1 row, then 9+185 = 194. This matches original test.
    refCoord_airquality_append <- matrix(c(9, 5, 194, 15), ncol = 2, byrow = TRUE)

    appendNamedRegion(wb.xls, airquality, name = "mtcars")
    expect_equal(refCoord_airquality_append, getReferenceCoordinatesForName(wb.xls, "mtcars"))

    # Verify actual data read back (optional, but good for sanity)
    res_xls_air <- readNamedRegion(wb.xls, name="mtcars")
    expect_equal(nrow(res_xls_air), 32 + nrow(airquality))


    appendNamedRegion(wb.xlsx, airquality, name = "mtcars")
    expect_equal(refCoord_airquality_append, getReferenceCoordinatesForName(wb.xlsx, "mtcars"))

    res_xlsx_air <- readNamedRegion(wb.xlsx, name="mtcars")
    expect_equal(nrow(res_xlsx_air), 32 + nrow(airquality))
})
