context("Workbook appendNamedRegion functionality")

# Helper function (specific to this test file)
test_overwrite_formula_helper <- function(
  wb,
  data_to_append,
  expected_read_data,
  expected_coords,
  name_to_use = "mtcars_formula",
  overwrite = TRUE
) {
  appendNamedRegion(wb, data_to_append, name = name_to_use, overwriteFormulaCells = overwrite)
  res <- readNamedRegion(wb, name = name_to_use)
  expect_equal(res, normalizeDataframe(expected_read_data), ignore_attr = c("worksheetScope", "row.names"))
  expect_equal(getReferenceCoordinatesForName(wb, name_to_use), expected_coords)
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

  mtcars_mod <- mtcars
  mtcars_mod$carb <- mtcars_mod$gear

  mtcars_carb_eq_gear <- mtcars
  mtcars_carb_eq_gear$carb <- mtcars_carb_eq_gear$gear

  expected_data_overwrite_false <- rbind(mtcars_mod, mtcars_carb_eq_gear)
  expected_coords_overwrite_false <- matrix(c(9, 5, 73, 15), ncol = 2, byrow = TRUE)

  test_overwrite_formula_helper(
    wb.xls,
    mtcars,
    expected_data_overwrite_false,
    expected_coords_overwrite_false,
    name_to_use = "mtcars_formula",
    overwrite = FALSE
  )
  test_overwrite_formula_helper(
    wb.xlsx,
    mtcars,
    expected_data_overwrite_false,
    expected_coords_overwrite_false,
    name_to_use = "mtcars_formula",
    overwrite = FALSE
  )
})

test_that("overwriteFormulaCells = TRUE works as expected", {
  wb.xls <- loadWorkbook(rsrc("testWorkbookAppend.xls"))
  wb.xlsx <- loadWorkbook(rsrc("testWorkbookAppend.xlsx"))

  mtcars_mod <- mtcars
  mtcars_mod$carb <- mtcars_mod$gear

  mtcars_carb_eq_gear <- mtcars
  mtcars_carb_eq_gear$carb <- mtcars_carb_eq_gear$gear

  expected_data_overwrite_true <- rbind(mtcars_mod, mtcars)
  expected_coords_overwrite_true <- matrix(c(9, 5, 73, 15), ncol = 2, byrow = TRUE)

  test_overwrite_formula_helper(
    wb.xls,
    mtcars,
    expected_data_overwrite_true,
    expected_coords_overwrite_true,
    name_to_use = "mtcars_formula",
    overwrite = TRUE
  )
  test_overwrite_formula_helper(
    wb.xlsx,
    mtcars,
    expected_data_overwrite_true,
    expected_coords_overwrite_true,
    name_to_use = "mtcars_formula",
    overwrite = TRUE
  )
})


test_that("appending different data (airquality) to 'mtcars' named region updates coordinates", {
  wb.xls <- loadWorkbook(rsrc("testWorkbookAppend.xls"))
  wb.xlsx <- loadWorkbook(rsrc("testWorkbookAppend.xlsx"))

  refCoord_airquality_append <- matrix(c(9, 5, 194, 15), ncol = 2, byrow = TRUE)

  appendNamedRegion(wb.xls, airquality, name = "mtcars")
  expect_equal(refCoord_airquality_append, getReferenceCoordinatesForName(wb.xls, "mtcars"))

  # Verify actual data read back (optional, but good for sanity)
  res_xls_air <- readNamedRegion(wb.xls, name = "mtcars")
  expect_equal(nrow(res_xls_air), 32 + nrow(airquality))

  appendNamedRegion(wb.xlsx, airquality, name = "mtcars")
  expect_equal(refCoord_airquality_append, getReferenceCoordinatesForName(wb.xlsx, "mtcars"))

  res_xlsx_air <- readNamedRegion(wb.xlsx, name = "mtcars")
  expect_equal(nrow(res_xlsx_air), 32 + nrow(airquality))
})
