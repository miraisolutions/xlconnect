# Helper function to avoid code duplication
test_error_warn <- function(wb, region, expected_df, warning_msg) {
  expect_warning(res <- try(readNamedRegion(wb, name = region)), warning_msg)
  expect_false(is(res, "try-error"))
  # if (is.null(attr(res, "worksheetScope"))) {
  # attr(expected_df, "worksheetScope") <- NULL
  # } else {
  # }
  expect_equal(expected_df, res)
}

add_expected_attr <- function(df) {
  attr(df, "worksheetScope") <- ""
  df
}

test_that("onErrorCell with XLC$ERROR.WARN works for xls", {
  wb.xls <- loadWorkbook(rsrc("testWorkbookErrorCell.xls"), create = FALSE)

  # Check that reading error cells with the warning flag set does not cause any issues (*.xls)
  onErrorCell(wb.xls, XLC$ERROR.WARN)

  # Test regions
  test_error_warn(
    wb.xls,
    "AA",
    add_expected_attr(data.frame(A = c("aa", "bb", "cc", NA, "ee", "ff"), stringsAsFactors = FALSE)),
    "Error detected in cell"
  )
  test_error_warn(
    wb.xls,
    "BB",
    add_expected_attr(data.frame(B = c(4.3, NA, -2.5, 1.6, NA, 9.7), stringsAsFactors = FALSE)),
    "Error detected in cell"
  )
  # Test with XLConnect.setCustomAttributes = FALSE
  withr::local_options(XLConnect.setCustomAttributes = FALSE)
  data_frame_without_worksheet_scope = data.frame(C = c(-53.2, NA, 34.1, -37.89, 0, 1.6), stringsAsFactors = FALSE)
  test_error_warn(wb.xls, "CC", data_frame_without_worksheet_scope, "Error detected in cell")
  # Reset to TRUE for remaining tests
  withr::local_options(XLConnect.setCustomAttributes = TRUE)
  test_error_warn(
    wb.xls,
    "DD",
    add_expected_attr(data.frame(D = c(8.2, 2, 1, -0.5, NA, 3.1), stringsAsFactors = FALSE)),
    "Error detected in cell"
  )
  test_error_warn(
    wb.xls,
    "EE",
    add_expected_attr(data.frame(E = c("zz", "yy", NA, "ww", "vv", "uu"), stringsAsFactors = FALSE)),
    "Error when trying to evaluate cell"
  )
})

test_that("onErrorCell with XLC$ERROR.WARN works for xlsx", {
  wb.xlsx <- loadWorkbook(rsrc("testWorkbookErrorCell.xlsx"), create = FALSE)

  # Check that reading error cells with the warning flag set does not cause any issues (*.xlsx)
  onErrorCell(wb.xlsx, XLC$ERROR.WARN)

  # Test regions
  test_error_warn(
    wb.xlsx,
    "AA",
    add_expected_attr(data.frame(A = c("aa", "bb", "cc", NA, "ee", "ff"), stringsAsFactors = FALSE)),
    "Error when trying to evaluate cell"
  )
  test_error_warn(
    wb.xlsx,
    "BB",
    add_expected_attr(data.frame(B = c(4.3, NA, -2.5, 1.6, NA, 9.7), stringsAsFactors = FALSE)),
    "Error detected in cell"
  )
  test_error_warn(
    wb.xlsx,
    "CC",
    add_expected_attr(data.frame(C = c(-53.2, NA, 34.1, -37.89, 0, 1.6), stringsAsFactors = FALSE)),
    "Error detected in cell"
  )
  test_error_warn(
    wb.xlsx,
    "DD",
    add_expected_attr(data.frame(D = c(8.2, 2, 1, -0.5, NA, 3.1), stringsAsFactors = FALSE)),
    "Error detected in cell"
  )
  test_error_warn(
    wb.xlsx,
    "EE",
    add_expected_attr(data.frame(E = c("zz", "yy", NA, "ww", "vv", "uu"), stringsAsFactors = FALSE)),
    "Error when trying to evaluate cell"
  )
})

test_that("onErrorCell with XLC$ERROR.STOP creates an error for xls", {
  wb.xls <- loadWorkbook(rsrc("testWorkbookErrorCell.xls"), create = FALSE)

  # Check that reading error cells with the stop flag set causes an exception (*.xls)
  onErrorCell(wb.xls, XLC$ERROR.STOP)

  expect_true(is(try(readNamedRegion(wb.xls, name = "AA")), "try-error"))
  expect_true(is(try(readNamedRegion(wb.xls, name = "BB")), "try-error"))
  expect_true(is(try(readNamedRegion(wb.xls, name = "CC")), "try-error"))
  expect_true(is(try(readNamedRegion(wb.xls, name = "DD")), "try-error"))
  expect_true(is(try(readNamedRegion(wb.xls, name = "EE")), "try-error"))
})

test_that("onErrorCell with XLC$ERROR.STOP creates an error for xlsx", {
  wb.xlsx <- loadWorkbook(rsrc("testWorkbookErrorCell.xlsx"), create = FALSE)

  # Check that reading error cells with the stop flag set causes an exception (*.xlsx)
  onErrorCell(wb.xlsx, XLC$ERROR.STOP)

  expect_true(is(try(readNamedRegion(wb.xlsx, name = "AA")), "try-error"))
  expect_true(is(try(readNamedRegion(wb.xlsx, name = "BB")), "try-error"))
  expect_true(is(try(readNamedRegion(wb.xlsx, name = "CC")), "try-error"))
  expect_true(is(try(readNamedRegion(wb.xlsx, name = "DD")), "try-error"))
  expect_true(is(try(readNamedRegion(wb.xlsx, name = "EE")), "try-error"))
})
