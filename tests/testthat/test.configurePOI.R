test_that("exception is thrown if we limit the number of files to just 1", {
  # Check that an exception is thrown if we limit the number of files to just 1
  configurePOI(zip_max_files = 1L)
  expect_error(loadWorkbook("resources/testLoadWorkbook.xlsx"))

  # Reset settings after test
  configurePOI()
})

test_that("zip bomb detection works with high inflate ratio", {
  # Check zip bomb detection with high inflate ratio
  configurePOI(zip_min_inflate_ratio = 0.99)
  expect_error(readWorksheetFromFile("resources/testZipBomb.xlsx", sheet = 1))

  # Reset settings after test
  configurePOI()
})

test_that("exception is thrown if we limit the max zip entry size to just 1 byte", {
  # Check that an exception is thrown if we limit the max zip entry size to just
  # 1 byte
  configurePOI(zip_max_entry_size = 1L)
  expect_error(readWorksheetFromFile("resources/testWorkbookReadWorksheet.xlsx", sheet = 1))

  # Reset settings after test
  configurePOI()
})

test_that("zip entries are stored in temp files when zip_entry_threshold_bytes is 0", {
  # Check storing of zip entries in temp files
  configurePOI(zip_entry_threshold_bytes = 0L)
  expect_true(is.data.frame(readWorksheetFromFile("resources/testWorkbookReadWorksheet.xlsx", sheet = 1)))

  # Reset settings after test
  configurePOI()
})
