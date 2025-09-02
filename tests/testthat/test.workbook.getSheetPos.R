test_that("getSheetPos returns correct positions for existing sheets in XLS", {
  wb.xls <- loadWorkbook("resources/testWorkbookGetSheetPos.xls", create = TRUE)
  createSheet(wb.xls, c("Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4"))
  # Check positions (*.xls)
  expected <- c(`Sheet 3` = 3, `Sheet 2` = 2, `Sheet 4` = 4, `Sheet 1` = 1)
  expect_equal(expected, getSheetPos(wb.xls, c("Sheet 3", "Sheet 2", "Sheet 4", "Sheet 1")))
})

test_that("getSheetPos returns correct positions for existing sheets in XLSX", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookGetSheetPos.xlsx", create = TRUE)
  createSheet(wb.xlsx, c("Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4"))
  # Check positions (*.xlsx)
  expected <- c(`Sheet 3` = 3, `Sheet 2` = 2, `Sheet 4` = 4, `Sheet 1` = 1)
  expect_equal(expected, getSheetPos(wb.xlsx, c("Sheet 3", "Sheet 2", "Sheet 4", "Sheet 1")))
})

test_that("getSheetPos returns 0 index for non-existing worksheets in XLS", {
  wb.xls <- loadWorkbook("resources/testWorkbookGetSheetPos.xls", create = TRUE)
  # Check that querying a non-existing worksheet results in a 0 index (*.xls)
  expect_equal(c(NotExisting = 0), getSheetPos(wb.xls, "NotExisting"))
  expect_equal(0, as.vector(getSheetPos(wb.xls, "%#?%+?[-")))
})

test_that("getSheetPos returns 0 index for non-existing worksheets in XLSX", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookGetSheetPos.xlsx", create = TRUE)
  # Check that querying a non-existing worksheet results in a 0 index (*.xlsx)
  expect_equal(c(NotExisting = 0), getSheetPos(wb.xlsx, "NotExisting"))
  expect_equal(0, as.vector(getSheetPos(wb.xlsx, "%#?%+?[-")))
})
