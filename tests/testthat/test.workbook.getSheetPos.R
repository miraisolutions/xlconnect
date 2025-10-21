test_that("getSheetPos returns correct positions for existing sheets in XLS", {
  wb.xls <- loadWorkbook("resources/testWorkbookGetSheetPos.xls", create = TRUE)
  createSheet(wb.xls, c("Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4"))
  # Check positions (*.xls)
  expected <- c(`Sheet 3` = 3, `Sheet 2` = 2, `Sheet 4` = 4, `Sheet 1` = 1)
  expect_equal(getSheetPos(wb.xls, c("Sheet 3", "Sheet 2", "Sheet 4", "Sheet 1")), expected)
})

test_that("getSheetPos returns correct positions for existing sheets in XLSX", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookGetSheetPos.xlsx", create = TRUE)
  createSheet(wb.xlsx, c("Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4"))
  # Check positions (*.xlsx)
  expected <- c(`Sheet 3` = 3, `Sheet 2` = 2, `Sheet 4` = 4, `Sheet 1` = 1)
  expect_equal(getSheetPos(wb.xlsx, c("Sheet 3", "Sheet 2", "Sheet 4", "Sheet 1")), expected)
})

test_that("getSheetPos returns 0 index for non-existing worksheets in XLS", {
  wb.xls <- loadWorkbook("resources/testWorkbookGetSheetPos.xls", create = TRUE)
  # Check that querying a non-existing worksheet results in a 0 index (*.xls)
  expect_equal(getSheetPos(wb.xls, "NotExisting"), c(NotExisting = 0))
  expect_equal(as.vector(getSheetPos(wb.xls, "%#?%+?[-")), 0)
})

test_that("getSheetPos returns 0 index for non-existing worksheets in XLSX", {
  wb.xlsx <- loadWorkbook("resources/testWorkbookGetSheetPos.xlsx", create = TRUE)
  # Check that querying a non-existing worksheet results in a 0 index (*.xlsx)
  expect_equal(getSheetPos(wb.xlsx, "NotExisting"), c(NotExisting = 0))
  expect_equal(as.vector(getSheetPos(wb.xlsx, "%#?%+?[-")), 0)
})
