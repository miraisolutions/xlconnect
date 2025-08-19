test_that("unhideSheet works correctly for xls", {
  wb.xls <- loadWorkbook(rsrc("testWorkbookHiddenSheets.xls"), create = FALSE)
  unhideSheet(wb.xls, 2)
  unhideSheet(wb.xls, "DDD")
  expect_false(isSheetHidden(wb.xls, 2))
  expect_false(isSheetVeryHidden(wb.xls, "DDD"))
})

test_that("unhideSheet works correctly for xlsx", {
  wb.xlsx <- loadWorkbook(rsrc("testWorkbookHiddenSheets.xlsx"), create = FALSE)
  unhideSheet(wb.xlsx, 2)
  unhideSheet(wb.xlsx, "DDD")
  expect_false(isSheetHidden(wb.xlsx, 2))
  expect_false(isSheetVeryHidden(wb.xlsx, "DDD"))
})

test_that("unhideSheet throws an error for non-existent sheets in xls", {
  wb.xls <- loadWorkbook(rsrc("testWorkbookHiddenSheets.xls"), create = FALSE)
  expect_error(unhideSheet(wb.xls, 58), "IllegalArgumentException")
  expect_error(unhideSheet(wb.xls, "SheetWhichDoesNotExist"), "IllegalArgumentException")
})

test_that("unhideSheet throws an error for non-existent sheets in xlsx", {
  wb.xlsx <- loadWorkbook(rsrc("testWorkbookHiddenSheets.xlsx"), create = FALSE)
  expect_error(unhideSheet(wb.xlsx, 58), "IllegalArgumentException")
  expect_error(unhideSheet(wb.xlsx, "SheetWhichDoesNotExist"), "IllegalArgumentException")
})
