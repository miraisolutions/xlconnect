test_that("unhiding sheets works correctly for xls format (assumes 'isSheetHidden' and 'isSheetVeryHidden' work correctly)", {
  wb.xls <- loadWorkbook(rsrc("testWorkbookHiddenSheets.xls"), create = FALSE)
  unhideSheet(wb.xls, 2)
  unhideSheet(wb.xls, "DDD")
  expect_false(isSheetHidden(wb.xls, 2))
  expect_false(isSheetVeryHidden(wb.xls, "DDD"))
})

test_that("unhiding sheets works correctly for xlsx format (assumes 'isSheetHidden' and 'isSheetVeryHidden' work correctly)", {
  wb.xlsx <- loadWorkbook(rsrc("testWorkbookHiddenSheets.xlsx"), create = FALSE)
  unhideSheet(wb.xlsx, 2)
  unhideSheet(wb.xlsx, "DDD")
  expect_false(isSheetHidden(wb.xlsx, 2))
  expect_false(isSheetVeryHidden(wb.xlsx, "DDD"))
})

test_that("attempting to unhide an illegal sheet throws an exception for xls format", {
  wb.xls <- loadWorkbook(rsrc("testWorkbookHiddenSheets.xls"), create = FALSE)
  expect_error(unhideSheet(wb.xls, 58), "IllegalArgumentException")
  expect_error(unhideSheet(wb.xls, "SheetWhichDoesNotExist"), "IllegalArgumentException")
})

test_that("attempting to unhide an illegal sheet throws an exception for xlsx format", {
  wb.xlsx <- loadWorkbook(rsrc("testWorkbookHiddenSheets.xlsx"), create = FALSE)
  expect_error(unhideSheet(wb.xlsx, 58), "IllegalArgumentException")
  expect_error(unhideSheet(wb.xlsx, "SheetWhichDoesNotExist"), "IllegalArgumentException")
})
