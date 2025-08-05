test_that("with.workbook correctly loads data from an XLS file", {
  wb.xls <- loadWorkbook(rsrc("testWithWorkbook.xls"), create = FALSE)
  with(
    wb.xls,
    {
      expect_true(all(dim(AA) == c(8, 3)), info = "XLS: Check dimensions of sheet AA")
      expect_true(all(dim(BB) == c(5, 5)), info = "XLS: Check dimensions of sheet BB")
    },
    header = FALSE
  )
})

test_that("with.workbook correctly loads data from an XLSX file", {
  wb.xlsx <- loadWorkbook(rsrc("testWithWorkbook.xlsx"), create = FALSE)
  with(
    wb.xlsx,
    {
      expect_true(all(dim(AA) == c(8, 3)), info = "XLSX: Check dimensions of sheet AA")
      expect_true(all(dim(BB) == c(5, 5)), info = "XLSX: Check dimensions of sheet BB")
    },
    header = FALSE
  )
})
