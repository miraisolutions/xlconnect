context("with.workbook functionality")

test_that("with.workbook correctly loads data from specified sheets", {
    wb.xls <- loadWorkbook(rsrc("testWithWorkbook.xls"),
        create = FALSE
    )
    wb.xlsx <- loadWorkbook(rsrc("testWithWorkbook.xlsx"),
        create = FALSE
    )

    # Test with XLS workbook
    with(wb.xls,
        {
            expect_true(all(dim(AA) == c(8, 3)), info = "XLS: Check dimensions of sheet AA")
            expect_true(all(dim(BB) == c(5, 5)), info = "XLS: Check dimensions of sheet BB")
        },
        header = FALSE # Assuming this applies to how with.workbook reads the sheets
    )

    # Test with XLSX workbook
    with(wb.xlsx,
        {
            expect_true(all(dim(AA) == c(8, 3)), info = "XLSX: Check dimensions of sheet AA")
            expect_true(all(dim(BB) == c(5, 5)), info = "XLSX: Check dimensions of sheet BB")
        },
        header = FALSE # Assuming this applies to how with.workbook reads the sheets
    )
})
