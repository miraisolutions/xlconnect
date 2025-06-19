test_that("test.workbook.testBug181", {
    checkNoException({
        wb <- loadWorkbook("resources/testBug181.xlsx"))
        obj <- readWorksheet(wb, "Sheet1")
        checkTrue(is.data.frame(obj))
    })
})

