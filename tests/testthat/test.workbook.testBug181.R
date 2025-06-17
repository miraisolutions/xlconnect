test_that("test.workbook.testBug181", {
    checkNoException({
        wb <- loadWorkbook(rsrc("resources/testBug181.xlsx"))
        obj <- readWorksheet(wb, "Sheet1")
        checkTrue(is.data.frame(obj))
    })
})

