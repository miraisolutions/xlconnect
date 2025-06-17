test_that("test.with.workbook", {
    wb.xls <- loadWorkbook(rsrc("resources/testWithWorkbook.xls"), 
        create = FALSE)
    wb.xlsx <- loadWorkbook(rsrc("resources/testWithWorkbook.xlsx"), 
        create = FALSE)
    with(wb.xls, {
        checkTrue(all(dim(AA) == c(8, 3)))
        checkTrue(all(dim(BB) == c(5, 5)))
    }, header = FALSE)
    with(wb.xlsx, {
        checkTrue(all(dim(AA) == c(8, 3)))
        checkTrue(all(dim(BB) == c(5, 5)))
    }, header = FALSE)
})

