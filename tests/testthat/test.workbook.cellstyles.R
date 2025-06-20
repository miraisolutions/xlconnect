test_that("test.workbook.cellstyles", {
    if (getOption("FULL.TEST.SUITE")) {
        file.xls <- "resources/cellstyles.xls"
        file.xlsx <- "resources/cellstyles.xlsx"
        file.remove(file.xls)
        file.remove(file.xlsx)
        wb.xls <- loadWorkbook(file.xls, create = TRUE)
        wb.xlsx <- loadWorkbook(file.xlsx, create = TRUE)
        styleName <- "MyStyle"
        anotherStyleName <- "MyOtherStyle"
        expect_false(existsCellStyle(wb.xls, styleName))
        expect_false(existsCellStyle(wb.xlsx, styleName))
        expect_error(getCellStyle(wb.xls, styleName), NA))
        expect_error(getCellStyle(wb.xlsx, styleName), NA))
        expect_error(createCellStyle(wb.xls, styleName), NA)
        expect_error(createCellStyle(wb.xlsx, styleName), NA)
        expect_true(existsCellStyle(wb.xls, styleName))
        expect_true(existsCellStyle(wb.xlsx, styleName))
        expect_error(createCellStyle(wb.xls, styleName), NA)
        expect_error(createCellStyle(wb.xlsx, styleName), NA)
        expect_error(getCellStyle(wb.xls, styleName), NA))
        expect_error(getCellStyle(wb.xlsx, styleName), NA))
        expect_false(existsCellStyle(wb.xls, anotherStyleName))
        expect_false(existsCellStyle(wb.xlsx, anotherStyleName))
        expect_error(getOrCreateCellStyle(wb.xls, anotherStyleName), NA))
        expect_error(getOrCreateCellStyle(wb.xlsx, anotherStyleName), NA))
        expect_true(existsCellStyle(wb.xls, anotherStyleName))
        expect_true(existsCellStyle(wb.xlsx, anotherStyleName))
        expect_error(getOrCreateCellStyle(wb.xls, anotherStyleName), NA))
        expect_error(getOrCreateCellStyle(wb.xlsx, anotherStyleName), NA))
    }
})

