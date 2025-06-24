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
        expect_error(getCellStyle(wb.xls, styleName), "IllegalArgumentException")
        expect_error(getCellStyle(wb.xlsx, styleName), "IllegalArgumentException")
        # createCellStyle does not throw an error if the style does not exist yet
        createCellStyle(wb.xls, styleName)
        createCellStyle(wb.xlsx, styleName)
        expect_true(existsCellStyle(wb.xls, styleName))
        expect_true(existsCellStyle(wb.xlsx, styleName))
        expect_error(createCellStyle(wb.xls, styleName), "IllegalArgumentException")
        expect_error(createCellStyle(wb.xlsx, styleName), "IllegalArgumentException")
        # getCellStyle does not throw an error if the style exists
        getCellStyle(wb.xls, styleName)
        getCellStyle(wb.xlsx, styleName)
        expect_false(existsCellStyle(wb.xls, anotherStyleName))
        expect_false(existsCellStyle(wb.xlsx, anotherStyleName))
        # getOrCreateCellStyle does not throw an error
        getOrCreateCellStyle(wb.xls, anotherStyleName)
        getOrCreateCellStyle(wb.xlsx, anotherStyleName)
        expect_true(existsCellStyle(wb.xls, anotherStyleName))
        expect_true(existsCellStyle(wb.xlsx, anotherStyleName))
        # getOrCreateCellStyle does not throw an error
        getOrCreateCellStyle(wb.xls, anotherStyleName)
        getOrCreateCellStyle(wb.xlsx, anotherStyleName)
    }
})

