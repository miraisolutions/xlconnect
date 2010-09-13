# Tests for writing to a worksheet
# * Additional tests to the ones defined in 'runit.workbook.writeAndReadWorksheet.R'
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.writeWorksheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookWriteWorksheet.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookWriteWorksheet.xlsx"), create = TRUE)
	
	# Check that trying to write an object which cannot be converted to a data.frame
	# causes an exception (*.xls)
	createSheet(wb.xls, "test1")
	checkException(writeWorksheet(wb.xls, search, "test1"))
	
	# Check that trying to write an object which cannot be converted to a data.frame
	# causes an exception (*.xlsx)
	createSheet(wb.xlsx, "test1")
	checkException(writeWorksheet(wb.xlsx, search, "test1"))
	
	# Check that attempting to write to a non-existing sheet causes an exception (*.xls)
	checkException(writeWorksheet(wb.xls, mtcars, "sheetDoesNotExist"))
	
	# Check that attempting to write to a non-existing sheet causes an exception (*.xlsx)
	checkException(writeWorksheet(wb.xlsx, mtcars, "sheetDoesNotExist"))
}
