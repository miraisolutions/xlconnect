# Tests for writing a named region
# * Additional tests to the ones defined in 'runit.workbook.writeAndReadNamedRegion.R'
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.writeNamedRegion <- function() {
	
	# Create workbooks
	wb.xls <- openWorkbook("resources/testWorkbookWriteNamedRegion.xls", create = TRUE)
	wb.xlsx <- openWorkbook("resources/testWorkbookWriteNamedRegion.xlsx", create = TRUE)
	
	# Check that trying to write an object which cannot be converted to a data.frame
	# causes an exception (*.xls)
	createName(wb.xls, "test1", "Test1!$C$8")
	checkException(writeNamedRegion(wb.xls, search, "test1"))
	
	# Check that trying to write an object which cannot be converted to a data.frame
	# causes an exception (*.xlsx)
	createName(wb.xlsx, "test1", "Test1!$C$8")
	checkException(writeNamedRegion(wb.xlsx, search, "test1"))
	
	# Check that attempting to write to a non-existing name causes an exception (*.xls)
	checkException(writeNamedRegion(wb.xls, mtcars, "nameDoesNotExist"))
	
	# Check that attempting to write to a non-existing name causes an exception (*.xlsx)
	checkException(writeNamedRegion(wb.xlsx, mtcars, "nameDoesNotExist"))
}
