# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.removeName <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook("resources/testWorkbookRemoveName.xls", create = FALSE)
	wb.xlsx <- loadWorkbook("resources/testWorkbookRemoveName.xlsx", create = FALSE)
	
	# Check that when removing a name from a worksheet it does not exist anymore (*.xls)
	# (assumes 'existsName' is working correctly)
	removeName(wb.xls, "AA")
	checkTrue(!existsName(wb.xls, "AA"))
	
	# Check that when removing a name from a worksheet it does not exist anymore (*.xlsx)
	# (assumes 'existsName' is working correctly)
	removeName(wb.xlsx, "AA")
	checkTrue(!existsName(wb.xlsx, "AA"))
	
	# Check that attempting to remove a non-existing name does not throw an exception (*.xls)
	checkNoException(removeName(wb.xls, "NameWhichDoesNotExist"))
	
	# Check that attempting to remove a non-existing name does not throw an exception (*.xlsx)
	checkNoException(removeName(wb.xlsx, "NameWhichDoesNotExist"))
}
