# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.removeSheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook("resources/testWorkbookRemoveSheet.xls", create = FALSE)
	wb.xlsx <- loadWorkbook("resources/testWorkbookRemoveSheet.xlsx", create = FALSE)
	
	# Check that removing a sheet works fine (*.xls)
	# (assumes 'existsSheet' to be working properly)
	removeSheet(wb.xls, "BBB")
	checkTrue(!existsSheet(wb.xls, "BBB"))
	
	# Check that removing a sheet works fine (*.xlsx)
	# (assumes 'existsSheet' to be working properly)
	removeSheet(wb.xlsx, "BBB")
	checkTrue(!existsSheet(wb.xlsx, "BBB"))
	
	# Check that removing a non-existing sheet does not throw an exception (*.xls)
	checkNoException(removeSheet(wb.xls, 35))
	checkNoException(removeSheet(wb.xls, "SheetWhichDoesNotExist"))
	
	# Check that removing a non-existing sheet does not throw an exception (*.xlsx)
	checkNoException(removeSheet(wb.xlsx, 35))
	checkNoException(removeSheet(wb.xlsx, "SheetWhichDoesNotExist"))
}
