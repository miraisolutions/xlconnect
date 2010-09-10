# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.setActiveSheet <- function() {
	
	# Create workbooks
	wb.xls <- openWorkbook("resources/testWorkbookSetActiveSheet.xls", create = FALSE)
	wb.xlsx <- openWorkbook("resources/testWorkbookSetActiveSheet.xlsx", create = FALSE)
	
	# Check that setting the active sheet works ok (*.xls)
	# (assumes that 'getActiveSheetIndex' works fine)
	setActiveSheet(wb.xls, 1)
	checkTrue(getActiveSheetIndex(wb.xls) == 1)
	setActiveSheet(wb.xls, 3)
	checkTrue(getActiveSheetIndex(wb.xls) == 3)
	setActiveSheet(wb.xls, "Sheet2")
	checkTrue(getActiveSheetIndex(wb.xls) == 2)
	
	# Check that setting the active sheet works ok (*.xlsx)
	# (assumes that 'getActiveSheetIndex' works fine)
	setActiveSheet(wb.xlsx, 1)
	checkTrue(getActiveSheetIndex(wb.xlsx) == 1)
	setActiveSheet(wb.xlsx, 3)
	checkTrue(getActiveSheetIndex(wb.xlsx) == 3)
	setActiveSheet(wb.xlsx, "Sheet2")
	checkTrue(getActiveSheetIndex(wb.xlsx) == 2)
	
	# Check that setting an illegal active sheet throws an exception (*.xls)
	checkException(setActiveSheet(wb.xls, 19))
	checkException(setActiveSheet(wb.xls, "SheetWhichDoesNotExist"))
	
	# Check that setting an illegal active sheet throws an exception (*.xlsx)
	checkException(setActiveSheet(wb.xlsx, 19))
	checkException(setActiveSheet(wb.xlsx, "SheetWhichDoesNotExist"))
}
