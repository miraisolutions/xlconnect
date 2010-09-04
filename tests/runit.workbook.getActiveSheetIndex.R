# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.getActiveSheetIndex <- function() {
	
	# Create workbooks
	wb.xls <- openWorkbook("resources/testWorkbookActiveSheetIndexAndName.xls", create = FALSE)
	wb.xlsx <- openWorkbook("resources/testWorkbookActiveSheetIndexAndName.xlsx", create = FALSE)
	
	# Check that the active sheet index is 5 (*.xls)
	checkTrue(getActiveSheetIndex(wb.xls) == 5)
	
	# Check that the active sheet index is 5 (*.xlsx)
	checkTrue(getActiveSheetIndex(wb.xlsx) == 5)
	
}
