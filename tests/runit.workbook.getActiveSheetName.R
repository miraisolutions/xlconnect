# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.getActiveSheetName <- function() {

	# Create workbooks
	wb.xls <- openWorkbook("resources/testWorkbookActiveSheetIndexAndName.xls", create = FALSE)
	wb.xlsx <- openWorkbook("resources/testWorkbookActiveSheetIndexAndName.xlsx", create = FALSE)
	
	# Check that the active sheet name is 'Fifth Sheet' (*.xls)
	checkTrue(getActiveSheetName(wb.xls) == "Fifth Sheet")
	
	# Check that the active sheet name is 'Fifth Sheet' (*.xlsx)
	checkTrue(getActiveSheetName(wb.xlsx) == "Fifth Sheet")
	
}
