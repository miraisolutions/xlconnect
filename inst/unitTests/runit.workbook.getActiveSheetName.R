# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.getActiveSheetName <- function() {

	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookActiveSheetIndexAndName.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookActiveSheetIndexAndName.xlsx"), create = FALSE)
	
	# Check that the active sheet name is 'Fifth Sheet' (*.xls)
	checkTrue(getActiveSheetName(wb.xls) == "Fifth Sheet")
	
	# Check that the active sheet name is 'Fifth Sheet' (*.xlsx)
	checkTrue(getActiveSheetName(wb.xlsx) == "Fifth Sheet")
	
}
