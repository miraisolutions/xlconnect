# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.unhideSheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xlsx"), create = FALSE)
	
	# Check that unhiding sheets works correctly (*.xls)
	# (assumes 'isSheetHidden' and 'isSheetVeryHidden' work correctly)
	unhideSheet(wb.xls, 2)
	unhideSheet(wb.xls, "DDD")
	checkTrue(!isSheetHidden(wb.xls, 2))
	checkTrue(!isSheetVeryHidden(wb.xls, "DDD"))
	
	# Check that unhiding sheets works correctly (*.xlsx)
	# (assumes 'isSheetHidden' and 'isSheetVeryHidden' work correctly)
	unhideSheet(wb.xlsx, 2)
	unhideSheet(wb.xlsx, "DDD")
	checkTrue(!isSheetHidden(wb.xlsx, 2))
	checkTrue(!isSheetVeryHidden(wb.xlsx, "DDD"))
	
	# Check that attempting to unhide an illegal sheet throws an exception (*.xls)
	checkException(unhideSheet(wb.xls, 58))
	checkException(unhideSheet(wb.xls, "SheetWhichDoesNotExist"))
	
	# Check that attempting to unhide an illegal sheet throws an exception (*.xlsx)
	checkException(unhideSheet(wb.xlsx, 58))
	checkException(unhideSheet(wb.xlsx, "SheetWhichDoesNotExist"))
}
