# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.isSheetHidden <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xlsx"), create = FALSE)
	
	# Check if sheets are hidden (*.xls)
	checkTrue(isSheetHidden(wb.xls, 2))
	checkTrue(isSheetHidden(wb.xls, "BBB"))
	checkTrue(!isSheetHidden(wb.xls, 1))
	checkTrue(!isSheetHidden(wb.xls, "AAA"))
	checkTrue(!isSheetHidden(wb.xls, 3))
	checkTrue(!isSheetHidden(wb.xls, "CCC"))
	checkTrue(!isSheetHidden(wb.xls, 4)) # Sheet is actually very hidden!
	checkTrue(!isSheetHidden(wb.xls, "DDD")) # Sheet is actually very hidden!
	
	# Check if sheets are hidden (*.xlsx)
	checkTrue(isSheetHidden(wb.xlsx, 2))
	checkTrue(isSheetHidden(wb.xlsx, "BBB"))
	checkTrue(!isSheetHidden(wb.xlsx, 1))
	checkTrue(!isSheetHidden(wb.xlsx, "AAA"))
	checkTrue(!isSheetHidden(wb.xlsx, 3))
	checkTrue(!isSheetHidden(wb.xlsx, "CCC"))
	checkTrue(!isSheetHidden(wb.xlsx, 4)) # Sheet is actually very hidden!
	checkTrue(!isSheetHidden(wb.xlsx, "DDD")) # Sheet is actually very hidden!
	
	# Check if quering invalid/non-existing sheets
	# causes an exception (*.xls)
	checkException(isSheetHidden(wb.xls, 200))
	checkException(isSheetHidden(wb.xls, "Sheet does not exist"))
	checkException(isSheetHidden(wb.xls, "'Illegal sheet name"))
	
	# Check if quering invalid/non-existing sheets
	# causes an exception (*.xlsx)
	checkException(isSheetHidden(wb.xlsx, 200))
	checkException(isSheetHidden(wb.xlsx, "Sheet does not exist"))
	checkException(isSheetHidden(wb.xlsx, "'Illegal sheet name"))
	
}
