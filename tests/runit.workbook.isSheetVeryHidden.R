# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.isSheetVeryHidden <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook("resources/testWorkbookHiddenSheets.xls", create = FALSE)
	wb.xlsx <- loadWorkbook("resources/testWorkbookHiddenSheets.xlsx", create = FALSE)
	
	# Check if sheets are hidden (*.xls)
	checkTrue(isSheetVeryHidden(wb.xls, 4))
	checkTrue(isSheetVeryHidden(wb.xls, "DDD"))
	checkTrue(!isSheetVeryHidden(wb.xls, 1))
	checkTrue(!isSheetVeryHidden(wb.xls, "AAA"))
	checkTrue(!isSheetVeryHidden(wb.xls, 2)) # Sheet is actually hidden only!
	checkTrue(!isSheetVeryHidden(wb.xls, "BBB")) # Sheet is actually hidden only!
	checkTrue(!isSheetVeryHidden(wb.xls, 3))
	checkTrue(!isSheetVeryHidden(wb.xls, "CCC"))
	
	# Check if sheets are hidden (*.xlsx)
	checkTrue(isSheetVeryHidden(wb.xlsx, 4))
	checkTrue(isSheetVeryHidden(wb.xlsx, "DDD"))
	checkTrue(!isSheetVeryHidden(wb.xlsx, 1))
	checkTrue(!isSheetVeryHidden(wb.xlsx, "AAA"))
	checkTrue(!isSheetVeryHidden(wb.xlsx, 2)) # Sheet is actually hidden only!
	checkTrue(!isSheetVeryHidden(wb.xlsx, "BBB")) # Sheet is actually hidden only!
	checkTrue(!isSheetVeryHidden(wb.xlsx, 3))
	checkTrue(!isSheetVeryHidden(wb.xlsx, "CCC"))
	
	# Check if quering invalid/non-existing sheets
	# causes an exception (*.xls)
	checkException(isSheetVeryHidden(wb.xls, 200))
	checkException(isSheetVeryHidden(wb.xls, "Sheet does not exist"))
	checkException(isSheetVeryHidden(wb.xls, "'Illegal sheet name"))
	
	# Check if quering invalid/non-existing sheets
	# causes an exception (*.xlsx)
	checkException(isSheetVeryHidden(wb.xlsx, 200))
	checkException(isSheetVeryHidden(wb.xlsx, "Sheet does not exist"))
	checkException(isSheetVeryHidden(wb.xlsx, "'Illegal sheet name"))
	
}
