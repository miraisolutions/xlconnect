# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.isSheetVisible <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xlsx"), create = FALSE)
	
	# Check if sheets are visible (*.xls)
	checkTrue(!isSheetVisible(wb.xls, 2))
	checkTrue(!isSheetVisible(wb.xls, "BBB"))
	checkTrue(isSheetVisible(wb.xls, 1))
	checkTrue(isSheetVisible(wb.xls, "AAA"))
	checkTrue(isSheetVisible(wb.xls, 3))
	checkTrue(isSheetVisible(wb.xls, "CCC"))
	checkTrue(!isSheetVisible(wb.xls, 4))
	checkTrue(!isSheetVisible(wb.xls, "DDD"))
	
	# Check if sheets are visible (*.xls)
	checkTrue(!isSheetVisible(wb.xlsx, 2))
	checkTrue(!isSheetVisible(wb.xlsx, "BBB"))
	checkTrue(isSheetVisible(wb.xlsx, 1))
	checkTrue(isSheetVisible(wb.xlsx, "AAA"))
	checkTrue(isSheetVisible(wb.xlsx, 3))
	checkTrue(isSheetVisible(wb.xlsx, "CCC"))
	checkTrue(!isSheetVisible(wb.xlsx, 4))
	checkTrue(!isSheetVisible(wb.xlsx, "DDD"))
	
}
