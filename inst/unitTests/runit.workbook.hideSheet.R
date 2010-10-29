# 
# This test assumes createSheet & isSheetHidden / isSheetVeryHidden to work correctly
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.hideSheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookHideSheet.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookHideSheet.xlsx"), create = TRUE)
	
	# Sheet names
	visibleSheet <- "VisibleSheet"
	hiddenSheet <- "HiddenSheet"
	veryHiddenSheet <- "VeryHiddenSheet"

	# Create some sheets
	createSheet(wb.xls, visibleSheet)
	createSheet(wb.xlsx, visibleSheet)
	createSheet(wb.xls, hiddenSheet)
	createSheet(wb.xlsx, hiddenSheet)
	createSheet(wb.xls, veryHiddenSheet)
	createSheet(wb.xlsx, veryHiddenSheet)
	
	# Check if sheets are hidden correspondingly (*.xls)
	hideSheet(wb.xls, hiddenSheet)
	checkTrue(isSheetHidden(wb.xls, hiddenSheet))
	hideSheet(wb.xls, veryHiddenSheet, veryHidden = TRUE)
	checkTrue(isSheetVeryHidden(wb.xls, veryHiddenSheet))
	
	# Check if sheets are hidden correspondingly (*.xls)
	hideSheet(wb.xlsx, hiddenSheet)
	checkTrue(isSheetHidden(wb.xlsx, hiddenSheet))
	hideSheet(wb.xlsx, veryHiddenSheet, veryHidden = TRUE)
	checkTrue(isSheetVeryHidden(wb.xlsx, veryHiddenSheet))
	
}
