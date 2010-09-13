# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.getSheets <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookSheets.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookSheets.xlsx"), create = FALSE)
	
	# Sheets defined in workbooks
	expectedSheets <- c("A1", "B 2", "$$", "=", "@}", "11. Oct.", "\"quote\"", "+0")
	
	# Check that all and only the expected sheets exist (*.xls)
	definedSheets <- getSheets(wb.xls)
	checkTrue(length(setdiff(expectedSheets, definedSheets)) == 0 && length(setdiff(definedSheets, expectedSheets)) == 0)
	
	# Check that all and only the expected sheets exist (*.xlsx)
	definedSheets <- getSheets(wb.xlsx)
	checkTrue(length(setdiff(expectedSheets, definedSheets)) == 0 && length(setdiff(definedSheets, expectedSheets)) == 0)
	
}
