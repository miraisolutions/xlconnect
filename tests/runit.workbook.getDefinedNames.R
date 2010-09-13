# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.getDefinedNames <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook("resources/testWorkbookDefinedNames.xls", create = FALSE)
	wb.xlsx <- loadWorkbook("resources/testWorkbookDefinedNames.xlsx", create = FALSE)
	
	# Names defined in workbooks
	expectedNames <- c("FirstName", "SecondName", "ThirdName", "FourthName", "FifthName")
	
	# Check that all and only the expected names exist (*.xls)
	definedNames <- getDefinedNames(wb.xls)
	checkTrue(length(setdiff(expectedNames, definedNames)) == 0 && length(setdiff(definedNames, expectedNames)) == 0)
	
	# Check that all and only the expected names exist (*.xlsx)
	definedNames <- getDefinedNames(wb.xlsx)
	checkTrue(length(setdiff(expectedNames, definedNames)) == 0 && length(setdiff(definedNames, expectedNames)) == 0)
	
}
