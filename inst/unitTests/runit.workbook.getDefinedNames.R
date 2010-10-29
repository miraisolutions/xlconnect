# TODO: Add comment
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.getDefinedNames <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookDefinedNames.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookDefinedNames.xlsx"), create = FALSE)
	
	# Names defined in workbooks
	expectedNamesValidOnly <- c("FirstName", "SecondName", "FourthName", "FifthName")
	expectedNamesAll <- c("FirstName", "SecondName", "ThirdName", "FourthName", "FifthName")
	
	# Check that all and only the expected names exist (*.xls)
	definedNames <- getDefinedNames(wb.xls, validOnly = TRUE)
	checkTrue(length(setdiff(expectedNamesValidOnly, definedNames)) == 0 && length(setdiff(definedNames, expectedNamesValidOnly)) == 0)
	definedNames <- getDefinedNames(wb.xls, validOnly = FALSE)
	checkTrue(length(setdiff(expectedNamesAll, definedNames)) == 0 && length(setdiff(definedNames, expectedNamesAll)) == 0)
	
	# Check that all and only the expected names exist (*.xlsx)
	definedNames <- getDefinedNames(wb.xlsx, validOnly = TRUE)
	checkTrue(length(setdiff(expectedNamesValidOnly, definedNames)) == 0 && length(setdiff(definedNames, expectedNamesValidOnly)) == 0)
	definedNames <- getDefinedNames(wb.xlsx, validOnly = FALSE)
	checkTrue(length(setdiff(expectedNamesAll, definedNames)) == 0 && length(setdiff(definedNames, expectedNamesAll)) == 0)
	
}
