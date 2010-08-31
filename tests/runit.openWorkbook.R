# Tests around opening an Excel workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.openWorkbook <- function() {
	
	# Check that an exception is thrown when trying to open
	# a non-existent file (*.xls)
	checkException(openWorkbook("resources/fileWhichDoesNotExist.xls"))
	
	# Check that an exception is thrown when trying to open
	# a non-existent file (*.xls)
	checkException(openWorkbook("resources/fileWhichDoesNotExist.xlsx"))
	
	# Check that an instance of the workbook class is returned
	# for an already existing file (*.xls)
	wb <- openWorkbook("resources/testOpenWorkbook.xls")
	checkTrue(is(wb, "workbook"))
	
	# Check that an instance of the workbook class is returned
	# for an already existing file (*.xlsx)
	wb <- openWorkbook("resources/testOpenWorkbook.xlsx")
	checkTrue(is(wb, "workbook"))
	
	# Check that an instance of the workbook class is returned
	# for a file created on-the-fly (*.xls)
	wb <- openWorkbook("resources/fileCreatedOnTheFly.xls", create = TRUE)
	checkTrue(is(wb, "workbook"))
	
	# Check that an instance of the workbook class is returned
	# for a file created on-the-fly (*.xlsx)
	wb <- openWorkbook("resources/fileCreatedOnTheFly.xlsx", create = TRUE)
	checkTrue(is(wb, "workbook"))
}

