# Tests around opening an Excel workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.loadWorkbook <- function() {
	
	# Check that an exception is thrown when trying to open
	# a non-existent file (*.xls)
	checkException(loadWorkbook("resources/fileWhichDoesNotExist.xls"))
	
	# Check that an exception is thrown when trying to open
	# a non-existent file (*.xls)
	checkException(loadWorkbook("resources/fileWhichDoesNotExist.xlsx"))
	
	# Check that an instance of the workbook class is returned
	# for an already existing file (*.xls)
	wb <- loadWorkbook("resources/testLoadWorkbook.xls")
	checkTrue(is(wb, "workbook"))
	
	# Check that an instance of the workbook class is returned
	# for an already existing file (*.xlsx)
	wb <- loadWorkbook("resources/testLoadWorkbook.xlsx")
	checkTrue(is(wb, "workbook"))
	
	# Check that an instance of the workbook class is returned
	# for a file created on-the-fly (*.xls)
	wb <- loadWorkbook("resources/fileCreatedOnTheFly.xls", create = TRUE)
	checkTrue(is(wb, "workbook"))
	
	# Check that an instance of the workbook class is returned
	# for a file created on-the-fly (*.xlsx)
	wb <- loadWorkbook("resources/fileCreatedOnTheFly.xlsx", create = TRUE)
	checkTrue(is(wb, "workbook"))
}

