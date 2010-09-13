# Tests around opening an Excel workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.loadWorkbook <- function() {
	
	# Check that an exception is thrown when trying to open
	# a non-existent file (*.xls)
	checkException(loadWorkbook(rsrc("resources/fileWhichDoesNotExist.xls")))
	
	# Check that an exception is thrown when trying to open
	# a non-existent file (*.xls)
	checkException(loadWorkbook(rsrc("resources/fileWhichDoesNotExist.xlsx")))
	
	# Check that an instance of the workbook class is returned
	# for an already existing file (*.xls)
	wb <- loadWorkbook(rsrc("resources/testLoadWorkbook.xls"))
	checkTrue(is(wb, "workbook"))
	
	# Check that an instance of the workbook class is returned
	# for an already existing file (*.xlsx)
	wb <- loadWorkbook(rsrc("resources/testLoadWorkbook.xlsx"))
	checkTrue(is(wb, "workbook"))
	
	# Check that an instance of the workbook class is returned
	# for a file created on-the-fly (*.xls)
	wb <- loadWorkbook(rsrc("resources/fileCreatedOnTheFly.xls"), create = TRUE)
	checkTrue(is(wb, "workbook"))
	
	# Check that an instance of the workbook class is returned
	# for a file created on-the-fly (*.xlsx)
	wb <- loadWorkbook(rsrc("resources/fileCreatedOnTheFly.xlsx"), create = TRUE)
	checkTrue(is(wb, "workbook"))
}

