# Tests for creating a worksheet
# 
# Author: Martin Studer, Mirai Solutions GmbH
###############################################################################

test.workbook.createSheet <- function() {
	
	# Create workbooks
	wb.xls <- openWorkbook("resources/createSheet.xls", create = TRUE)
	wb.xlsx <- openWorkbook("resources/createSheet.xlsx", create = TRUE)
	
	# Check that an exception is thrown when trying to create
	# a worksheet with a single quote at the beginning (*.xls)
	checkException(createSheet(wb.xls, "'Invalid Sheet Name"))
	
	# Check that an exception is thrown when trying to create
	# a worksheet with a single quote at the beginning (*.xlsx)
	# TODO: Check again, this currently does not produce an error - POI bug?
	# checkException(createSheet(wb.xlsx, "'Invalid Sheet Name"))
	
	# Check that an exception is thrown when trying to create
	# a worksheet with a single quote at the end (*.xls)
	checkException(createSheet(wb.xls, "Invalid Sheet Name'"))
	
	# Check that an exception is thrown when trying to create
	# a worksheet with a single quote at the end (*.xlsx)
	# TODO: Check again, this currently does not produce an error - POI bug?
	# checkException(createSheet(wb.xlsx, "Invalid Sheet Name'"))
	
	# Check that an exception is thrown when trying to create
	# a worksheet with a very long name (> 30 characters) (*.xls)
	# TODO: Check again, this currently does not produce an exception - rather truncates: probably POI inconsistency
	# checkException(createSheet(wb.xls, "A very very very very very very very very long name"))
	
	# Check that an exception is thrown when trying to create
	# a worksheet with a very long name (> 30 characters) (*.xlsx)
	checkException(createSheet(wb.xlsx, "A very very very very very very very very long name"))
	
	
	sheetName <- "My Sheet"

	# Check if creating a legal worksheet is working properly (*.xls)
	# (assumes method existsSheet working properly)
	try(createSheet(wb.xls, sheetName))
	checkTrue(existsSheet(wb.xls, sheetName))
	
	# Check if creating a legal worksheet is working properly (*.xlsx)
	# (assumes method existsSheet working properly)
	try(createSheet(wb.xlsx, sheetName))
	checkTrue(existsSheet(wb.xlsx, sheetName))	
	
	# Trying to create an already existing sheet should not cause 
	# any issues (just skips) (*.xls)
	try(createSheet(wb.xls, sheetName))
	checkTrue(existsSheet(wb.xls, sheetName))
	
	# Trying to create an already existing sheet should not cause 
	# any issues (just skips) (*.xlsx)
	try(createSheet(wb.xlsx, sheetName))
	checkTrue(existsSheet(wb.xlsx, sheetName))
	
}

