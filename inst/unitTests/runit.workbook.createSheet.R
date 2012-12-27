#############################################################################
#
# XLConnect
# Copyright (C) 2010-2013 Mirai Solutions GmbH
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
#############################################################################

#############################################################################
#
# Tests around creating Excel worksheets
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.createSheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/createSheet.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/createSheet.xlsx"), create = TRUE)
	
	# Check that an exception is thrown when trying to create
	# a worksheet with a single quote at the beginning (*.xls)
	checkException(createSheet(wb.xls, "'Invalid Sheet Name"))
	
	# Check that an exception is thrown when trying to create
	# a worksheet with a single quote at the beginning (*.xlsx)
	checkException(createSheet(wb.xlsx, "'Invalid Sheet Name"))
	
	# Check that an exception is thrown when trying to create
	# a worksheet with a single quote at the end (*.xls)
	checkException(createSheet(wb.xls, "Invalid Sheet Name'"))
	
	# Check that an exception is thrown when trying to create
	# a worksheet with a single quote at the end (*.xlsx)
	checkException(createSheet(wb.xlsx, "Invalid Sheet Name'"))
	
	# Check that an exception is thrown when trying to create
	# a worksheet with a very long name (> 31 characters) (*.xls)
	checkException(createSheet(wb.xls, "A very very very very very very very very long name"))
	
	# Check that an exception is thrown when trying to create
	# a worksheet with a very long name (> 31 characters) (*.xlsx)
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

