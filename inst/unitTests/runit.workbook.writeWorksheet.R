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
# Tests around writing to an Excel worksheet
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.writeWorksheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookWriteWorksheet.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookWriteWorksheet.xlsx"), create = TRUE)
	
	# Check that trying to write an object which cannot be converted to a data.frame
	# causes an exception (*.xls)
	createSheet(wb.xls, "test1")
	checkException(writeWorksheet(wb.xls, search, "test1"))
	
	# Check that trying to write an object which cannot be converted to a data.frame
	# causes an exception (*.xlsx)
	createSheet(wb.xlsx, "test1")
	checkException(writeWorksheet(wb.xlsx, search, "test1"))
	
	# Check that attempting to write to a non-existing sheet causes an exception (*.xls)
	checkException(writeWorksheet(wb.xls, mtcars, "sheetDoesNotExist"))
	
	# Check that attempting to write to a non-existing sheet causes an exception (*.xlsx)
	checkException(writeWorksheet(wb.xlsx, mtcars, "sheetDoesNotExist"))
}
