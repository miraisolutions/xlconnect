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
# Tests around writing named regions in an Excel workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.writeNamedRegion <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookWriteNamedRegion.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookWriteNamedRegion.xlsx"), create = TRUE)
	
	# Check that trying to write an object which cannot be converted to a data.frame
	# causes an exception (*.xls)
	createName(wb.xls, "test1", "Test1!$C$8")
	checkException(writeNamedRegion(wb.xls, search, "test1"))
	
	# Check that trying to write an object which cannot be converted to a data.frame
	# causes an exception (*.xlsx)
	createName(wb.xlsx, "test1", "Test1!$C$8")
	checkException(writeNamedRegion(wb.xlsx, search, "test1"))
	
	# Check that attempting to write to a non-existing name causes an exception (*.xls)
	checkException(writeNamedRegion(wb.xls, mtcars, "nameDoesNotExist"))
	
	# Check that attempting to write to a non-existing name causes an exception (*.xlsx)
	checkException(writeNamedRegion(wb.xlsx, mtcars, "nameDoesNotExist"))
	
	# Check that attempting to write to a name which referes to a non-existing sheet
	# causes an exception (*.xls)
	createName(wb.xls, "nope", "NonExistingSheet!A1")
	checkException(writeNamedRegion(wb.xls, mtcars, "nope"))
	
	# Check that attempting to write to a name which referes to a non-existing sheet
	# causes an exception (*.xlsx)
	createName(wb.xlsx, "nope", "NonExistingSheet!A1")
	checkException(writeNamedRegion(wb.xlsx, mtcars, "nope"))
  
	# Check that writing an empty data.frame does not cause an error (*.xls)
  createSheet(wb.xls, "empty")
	createName(wb.xls, "empty1", "empty!A1")
  createName(wb.xls, "empty2", "empty!D10")
	checkNoException(writeNamedRegion(wb.xls, data.frame(), "empty1"))
	checkNoException(writeNamedRegion(wb.xls, data.frame(a = character(0), b = numeric(0)), "empty2"))
	
	# Check that writing an empty data.frame does not cause an error (*.xlsx)
	createSheet(wb.xlsx, "empty")
	createName(wb.xlsx, "empty1", "empty!A1")
	createName(wb.xlsx, "empty2", "empty!D10")
	checkNoException(writeNamedRegion(wb.xlsx, data.frame(), "empty1"))
	checkNoException(writeNamedRegion(wb.xlsx, data.frame(a = character(0), b = numeric(0)), "empty2"))
	
}
