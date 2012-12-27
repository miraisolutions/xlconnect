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
# Tests around removing Excel worksheets
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.removeSheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookRemoveSheet.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookRemoveSheet.xlsx"), create = FALSE)
	
	# Check that removing a sheet works fine (*.xls)
	# (assumes 'existsSheet' to be working properly)
	removeSheet(wb.xls, "BBB")
	checkTrue(!existsSheet(wb.xls, "BBB"))
	
	# Check that removing a sheet works fine (*.xlsx)
	# (assumes 'existsSheet' to be working properly)
	removeSheet(wb.xlsx, "BBB")
	checkTrue(!existsSheet(wb.xlsx, "BBB"))
	
	# Check that removing a non-existing sheet does not throw an exception (*.xls)
	checkNoException(removeSheet(wb.xls, 35))
	checkNoException(removeSheet(wb.xls, "SheetWhichDoesNotExist"))
	
	# Check that removing a non-existing sheet does not throw an exception (*.xlsx)
	checkNoException(removeSheet(wb.xlsx, 35))
	checkNoException(removeSheet(wb.xlsx, "SheetWhichDoesNotExist"))
}
