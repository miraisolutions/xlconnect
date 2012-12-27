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
# Tests around removing Excel names
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.removeName <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookRemoveName.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookRemoveName.xlsx"), create = FALSE)
	
	# Check that when removing a name from a worksheet it does not exist anymore (*.xls)
	# (assumes 'existsName' is working correctly)
	removeName(wb.xls, "AA")
	checkTrue(!existsName(wb.xls, "AA"))
	
	# Check that when removing a name from a worksheet it does not exist anymore (*.xlsx)
	# (assumes 'existsName' is working correctly)
	removeName(wb.xlsx, "AA")
	checkTrue(!existsName(wb.xlsx, "AA"))
	
	# Check that attempting to remove a non-existing name does not throw an exception (*.xls)
	checkNoException(removeName(wb.xls, "NameWhichDoesNotExist"))
	
	# Check that attempting to remove a non-existing name does not throw an exception (*.xlsx)
	checkNoException(removeName(wb.xlsx, "NameWhichDoesNotExist"))
}
