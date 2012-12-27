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
# Tests around cloning Excel worksheets
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.cloneSheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookCloneSheet.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookCloneSheet.xlsx"), create = FALSE)
	
	# Check cloning of worksheet (*.xls)
	checkNoException(cloneSheet(wb.xls, sheet = "Test1", name = "Clone1"))
	checkNoException(readWorksheet(wb.xls, sheet = "Clone1"))
	
	# Check cloning of worksheet (*.xlsx)
	checkNoException(cloneSheet(wb.xlsx, sheet = "Test1", name = "Clone1"))
	checkNoException(readWorksheet(wb.xlsx, sheet = "Clone1"))
	
	# Check that attempting to clone a non-existing worksheet throws an exception (*.xls)
	checkException(cloneSheet(wb.xls, sheet = "NotThere", name = "MyClone"))
	
	# Check that attempting to clone a non-existing worksheet throws an exception (*.xlsx)
	checkException(cloneSheet(wb.xlsx, sheet = "NotThere", name = "MyClone"))
	
	# Check that attempting to assign an invalid name to a cloned sheet throws an exception (*.xls)
	checkException(cloneSheet(wb.xls, sheet = "Test1", name = "'illegal"))
	
	# Check that attempting to assign an invalid name to a cloned sheet throws an exception (*.xlsx)
	checkException(cloneSheet(wb.xlsx, sheet = "Test1", name = "'illegal"))
	
}

