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
# Tests around querying the last (non-empty) row of a worksheet
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.getLastRow <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookGetLastRow.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookGetLastRow.xlsx"), create = FALSE)
	
	# Check if last row is determined correctly (*.xls)
	checkEquals(as.vector(getLastRow(wb.xls, "mtcars")), 33)
	checkEquals(as.vector(getLastRow(wb.xls, "mtcars2")), 41)
	checkEquals(as.vector(getLastRow(wb.xls, "mtcars3")), 49)
	
	# Check if last row is determined correctly (*.xlsx)
	checkEquals(as.vector(getLastRow(wb.xlsx, "mtcars")), 33)
	checkEquals(as.vector(getLastRow(wb.xlsx, "mtcars2")), 41)
	checkEquals(as.vector(getLastRow(wb.xlsx, "mtcars3")), 49)
	
	# Check that querying the last row of a non-existing worksheet throws an exception (*.xls)
	checkException(getLastRow(wb.xls, "doesNotExist"))
	
	# Check that querying the last row of a non-existing worksheet throws an exception (*.xlsx)
	checkException(getLastRow(wb.xlsx, "doesNotExist"))
	
	# Last row of an empty worksheet is 1 (.xls)
	checkEquals(getLastRow(wb.xls, "empty"), c(empty = 1))
	
	# Last row of an empty worksheet is 1 (.xlsx)
	checkEquals(getLastRow(wb.xlsx, "empty"), c(empty = 1))
}
