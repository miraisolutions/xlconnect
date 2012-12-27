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
# Tests around appending data to worksheets
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.appendWorksheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookAppend.xls"))
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookAppend.xlsx"))
	
	# Check that appending data to a worksheet produces the expected result (*.xls)
	appendWorksheet(wb.xls, mtcars, sheet = "mtcars")
	res = readWorksheet(wb.xls, sheet = "mtcars")
	checkEquals(getLastRow(wb.xls, "mtcars"), c(mtcars = 73))
	checkEquals(res, normalizeDataframe(rbind(mtcars, mtcars)), check.attributes = FALSE, check.names = TRUE)
	
	# Check that appending data to a named region produces the expected result (*.xlsx)
	appendWorksheet(wb.xlsx, mtcars, sheet = "mtcars")
	res = readWorksheet(wb.xlsx, sheet = "mtcars")
	checkEquals(getLastRow(wb.xlsx, "mtcars"), c(mtcars = 73))
	checkEquals(res, normalizeDataframe(rbind(mtcars, mtcars)), check.attributes = FALSE, check.names = TRUE)
	
	# Check that trying to append to an non-existing worksheet throws an error (*.xls)
	checkException(appendWorksheet(wb.xls, mtcars, sheet = "doesNotExist"))
	
	# Check that trying to append to an non-existing worksheet throws an error (*.xlsx)
	checkException(appendWorksheet(wb.xlsx, mtcars, sheet = "doesNotExist"))
}
