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
# Tests around setting Excel worksheet positions
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.setSheetPos <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookSetSheetPos.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookSetSheetPos.xlsx"), create = TRUE)
	
	createSheet(wb.xls, c("A", "B", "C", "D"))
	createSheet(wb.xlsx, c("A", "B", "C", "D"))
	
	# Check positions (*.xls)
	setSheetPos(wb.xls, sheet = c("D", "B"), pos = c(2, 1))
	checkEquals(getSheets(wb.xls), c("B", "A", "D", "C"))
	checkEquals(getSheetPos(wb.xls, sheet = c("A", "B", "C", "D")), c("A" = 2, "B" = 1, "C" = 4, "D" = 3))
	
	# Check positions (*.xlsx)
	setSheetPos(wb.xlsx, sheet = c("D", "B"), pos = c(2, 1))
	checkEquals(getSheets(wb.xlsx), c("B", "A", "D", "C"))
	checkEquals(getSheetPos(wb.xlsx, sheet = c("A", "B", "C", "D")), c("A" = 2, "B" = 1, "C" = 4, "D" = 3))
	
	# Check that trying to set a non-existing index (out of bounds) results in an exception (*.xls)
	checkException(setSheetPos(wb.xls, sheet = "A", pos = -1))
	checkException(setSheetPos(wb.xls, sheet = "A", pos = 36298))
	
	# Check that trying to set a non-existing position (out of bounds) results in an exception (*.xlsx)
	checkException(setSheetPos(wb.xlsx, sheet = "A", pos = -1))
	checkException(setSheetPos(wb.xlsx, sheet = "A", pos = 36298))
	
	# Check that trying to set the position of a non-existing worksheet results in an exception (*.xls)
	checkException(setSheetPos(wb.xls, sheet = "NotThere", pos = 2))
	
	# Check that trying to set the position of a non-existing worksheet results in an exception (*.xlsx)
	checkException(setSheetPos(wb.xlsx, sheet = "NotThere", pos = 2))
	
}
