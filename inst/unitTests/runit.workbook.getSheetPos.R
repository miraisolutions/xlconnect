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
# Tests around querying Excel worksheet positions
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.getSheetPos <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookGetSheetPos.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookGetSheetPos.xlsx"), create = TRUE)
	
	createSheet(wb.xls, c("Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4"))
	createSheet(wb.xlsx, c("Sheet 1", "Sheet 2", "Sheet 3", "Sheet 4"))
	
	# Check positions (*.xls)
	checkEquals(getSheetPos(wb.xls, c("Sheet 3", "Sheet 2", "Sheet 4", "Sheet 1")), 
		c("Sheet 3" = 3, "Sheet 2" = 2, "Sheet 4" = 4, "Sheet 1" = 1))
	
	# Check positions (*.xlsx)
	checkEquals(getSheetPos(wb.xlsx, c("Sheet 3", "Sheet 2", "Sheet 4", "Sheet 1")), 
		c("Sheet 3" = 3, "Sheet 2" = 2, "Sheet 4" = 4, "Sheet 1" = 1))
	
	# Check that querying a non-existing worksheet results in a 0 index (*.xls)
	checkEquals(getSheetPos(wb.xls, "NotExisting"), c("NotExisting" = 0))
	checkEquals(as.vector(getSheetPos(wb.xls, "%#?%+?[-")), 0)
	
	# Check that querying a non-existing worksheet results in a 0 index (*.xls)
	checkEquals(getSheetPos(wb.xlsx, "NotExisting"), c("NotExisting" = 0))
	checkEquals(as.vector(getSheetPos(wb.xlsx, "%#?%+?[-")), 0)
	
}
