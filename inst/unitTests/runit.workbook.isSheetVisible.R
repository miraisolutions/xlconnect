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
# Tests around checking the visibility of Excel worksheets
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.isSheetVisible <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xlsx"), create = FALSE)
	
	# Check if sheets are visible (*.xls)
	checkTrue(!isSheetVisible(wb.xls, 2))
	checkTrue(!isSheetVisible(wb.xls, "BBB"))
	checkTrue(isSheetVisible(wb.xls, 1))
	checkTrue(isSheetVisible(wb.xls, "AAA"))
	checkTrue(isSheetVisible(wb.xls, 3))
	checkTrue(isSheetVisible(wb.xls, "CCC"))
	checkTrue(!isSheetVisible(wb.xls, 4))
	checkTrue(!isSheetVisible(wb.xls, "DDD"))
	
	# Check if sheets are visible (*.xls)
	checkTrue(!isSheetVisible(wb.xlsx, 2))
	checkTrue(!isSheetVisible(wb.xlsx, "BBB"))
	checkTrue(isSheetVisible(wb.xlsx, 1))
	checkTrue(isSheetVisible(wb.xlsx, "AAA"))
	checkTrue(isSheetVisible(wb.xlsx, 3))
	checkTrue(isSheetVisible(wb.xlsx, "CCC"))
	checkTrue(!isSheetVisible(wb.xlsx, 4))
	checkTrue(!isSheetVisible(wb.xlsx, "DDD"))
	
}
