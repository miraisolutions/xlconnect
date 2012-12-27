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

test.workbook.isSheetHidden <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xlsx"), create = FALSE)
	
	# Check if sheets are hidden (*.xls)
	checkTrue(isSheetHidden(wb.xls, 2))
	checkTrue(isSheetHidden(wb.xls, "BBB"))
	checkTrue(!isSheetHidden(wb.xls, 1))
	checkTrue(!isSheetHidden(wb.xls, "AAA"))
	checkTrue(!isSheetHidden(wb.xls, 3))
	checkTrue(!isSheetHidden(wb.xls, "CCC"))
	checkTrue(!isSheetHidden(wb.xls, 4)) # Sheet is actually very hidden!
	checkTrue(!isSheetHidden(wb.xls, "DDD")) # Sheet is actually very hidden!
	
	# Check if sheets are hidden (*.xlsx)
	checkTrue(isSheetHidden(wb.xlsx, 2))
	checkTrue(isSheetHidden(wb.xlsx, "BBB"))
	checkTrue(!isSheetHidden(wb.xlsx, 1))
	checkTrue(!isSheetHidden(wb.xlsx, "AAA"))
	checkTrue(!isSheetHidden(wb.xlsx, 3))
	checkTrue(!isSheetHidden(wb.xlsx, "CCC"))
	checkTrue(!isSheetHidden(wb.xlsx, 4)) # Sheet is actually very hidden!
	checkTrue(!isSheetHidden(wb.xlsx, "DDD")) # Sheet is actually very hidden!
	
	# Check if quering invalid/non-existing sheets
	# causes an exception (*.xls)
	checkException(isSheetHidden(wb.xls, 200))
	checkException(isSheetHidden(wb.xls, "Sheet does not exist"))
	checkException(isSheetHidden(wb.xls, "'Illegal sheet name"))
	
	# Check if quering invalid/non-existing sheets
	# causes an exception (*.xlsx)
	checkException(isSheetHidden(wb.xlsx, 200))
	checkException(isSheetHidden(wb.xlsx, "Sheet does not exist"))
	checkException(isSheetHidden(wb.xlsx, "'Illegal sheet name"))
	
}
