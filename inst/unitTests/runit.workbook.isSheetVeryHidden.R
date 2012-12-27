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

test.workbook.isSheetVeryHidden <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xlsx"), create = FALSE)
	
	# Check if sheets are hidden (*.xls)
	checkTrue(isSheetVeryHidden(wb.xls, 4))
	checkTrue(isSheetVeryHidden(wb.xls, "DDD"))
	checkTrue(!isSheetVeryHidden(wb.xls, 1))
	checkTrue(!isSheetVeryHidden(wb.xls, "AAA"))
	checkTrue(!isSheetVeryHidden(wb.xls, 2)) # Sheet is actually hidden only!
	checkTrue(!isSheetVeryHidden(wb.xls, "BBB")) # Sheet is actually hidden only!
	checkTrue(!isSheetVeryHidden(wb.xls, 3))
	checkTrue(!isSheetVeryHidden(wb.xls, "CCC"))
	
	# Check if sheets are hidden (*.xlsx)
	checkTrue(isSheetVeryHidden(wb.xlsx, 4))
	checkTrue(isSheetVeryHidden(wb.xlsx, "DDD"))
	checkTrue(!isSheetVeryHidden(wb.xlsx, 1))
	checkTrue(!isSheetVeryHidden(wb.xlsx, "AAA"))
	checkTrue(!isSheetVeryHidden(wb.xlsx, 2)) # Sheet is actually hidden only!
	checkTrue(!isSheetVeryHidden(wb.xlsx, "BBB")) # Sheet is actually hidden only!
	checkTrue(!isSheetVeryHidden(wb.xlsx, 3))
	checkTrue(!isSheetVeryHidden(wb.xlsx, "CCC"))
	
	# Check if quering invalid/non-existing sheets
	# causes an exception (*.xls)
	checkException(isSheetVeryHidden(wb.xls, 200))
	checkException(isSheetVeryHidden(wb.xls, "Sheet does not exist"))
	checkException(isSheetVeryHidden(wb.xls, "'Illegal sheet name"))
	
	# Check if quering invalid/non-existing sheets
	# causes an exception (*.xlsx)
	checkException(isSheetVeryHidden(wb.xlsx, 200))
	checkException(isSheetVeryHidden(wb.xlsx, "Sheet does not exist"))
	checkException(isSheetVeryHidden(wb.xlsx, "'Illegal sheet name"))
	
}
