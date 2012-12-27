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
# Tests around unhiding Excel worksheets
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.unhideSheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookHiddenSheets.xlsx"), create = FALSE)
	
	# Check that unhiding sheets works correctly (*.xls)
	# (assumes 'isSheetHidden' and 'isSheetVeryHidden' work correctly)
	unhideSheet(wb.xls, 2)
	unhideSheet(wb.xls, "DDD")
	checkTrue(!isSheetHidden(wb.xls, 2))
	checkTrue(!isSheetVeryHidden(wb.xls, "DDD"))
	
	# Check that unhiding sheets works correctly (*.xlsx)
	# (assumes 'isSheetHidden' and 'isSheetVeryHidden' work correctly)
	unhideSheet(wb.xlsx, 2)
	unhideSheet(wb.xlsx, "DDD")
	checkTrue(!isSheetHidden(wb.xlsx, 2))
	checkTrue(!isSheetVeryHidden(wb.xlsx, "DDD"))
	
	# Check that attempting to unhide an illegal sheet throws an exception (*.xls)
	checkException(unhideSheet(wb.xls, 58))
	checkException(unhideSheet(wb.xls, "SheetWhichDoesNotExist"))
	
	# Check that attempting to unhide an illegal sheet throws an exception (*.xlsx)
	checkException(unhideSheet(wb.xlsx, 58))
	checkException(unhideSheet(wb.xlsx, "SheetWhichDoesNotExist"))
}
