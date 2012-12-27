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
# Tests around setting the active sheet of an Excel workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.setActiveSheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookSetActiveSheet.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookSetActiveSheet.xlsx"), create = FALSE)
	
	# Check that setting the active sheet works ok (*.xls)
	# (assumes that 'getActiveSheetIndex' works fine)
	setActiveSheet(wb.xls, 1)
	checkTrue(getActiveSheetIndex(wb.xls) == 1)
	setActiveSheet(wb.xls, 3)
	checkTrue(getActiveSheetIndex(wb.xls) == 3)
	setActiveSheet(wb.xls, "Sheet2")
	checkTrue(getActiveSheetIndex(wb.xls) == 2)
	
	# Check that setting the active sheet works ok (*.xlsx)
	# (assumes that 'getActiveSheetIndex' works fine)
	setActiveSheet(wb.xlsx, 1)
	checkTrue(getActiveSheetIndex(wb.xlsx) == 1)
	setActiveSheet(wb.xlsx, 3)
	checkTrue(getActiveSheetIndex(wb.xlsx) == 3)
	setActiveSheet(wb.xlsx, "Sheet2")
	checkTrue(getActiveSheetIndex(wb.xlsx) == 2)
	
	# Check that setting an illegal active sheet throws an exception (*.xls)
	checkException(setActiveSheet(wb.xls, 19))
	checkException(setActiveSheet(wb.xls, "SheetWhichDoesNotExist"))
	
	# Check that setting an illegal active sheet throws an exception (*.xlsx)
	checkException(setActiveSheet(wb.xlsx, 19))
	checkException(setActiveSheet(wb.xlsx, "SheetWhichDoesNotExist"))
}
