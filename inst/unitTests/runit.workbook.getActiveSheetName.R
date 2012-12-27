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
# Tests around querying the active sheet name of an Excel workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.getActiveSheetName <- function() {

	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookActiveSheetIndexAndName.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookActiveSheetIndexAndName.xlsx"), create = FALSE)
	
	# Check that the active sheet name is 'Fifth Sheet' (*.xls)
	checkTrue(getActiveSheetName(wb.xls) == "Fifth Sheet")
	
	# Check that the active sheet name is 'Fifth Sheet' (*.xlsx)
	checkTrue(getActiveSheetName(wb.xlsx) == "Fifth Sheet")
	
}
