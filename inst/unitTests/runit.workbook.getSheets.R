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
# Tests around querying Excel worksheets
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.getSheets <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookSheets.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookSheets.xlsx"), create = FALSE)
	
	# Sheets defined in workbooks
	expectedSheets <- c("A1", "B 2", "$$", "=", "@}", "11. Oct.", "\"quote\"", "+0")
	
	# Check that all and only the expected sheets exist (*.xls)
	definedSheets <- getSheets(wb.xls)
	checkTrue(length(setdiff(expectedSheets, definedSheets)) == 0 && length(setdiff(definedSheets, expectedSheets)) == 0)
	
	# Check that all and only the expected sheets exist (*.xlsx)
	definedSheets <- getSheets(wb.xlsx)
	checkTrue(length(setdiff(expectedSheets, definedSheets)) == 0 && length(setdiff(definedSheets, expectedSheets)) == 0)
	
}
