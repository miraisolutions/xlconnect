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
# Tests around querying defined Excel names
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.getDefinedNames <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookDefinedNames.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookDefinedNames.xlsx"), create = FALSE)
	
	# Names defined in workbooks
	expectedNamesValidOnly <- c("FirstName", "SecondName", "FourthName", "FifthName")
	expectedNamesAll <- c("FirstName", "SecondName", "ThirdName", "FourthName", "FifthName")
	
	# Check that all and only the expected names exist (*.xls)
	definedNames <- getDefinedNames(wb.xls, validOnly = TRUE)
	checkTrue(length(setdiff(expectedNamesValidOnly, definedNames)) == 0 && length(setdiff(definedNames, expectedNamesValidOnly)) == 0)
	definedNames <- getDefinedNames(wb.xls, validOnly = FALSE)
	checkTrue(length(setdiff(expectedNamesAll, definedNames)) == 0 && length(setdiff(definedNames, expectedNamesAll)) == 0)
	
	# Check that all and only the expected names exist (*.xlsx)
	definedNames <- getDefinedNames(wb.xlsx, validOnly = TRUE)
	checkTrue(length(setdiff(expectedNamesValidOnly, definedNames)) == 0 && length(setdiff(definedNames, expectedNamesValidOnly)) == 0)
	definedNames <- getDefinedNames(wb.xlsx, validOnly = FALSE)
	checkTrue(length(setdiff(expectedNamesAll, definedNames)) == 0 && length(setdiff(definedNames, expectedNamesAll)) == 0)
	
}
