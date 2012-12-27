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
# Tests around renaming Excel worksheets
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.renameSheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookRenameSheet.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookRenameSheet.xlsx"), create = TRUE)
	
	# Check that renaming a sheet (using its name) works fine (*.xls)
	# (assumes 'createSheet' and 'existsSheet' to be working properly)
	createSheet(wb.xls, name = "OldName1")
	renameSheet(wb.xls, sheet = "OldName1", newName = "NewName1")
	checkTrue(existsSheet(wb.xls, "NewName1"))
	checkTrue(!existsSheet(wb.xls, "OldName1"))
	
	# Check that renaming a sheet (using its name) works fine (*.xlsx)
	# (assumes 'createSheet' and 'existsSheet' to be working properly)
	createSheet(wb.xlsx, name = "OldName1")
	renameSheet(wb.xlsx, sheet = "OldName1", newName = "NewName1")
	checkTrue(existsSheet(wb.xlsx, "NewName1"))
	checkTrue(!existsSheet(wb.xlsx, "OldName1"))
	
	# Check that renaming a sheet (using its index) works fine (*.xls)
	# (assumes 'createSheet' and 'existsSheet' to be working properly)
	createSheet(wb.xls, name = "OldName2")
	renameSheet(wb.xls, sheet = 2, newName = "NewName2")
	checkTrue(existsSheet(wb.xls, "NewName2"))
	checkTrue(!existsSheet(wb.xls, "OldName2"))
	
	# Check that renaming a sheet (using its index) works fine (*.xlsx)
	# (assumes 'createSheet' and 'existsSheet' to be working properly)
	createSheet(wb.xlsx, name = "OldName2")
	renameSheet(wb.xlsx, sheet = 2, newName = "NewName2")
	checkTrue(existsSheet(wb.xlsx, "NewName2"))
	checkTrue(!existsSheet(wb.xlsx, "OldName2"))
	
	# Check that renaming a non-existing sheet throws an exception (*.xls)
	checkException(renameSheet(wb.xls, sheet = "NonExisting", newName = "ShouldStillNotExist"))
	
	# Check that renaming a non-existing sheet throws an exception (*.xlsx)
	checkException(renameSheet(wb.xlsx, sheet = "NonExisting", newName = "ShouldStillNotExist"))
	
	# Check that renaming a sheet with an invalid new name throws an exception (*.xls)
	createSheet(wb.xls, name = "SomeName")
	checkException(renameSheet(wb.xls, sheet = "SomeName", newName = "'invalid"))
	
	# Check that renaming a sheet with an invalid new name throws an exception (*.xlsx)
	createSheet(wb.xlsx, name = "SomeName")
	checkException(renameSheet(wb.xlsx, sheet = "SomeName", newName = "'invalid"))
	
}
