#############################################################################
#
# XLConnect
# Copyright (C) 2010-2025 Mirai Solutions GmbH
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
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
#############################################################################

#############################################################################
#
# Tests around writing to an Excel worksheet
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.writeWorksheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookWriteWorksheet.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookWriteWorksheet.xlsx"), create = TRUE)
	
	# Check that trying to write an object which cannot be converted to a data.frame
	# causes an exception (*.xls)
	createSheet(wb.xls, "test1")
	checkException(writeWorksheet(wb.xls, search, "test1"))
	
	# Check that trying to write an object which cannot be converted to a data.frame
	# causes an exception (*.xlsx)
	createSheet(wb.xlsx, "test1")
	checkException(writeWorksheet(wb.xlsx, search, "test1"))
	
	# Check that attempting to write to a non-existing sheet causes an exception (*.xls)
	checkException(writeWorksheet(wb.xls, mtcars, "sheetDoesNotExist"))
	
	# Check that attempting to write to a non-existing sheet causes an exception (*.xlsx)
	checkException(writeWorksheet(wb.xlsx, mtcars, "sheetDoesNotExist"))
	
	# Load workbooks with formulas on them
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookOverwriteFormulas.xls"))
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookOverwriteFormulas.xlsx"))
	
	# Initialize mtcars_mod as is defined with the formula in the carb column (set equal to the gear column)
	mtcars_mod = mtcars
	mtcars_mod$carb = mtcars_mod$gear
	
	test_overwrite_formula <- function(wb, expected, overwrite = TRUE) {
	  writeWorksheet(wb, mtcars, "mtcars_formula", startRow = 9, startCol = 5, overwriteFormulaCells = overwrite)
	  res = readWorksheet(wb, "mtcars_formula")
	  checkEquals(res, normalizeDataframe(expected), check.attributes = FALSE, check.names = TRUE)
	}
	
	# Check that formulas can be kept in existing named region (*.xls)
	test_overwrite_formula(wb.xls, mtcars_mod, overwrite = FALSE)
	
	# Check that formulas can be kept in existing named region (*.xlsx)
	test_overwrite_formula(wb.xlsx, mtcars_mod, overwrite = FALSE)
	
	# Check that formulas are overwritten (*.xls)
	test_overwrite_formula(wb.xls, mtcars)
	
	# Check that formulas are overwritten (*.xlsx)
	test_overwrite_formula(wb.xlsx, mtcars)
}
