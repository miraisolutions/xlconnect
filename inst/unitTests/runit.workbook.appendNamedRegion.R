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
# Tests around appending data to named regions
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.appendNamedRegion <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookAppend.xls"))
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookAppend.xlsx"))
	
	refCoord = matrix(c(9, 5, 73, 15), ncol = 2, byrow = TRUE)
	
	# Check that appending data to a named region produces the expected result (*.xls)
	appendNamedRegion(wb.xls, mtcars, name = "mtcars")
	res = readNamedRegion(wb.xls, name = "mtcars")
	checkEquals(res, normalizeDataframe(rbind(mtcars, mtcars)), check.attributes = FALSE, check.names = TRUE)
	checkEquals(getReferenceCoordinatesForName(wb.xls, "mtcars"), refCoord)
	
	# Check that appending data to a named region produces the expected result (*.xlsx)
	appendNamedRegion(wb.xlsx, mtcars, name = "mtcars")
	res = readNamedRegion(wb.xlsx, name = "mtcars")
	checkEquals(res, normalizeDataframe(rbind(mtcars, mtcars)), check.attributes = FALSE, check.names = TRUE)
	checkEquals(getReferenceCoordinatesForName(wb.xlsx, "mtcars"), refCoord)
	
	# Check that trying to append to an non-existing named region throws an error (*.xls)
	checkException(appendNamedRegion(wb.xls, mtcars, name = "doesNotExist"))
	
	# Check that trying to append to an non-existing named region throws an error (*.xlsx)
	checkException(appendNamedRegion(wb.xlsx, mtcars, name = "doesNotExist"))
	
	# Re-create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookAppend.xls"))
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookAppend.xlsx"))
	
	refCoord = matrix(c(9, 5, 194, 15), ncol = 2, byrow = TRUE)
	
	# Check that appending data to a named region with a different structure results
	# in the correct bounding box for the re-defined named region (*.xls)
	appendNamedRegion(wb.xls, airquality, name = "mtcars")
	checkEquals(getReferenceCoordinatesForName(wb.xls, "mtcars"), refCoord)
	
	# Check that appending data to a named region with a different structure results
	# in the correct bounding box for the re-defined named region (*.xlsx)
	appendNamedRegion(wb.xlsx, airquality, name = "mtcars")
	checkEquals(getReferenceCoordinatesForName(wb.xlsx, "mtcars"), refCoord)
}
