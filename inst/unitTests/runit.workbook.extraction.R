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
# Test workbook extraction & replacement operators
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.extraction <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookExtractionOperators.xls"), create = TRUE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookExtractionOperators.xlsx"), create = TRUE)
	
	# Check that writing a data set to a worksheet works (*.xls) 
	wb.xls["mtcars1"] = mtcars
	checkTrue("mtcars1" %in% getSheets(wb.xls))
	checkEquals(as.vector(getLastRow(wb.xls, "mtcars1")), 33)
	wb.xls["mtcars2", startRow = 6, startCol = 11, header = FALSE] = mtcars
	checkTrue("mtcars2" %in% getSheets(wb.xls))
	checkEquals(as.vector(getLastRow(wb.xls, "mtcars2")), 37)
	
	# Check that writing a data set to a worksheet works (*.xlsx) 
	wb.xlsx["mtcars1"] = mtcars
	checkTrue("mtcars1" %in% getSheets(wb.xlsx))
	checkEquals(as.vector(getLastRow(wb.xlsx, "mtcars1")), 33)
	wb.xlsx["mtcars2", startRow = 6, startCol = 11, header = FALSE] = mtcars
	checkTrue("mtcars2" %in% getSheets(wb.xlsx))
	checkEquals(as.vector(getLastRow(wb.xlsx, "mtcars2")), 37)
	
	# Check that reading data from a worksheet works (*.xls)
	checkEquals(dim(wb.xls["mtcars1"]), c(32, 11))
	checkEquals(dim(wb.xls["mtcars2"]), c(31, 11))
	
	# Check that reading data from a worksheet works (*.xlsx)
	checkEquals(dim(wb.xlsx["mtcars1"]), c(32, 11))
	checkEquals(dim(wb.xlsx["mtcars2"]), c(31, 11))
	
	# Check that writing data to a named region works (*.xls)
	wb.xls[["mtcars3", "mtcars3!$B$7"]] = mtcars
	checkTrue("mtcars3" %in% getDefinedNames(wb.xls))
	checkEquals(as.vector(getLastRow(wb.xls, "mtcars3")), 39)
	wb.xls[["mtcars4", "mtcars4!$D$8", rownames = "Car"]] = mtcars
	checkTrue("mtcars4" %in% getDefinedNames(wb.xls))
	checkEquals(as.vector(getLastRow(wb.xls, "mtcars4")), 40)
	
	# Check that writing data to a named region works (*.xlsx)
	wb.xlsx[["mtcars3", "mtcars3!$B$7"]] = mtcars
	checkTrue("mtcars3" %in% getDefinedNames(wb.xlsx))
	checkEquals(as.vector(getLastRow(wb.xlsx, "mtcars3")), 39)
	wb.xlsx[["mtcars4", "mtcars4!$D$8", rownames = "Car"]] = mtcars
	checkTrue("mtcars4" %in% getDefinedNames(wb.xlsx))
	checkEquals(as.vector(getLastRow(wb.xlsx, "mtcars4")), 40)
	
	# Check that reading data from a named region works (*.xls)
	checkEquals(dim(wb.xls[["mtcars3"]]), c(32, 11))
	checkEquals(dim(wb.xls[["mtcars4"]]), c(32, 12))
	
	# Check that reading data from a named region works (*.xlsx)
	checkEquals(dim(wb.xlsx[["mtcars3"]]), c(32, 11))
	checkEquals(dim(wb.xlsx[["mtcars4"]]), c(32, 12))
	
}
