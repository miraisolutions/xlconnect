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
# Tests around getting the bounding box coordinates from an Excel worksheet
# 
# Author: Nicola Lambiase, Mirai Solutions GmbH
#
#############################################################################

test.workbook.getBoundingBox <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookReadWorksheet.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookReadWorksheet.xlsx"), create = FALSE)
	
	dim1 <- matrix(c(17, 6, 25, 9), dimnames = list(c(), c("Test5")))
	dim2 <- matrix(c(17, 7, 24, 9), dimnames = list(c(), c("Test5")))
	dim3 <- matrix(c(17, 7, 25, 9), dimnames = list(c(), c("Test5")))
	dim4 <- matrix(c(17, 6, 24, 9), dimnames = list(c(), c("Test5")))
	dim5 <- matrix(c(11, 6, 16, 9, 8, 4, 16, 7, 17, 6, 25, 9), ncol=3, dimnames = list(c(), c("Test1","Test4","Test5")))
	
	# Checking the dimensions of a bounding box when no row/column is specified (*.xls)
	res <- getBoundingBox(wb.xls, sheet = "Test5")
	checkEquals(res, dim1)
	
	# Checking the dimensions of a bounding box when no row/column is specified (*.xlsx)
	res <- getBoundingBox(wb.xlsx, sheet = "Test5")
	checkEquals(res, dim1)

	# Checking the dimensions of a bounding box when start and end cells are specified (*.xls)
	res <- getBoundingBox(wb.xls, sheet = "Test5", startRow=17, startCol=7, endRow=24, endCol=9)
	checkEquals(res, dim2)
	
	# Checking the dimensions of a bounding box when start and end cells are specified (*.xlsx)
	res <- getBoundingBox(wb.xlsx, sheet = "Test5", startRow=17, startCol=7, endRow=24, endCol=9)
	checkEquals(res, dim2)
	
	# Checking the dimensions of a bounding box when start and end columns are specified (*.xls)
	res <- getBoundingBox(wb.xls, sheet = "Test5", startCol=7, endCol=9)
	checkEquals(res, dim3)
	
	# Checking the dimensions of a bounding box when start and end columns are specified (*.xlsx)
	res <- getBoundingBox(wb.xlsx, sheet = "Test5", startCol=7, endCol=9)
	checkEquals(res, dim3)
	
	# Checking the dimensions of a bounding box when only the end row is specified (*.xls)
	res <- getBoundingBox(wb.xls, sheet = "Test5", endRow=24)
	checkEquals(res, dim4)
	
	# Checking the dimensions of a bounding box when only the end row is specified (*.xlsx)
	res <- getBoundingBox(wb.xlsx, sheet = "Test5", endRow=24)
	checkEquals(res, dim4)
	
	# Checking the dimensions of bounding boxes when multiple sheets are specified (*.xls)
	res <- getBoundingBox(wb.xls, sheet = c("Test1","Test4","Test5"))
	checkEquals(res, dim5)
	
	# Checking the dimensions of bounding boxes when multiple sheets are specified (*.xlsx)
	res <- getBoundingBox(wb.xlsx, sheet = c("Test1","Test4","Test5"))
	checkEquals(res, dim5)
}