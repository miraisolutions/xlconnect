#############################################################################
#
# XLConnect
# Copyright (C) 2010-2012 Mirai Solutions GmbH
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
# Tests around clearing named regions from an Excel Workbook
# 
# Author: Nicola Lambiase, Mirai Solutions GmbH
#
#############################################################################

test.workbook.clearNamedRegion <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookClearCells.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookClearCells.xlsx"), create = FALSE)
	
	checkDf <- data.frame(
			"eight" = 36:40,
			stringsAsFactors = F
	)
	
	# Check that clearing all the named regions from a sheet returns an empty sheet (*.xls)
	clearNamedRegion(wb.xls, c("region1", "region2"))
	res <- readWorksheet(wb.xls, "clearNamedRegion", endRow = 7, header = TRUE)
	checkEquals(res, checkDf)
	res <- readWorksheet(wb.xls, "clearNamedRegion", startRow=10, endRow = 15, header = TRUE)
	checkEquals(res, checkDf)
	
	# Check that clearing all the named regions from a sheet returns an empty sheet (*.xlsx)
	clearNamedRegion(wb.xlsx, c("region1", "region2"))
	res <- readWorksheet(wb.xlsx, "clearNamedRegion", endRow = 7, header = TRUE)
	checkEquals(res, checkDf)
	res <- readWorksheet(wb.xlsx, "clearNamedRegion", startRow=10, endRow = 15, header = TRUE)
	checkEquals(res, checkDf)
}
