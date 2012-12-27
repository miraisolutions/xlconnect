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
			"one" = 1:5,	
			"two" = 6:10,
			"three" = 11:15,
			"four" = 16:20,
			"five" = 21:25,
			"six" = 26:30,
			"seven" = 31:35,
			stringsAsFactors = F
	)
	
	# Check that clearing 2 of 3 named regions in a sheet returns only the third one (*.xls)
	clearNamedRegion(wb.xls, c("region1", "region2"))
	res <- readWorksheet(wb.xls, "clearNamedRegion", header = TRUE)
	checkEquals(res, checkDf)
	
	# Check that clearing 2 of 3 named regions in a sheet returns only the third one (*.xlsx)
	clearNamedRegion(wb.xlsx, c("region1", "region2"))
	res <- readWorksheet(wb.xlsx, "clearNamedRegion", header = TRUE)
	checkEquals(res, checkDf)
}
