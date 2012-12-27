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
# Tests around clearing ranges from an Excel Workbook
# 
# Author: Nicola Lambiase, Mirai Solutions GmbH
#
#############################################################################

test.workbook.clearRange <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookClearCells.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookClearCells.xlsx"), create = FALSE)
	
	checkDf <- data.frame(
			"one" = 1:5,
			"two" = c(NA, NA, 8, 9, 10),
			"three" = c(NA, NA, 13, 14, 15),
			"four" = 16:20,
			"five" = c(21, 22, NA, NA, 25),
			"six" = c(26, 27, NA, NA, 30),
			"seven" = 31:35,
			stringsAsFactors = F
	)
	
	# Check that clearing ranges returns the desired result (*.xls)
	clearRange(wb.xls, "clearRange", c(c(4, 4, 5, 5), c(6, 7, 7, 8)))
	res <- readWorksheet(wb.xls, "clearRange", region = "C3:I8", header = TRUE)
	checkEquals(res, checkDf)
	
	# Check that clearing ranges returns the desired result (*.xlsx)
	clearRange(wb.xlsx, "clearRange", c(c(4, 4, 5, 5), c(6, 7, 7, 8)))
	res <- readWorksheet(wb.xlsx, "clearRange", region = "C3:I8", header = TRUE)
	checkEquals(res, checkDf)
}