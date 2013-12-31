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
# Tests around reading Excel tables from an Excel workbook
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.readTable <- function() {
	
	# Create workbooks
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookReadTable.xlsx"), create = FALSE)
	
	checkDf <- data.frame(
	  "NumericColumn" = c(-23.63, NA, NA, 5.8, 3),
	  "StringColumn" = c("Hello", NA, NA, NA, "World"),
	  "BooleanColumn" = c(TRUE, FALSE, FALSE, NA, NA),
	  "DateTimeColumn" = as.POSIXct(c(NA, NA, "2010-09-09 21:03:07", "2010-09-10 21:03:07", "2010-09-11 21:03:07")),
	  stringsAsFactors = F
	)
  
	# Check that reading an Excel table works as expected (*.xlsx)
	res <- readTable(wb.xlsx, sheet = "Test", table = "TestTable1")
	checkEquals(res, checkDf)
}
