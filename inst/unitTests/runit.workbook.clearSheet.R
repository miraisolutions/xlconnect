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
# Tests around clearing sheets from an Excel Workbook
# 
# Author: Nicola Lambiase, Mirai Solutions GmbH
#
#############################################################################

test.workbook.clearSheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookClearCells.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookClearCells.xlsx"), create = FALSE)
	
	# Check that clearing sheets returns empty sheets (*.xls)
	clearSheet(wb.xls, c("clearSheet1", "clearSheet2"))
	res1 <- getLastRow(wb.xls, "clearSheet1")
	res2 <- getLastRow(wb.xls, "clearSheet2")
	checkEquals(c(res1, res2), c(clearSheet1 = 1, clearSheet2=1))

	# Check that clearing sheets returns empty sheets (*.xlsx)
	clearSheet(wb.xlsx, c("clearSheet1", "clearSheet2"))
	res1 <- getLastRow(wb.xlsx, "clearSheet1")
	res2 <- getLastRow(wb.xlsx, "clearSheet2")
	checkEquals(c(res1, res2), c(clearSheet1 = 1, clearSheet2=1))

}