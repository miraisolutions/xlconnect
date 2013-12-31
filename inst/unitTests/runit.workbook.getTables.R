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
# Tests around querying Excel tables on a worksheet
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.getTables <- function() {
	
	# Create workbooks
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookReadTable.xlsx"), create = FALSE)
  
	# Check that querying Excel table works as expected (*.xlsx)
	res <- getTables(wb.xlsx, sheet = "Test", simplify = TRUE)
	checkEquals(res, "TestTable1")
  
  res <- getTables(wb.xlsx, sheet = "Test", simplify = FALSE)
  checkEquals(res, list(Test = "TestTable1"))
  
  res <- getTables(wb.xlsx, sheet = "NoTableHere", simplify = TRUE)
  checkEquals(res, character(0))
  
  # Check that trying to query tables from an non-existent sheet throws an exception (*.xlsx)
  checkException(getTables(wb.xlsx, sheet = "DoesNotExist"))
}
