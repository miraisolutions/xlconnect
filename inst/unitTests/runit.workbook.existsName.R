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
# Tests around checking existence of Excel names
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.existsName <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookExistsNameAndSheet.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookExistsNameAndSheet.xlsx"), create = FALSE)
	
	# Check that the following names exists (*.xls)
	checkTrue(existsName(wb.xls, "AA"))
	checkTrue(existsName(wb.xls, "BB"))
	checkTrue(existsName(wb.xls, "CC"))
	
	# Check that the following do NOT exists (*.xls)
	checkTrue(!existsName(wb.xls, "DD"))
	checkTrue(!existsName(wb.xls, "'illegal name"))
	checkTrue(!existsName(wb.xls, "%&$$-^~@afk20 235-??a?"))
	
	# Check that the following names exists (*.xlsx)
	checkTrue(existsName(wb.xlsx, "AA"))
	checkTrue(existsName(wb.xlsx, "BB"))
	checkTrue(existsName(wb.xlsx, "CC"))
	
	# Check that the following do NOT exists (*.xlsx)
	checkTrue(!existsName(wb.xlsx, "DD"))
	checkTrue(!existsName(wb.xlsx, "'illegal name"))
	checkTrue(!existsName(wb.xlsx, "%&$$-^~@afk20 235-??a?"))
}
