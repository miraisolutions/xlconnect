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
# Tests around checking existence of Excel worksheets
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.existsSheet <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookExistsNameAndSheet.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookExistsNameAndSheet.xlsx"), create = FALSE)
	
	# Check that the following sheets exists (*.xls)
	checkTrue(existsSheet(wb.xls, "AAA"))
	checkTrue(existsSheet(wb.xls, "BBB"))
	checkTrue(existsSheet(wb.xls, "CCC"))
	
	# Check that the following do NOT exists (*.xls)
	checkTrue(!existsSheet(wb.xls, "DDD"))
	checkTrue(!existsSheet(wb.xls, "'illegal name"))
	checkTrue(!existsSheet(wb.xls, "%&$$-^~@afk20 235-??a?"))
	
	# Check that the following names exists (*.xlsx)
	checkTrue(existsSheet(wb.xlsx, "AAA"))
	checkTrue(existsSheet(wb.xlsx, "BBB"))
	checkTrue(existsSheet(wb.xlsx, "CCC"))
	
	# Check that the following do NOT exists (*.xlsx)
	checkTrue(!existsSheet(wb.xlsx, "DDD"))
	checkTrue(!existsSheet(wb.xlsx, "'illegal name"))
	checkTrue(!existsSheet(wb.xlsx, "%&$$-^~@afk20 235-??a?"))
	
}
