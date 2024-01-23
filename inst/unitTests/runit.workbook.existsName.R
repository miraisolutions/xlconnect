#############################################################################
#
# XLConnect
# Copyright (C) 2010-2024 Mirai Solutions GmbH
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
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
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
	checkEquals(existsName(wb.xls, "AA"), TRUE, check.attributes = FALSE)
	checkEquals(existsName(wb.xls, "BB"), TRUE, check.attributes = FALSE)
	checkEquals(existsName(wb.xls, "CC"), TRUE, check.attributes = FALSE)
	
	# Check that the following do NOT exists (*.xls)
	checkEquals(existsName(wb.xls, "DD"), FALSE, check.attributes = FALSE)
	checkEquals(existsName(wb.xls, "'illegal name"), FALSE, check.attributes = FALSE)
	checkEquals(existsName(wb.xls, "%&$$-^~@afk20 235-??a?"), FALSE, check.attributes = FALSE)
	
	# Check that the following names exists (*.xlsx)
	checkEquals(existsName(wb.xlsx, "AA"), TRUE, check.attributes = FALSE)
	checkEquals(existsName(wb.xlsx, "BB"), TRUE, check.attributes = FALSE)
	checkEquals(existsName(wb.xlsx, "CC"), TRUE, check.attributes = FALSE)
	
	# Check that the following do NOT exists (*.xlsx)
	checkEquals(existsName(wb.xlsx, "DD"), FALSE, check.attributes = FALSE)
	checkEquals(existsName(wb.xlsx, "'illegal name"), FALSE, check.attributes = FALSE)
	checkEquals(existsName(wb.xlsx, "%&$$-^~@afk20 235-??a?"), FALSE, check.attributes = FALSE)

	# TODO check with attributes - where was the name found ?
}
