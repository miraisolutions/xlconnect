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
# Tests around querying Excel reference formulas
# 
# Author: Martin Studer, Mirai Solutions GmbH
#
#############################################################################

test.workbook.getReferenceFormula <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookReferenceFormula.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookReferenceFormula.xlsx"), create = FALSE)
	
	# Check if reference formulas match (*.xls)
	checkTrue(getReferenceFormula(wb.xls, "FirstName") == "Tabelle1!$A$1")
	checkTrue(substring(getReferenceFormula(wb.xls, "SecondName"), 1, 5) == "#REF!")
	
	# Check if reference formulas match (*.xlsx)
	checkTrue(getReferenceFormula(wb.xlsx, "FirstName") == "Tabelle1!$A$1")
	checkTrue(substring(getReferenceFormula(wb.xlsx, "SecondName"), 1, 5) == "#REF!")
	
}
