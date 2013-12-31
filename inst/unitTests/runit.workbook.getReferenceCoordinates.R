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
# Tests around querying Excel reference coordinates
# 
# Author: Thomas Themel, Mirai Solutions GmbH
#
#############################################################################

test.workbook.getReferenceCoordinates <- function() {
	
	# Create workbooks
	wb.xls <- loadWorkbook(rsrc("resources/testWorkbookReferenceFormula.xls"), create = FALSE)
	wb.xlsx <- loadWorkbook(rsrc("resources/testWorkbookReferenceFormula.xlsx"), create = FALSE)
	
	# Check if reference formulas match (*.xls)
	checkTrue(all(getReferenceCoordinatesForName(wb.xls, "FirstName") == matrix(c(1,1,1,1), nrow = 2, byrow=TRUE)))
	checkException(getReferenceCoordinatesForName(wb.xls, "NonExistentName"))

	# Check if reference positions match (*.xlsx)
	checkTrue(all(getReferenceCoordinatesForName(wb.xlsx, "FirstName") == matrix(c(1,1,1,1), nrow = 2, byrow=TRUE)))
	checkException(getReferenceCoordinatesForName(wb.xlsx, "NonExistentName"))
}
